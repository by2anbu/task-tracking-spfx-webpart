
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import {
    DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, TextField, Stack,
    PrimaryButton, Dialog, DialogType, DialogFooter, DefaultButton, ConstrainMode,
    IGroup, IDetailsGroupDividerProps, Icon, FontIcon
} from 'office-ui-fabric-react';
import { taskService } from '../../../../services/sp-service';
import { IMainTask, ISubTask } from '../../../../services/interfaces';

export interface ICorrespondenceViewProps {
    userEmail: string;
    isAdmin: boolean;
}

export interface ICorrespondenceViewState {
    items: any[];
    filteredItems: any[];
    loading: boolean;
    groups: IGroup[];

    // Filters
    filterParentTitle: string;
    filterChildTitle: string;
    filterFrom: string;
    filterTo: string;

    // Lookups
    mainTaskMap: { [key: number]: IMainTask };
    subTaskMap: { [key: number]: ISubTask };

    // Dialog
    showDialog: boolean;
    selectedItem: any | undefined;
    attachments: any[];
    readIds: number[]; // Track dismissed notifications
}

export class CorrespondenceView extends React.Component<ICorrespondenceViewProps, ICorrespondenceViewState> {
    constructor(props: ICorrespondenceViewProps) {
        super(props);
        this.state = {
            items: [],
            filteredItems: [],
            loading: true,
            groups: [],
            filterParentTitle: '',
            filterChildTitle: '',
            filterFrom: '',
            filterTo: '',
            mainTaskMap: {},
            subTaskMap: {},
            showDialog: false,
            selectedItem: undefined,
            attachments: [],
            readIds: []
        };
    }

    public async componentDidMount(): Promise<void> {
        // Load Read Status from LocalStorage
        try {
            const stored = localStorage.getItem('task_tracking_dismissed_notifications');
            if (stored) {
                const ids = JSON.parse(stored);
                this.setState({ readIds: ids });
            }
        } catch (e) {
            console.error("Error loading read status", e);
        }
        await this.loadData();
    }

    private loadData = async (): Promise<void> => {
        this.setState({ loading: true });
        try {
            // Permissions Logic
            let relevantEmails: any[] = [];

            if (this.props.isAdmin) {
                relevantEmails = await taskService.getAllCorrespondence();
            } else {
                const userEmail = (this.props.userEmail || '').toLowerCase();
                // Security Fix: Use server-side filtering
                relevantEmails = await taskService.getPermissionsAwareCorrespondence(userEmail);
            }

            const mainTasks = await taskService.getAllMainTasks();
            const subTasks = await taskService.getAllSubTasks();

            const mainMap: { [key: number]: IMainTask } = {};
            mainTasks.forEach((t: IMainTask) => {
                mainMap[t.Id] = t;
            });

            const subMap: { [key: number]: ISubTask } = {};
            subTasks.forEach((t: ISubTask) => {
                subMap[t.Id] = t;
            });

            this.setState({
                items: relevantEmails,
                filteredItems: relevantEmails,
                mainTaskMap: mainMap,
                subTaskMap: subMap,
                loading: false
            }, this.applyFilters);

        } catch (error) {
            console.error(error);
            this.setState({ loading: false });
        }
    }

    private applyFilters = (): void => {
        const { items, filterParentTitle, filterChildTitle, filterFrom, filterTo, mainTaskMap, subTaskMap } = this.state;

        const filtered = items.filter(item => {
            const parent = mainTaskMap[item.ParentTaskID];
            const parentTitle = parent ? parent.Title : '';
            if (filterParentTitle && parentTitle.toLowerCase().indexOf(filterParentTitle.toLowerCase()) === -1) return false;

            const child = item.ChildTaskID ? subTaskMap[item.ChildTaskID] : null;
            const childTitle = child ? child.Task_Title : '';
            if (filterChildTitle) {
                if (childTitle.toLowerCase().indexOf(filterChildTitle.toLowerCase()) === -1) return false;
            }

            const fromVal = (item.FromAddress || item.Author?.Title || '').toLowerCase();
            if (filterFrom && fromVal.indexOf(filterFrom.toLowerCase()) === -1) return false;

            const toVal = (item.ToAddress || '').toLowerCase();
            if (filterTo && toVal.indexOf(filterTo.toLowerCase()) === -1) return false;

            return true;
        });

        const { sortedItems, groups } = this.groupItems(filtered);
        this.setState({ filteredItems: sortedItems, groups: groups });
    }

    // Grouping Logic: Main Task -> Sub Task -> Correspondence
    private groupItems = (items: any[]): { sortedItems: any[], groups: IGroup[] } => {
        const { mainTaskMap, subTaskMap } = this.state;
        const groups: IGroup[] = [];
        let sortedItems: any[] = [];

        // 1. Group by Parent Task
        const parentGroups: { [key: number]: any[] } = {};
        items.forEach(item => {
            const pid = item.ParentTaskID || 0;
            if (!parentGroups[pid]) parentGroups[pid] = [];
            parentGroups[pid].push(item);
        });

        // Sort Parent Tasks by ID descending (newest first)
        const parentIds = Object.keys(parentGroups).map(Number).sort((a, b) => b - a);

        let globalIndex = 0;

        parentIds.forEach(pid => {
            const parentItems = parentGroups[pid];
            const mainTask = mainTaskMap[pid];
            const parentTitle = mainTask ? mainTask.Title : `Unknown Parent (${pid})`;

            // Create Level 1 Group (Main Task)
            const mainGroup: IGroup = {
                key: `parent_${pid}`,
                name: parentTitle,
                startIndex: globalIndex,
                count: parentItems.length,
                level: 0,
                isCollapsed: false,
                data: { type: 'main', task: mainTask }
            };

            // 2. Group by Child Task within Parent
            const childGroups: { [key: string]: any[] } = {}; // Use string key to handle '0' (no child)
            parentItems.forEach(item => {
                const cid = item.ChildTaskID || 0;
                if (!childGroups[cid]) childGroups[cid] = [];
                childGroups[cid].push(item);
            });

            // Sort Child IDs
            const childIds = Object.keys(childGroups).map(Number).sort((a, b) => b - a);

            // If we have sub-groups (Subtasks)
            const subGroups: IGroup[] = [];

            childIds.forEach(cid => {
                const childItems = childGroups[cid];
                // Sort items by Date Descending
                childItems.sort((a, b) => new Date(b.Created).getTime() - new Date(a.Created).getTime());

                // Add to flat list
                sortedItems = sortedItems.concat(childItems);

                const subTask = subTaskMap[cid];
                const childTitle = subTask ? subTask.Task_Title : (cid === 0 ? 'General (No Subtask)' : `Unknown Subtask (${cid})`);

                const subGroup: IGroup = {
                    key: `parent_${pid}_child_${cid}`,
                    name: childTitle,
                    startIndex: globalIndex,
                    count: childItems.length,
                    level: 1,
                    isCollapsed: false,
                    data: { type: 'sub', task: subTask, count: childItems.length }
                };
                subGroups.push(subGroup);
                globalIndex += childItems.length;
            });

            // If there's only one child group and it's 0 (General), maybe we don't need a sub-group?
            // Requirement says "Level 1 - Main, Level 2 - Sub". Let's keep it consistent.

            mainGroup.children = subGroups;
            groups.push(mainGroup);
        });

        return { sortedItems, groups };
    }

    private _onFilterChange = (key: keyof ICorrespondenceViewState, val: string): void => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        this.setState({ [key]: val } as any, this.applyFilters);
    }

    private _onClearFilters = (): void => {
        this.setState({
            filterParentTitle: '',
            filterChildTitle: '',
            filterFrom: '',
            filterTo: ''
        }, this.applyFilters);
    }

    private _onItemInvoked = async (item: any): Promise<void> => {
        this.setState({ showDialog: true, selectedItem: item, attachments: [] });
        if (item.Id) {
            try {
                const atts = await taskService.getAttachments('Task Correspondence', item.Id);
                this.setState({ attachments: atts });
            } catch (e) { console.error('Error fetching attachments', e); }
        }
    }

    private _closeDialog = (): void => {
        this.setState({ showDialog: false, selectedItem: undefined, attachments: [] });
    }

    private _onExport = (): void => {
        const { filteredItems, mainTaskMap, subTaskMap } = this.state;
        if (!filteredItems || filteredItems.length === 0) return;

        const headers = ["ID", "Date", "Parent Task ID", "Parent Task Title", "Child Task ID", "Child Task Title", "Subject", "From", "To", "Message Body"];

        const cvsRows = filteredItems.map(item => {
            const date = new Date(item.Created).toLocaleDateString() + ' ' + new Date(item.Created).toLocaleTimeString();
            const parent = mainTaskMap[item.ParentTaskID];
            const child = item.ChildTaskID ? subTaskMap[item.ChildTaskID] : null;
            const parentTitle = parent ? parent.Title : '';
            const childTitle = child ? child.Task_Title : '';
            const cleanBody = (item.MessageBody || '').replace(/<[^>]+>/g, '').replace(/[\n\r]+/g, ' ').replace(/"/g, '""');

            return [
                item.Id,
                date,
                item.ParentTaskID,
                `"${parentTitle}"`,
                item.ChildTaskID || '',
                `"${childTitle}"`,
                `"${(item.Title || '').replace(/"/g, '""')}"`,
                `"${(item.FromAddress || item.Author?.Title || '').replace(/"/g, '""')}"`,
                `"${(item.ToAddress || '').replace(/"/g, '""')}"`,
                `"${cleanBody}"`
            ].join(',');
        });

        const csvContent = [headers.join(',')].concat(cvsRows).join('\n');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const nav = (window.navigator as any);
        if (nav.msSaveOrOpenBlob) {
            nav.msSaveOrOpenBlob(new Blob([csvContent]), `Correspondence_Export.csv`);
        } else {
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.setAttribute("href", url);
            link.setAttribute("download", `Correspondence_Export_${new Date().toISOString().slice(0, 10)}.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }

    // Custom Group Header Renderer
    private _onRenderGroupHeader = (props: IDetailsGroupDividerProps | undefined): JSX.Element | null => {
        if (!props || !props.group) return null;
        const group = props.group;
        const data = group.data;

        // Status Colors
        const getStatusColor = (status: string | undefined) => {
            if (!status) return '#605e5c'; // Gray
            if (status === 'Completed') return '#107c10'; // Green
            if (status === 'In Progress') return '#ffb900'; // Yellow
            return '#605e5c'; // Gray (Not Started)
        };

        // Format date helper
        const formatDate = (dateStr: string | undefined): string => {
            if (!dateStr) return '-';
            const d = new Date(dateStr);
            return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        };

        // Truncate text
        const truncate = (text: string | undefined, maxLen: number): string => {
            if (!text) return '-';
            return text.length > maxLen ? text.substring(0, maxLen) + '...' : text;
        };

        if (group.level === 0) {
            // Level 1: Main Task - Detailed Header
            const task = data.task as IMainTask;
            if (!task) return <div style={{ padding: 10, fontWeight: 600 }}>{group.name}</div>;

            const color = getStatusColor(task.Status);
            const percentComplete = task.PercentComplete ? (task.PercentComplete * 100).toFixed(0) : '0';

            return (
                <div style={{
                    display: 'grid',
                    gridTemplateColumns: '30px 250px 200px 100px 100px 100px 100px 100px',
                    gap: '10px',
                    alignItems: 'center',
                    padding: '12px 10px',
                    backgroundColor: '#f3f2f1',
                    borderBottom: '2px solid #0078d4',
                    cursor: 'pointer',
                    fontSize: 13
                }} onClick={() => props.onToggleCollapse!(group)}>
                    <Icon iconName={group.isCollapsed ? 'ChevronRight' : 'ChevronDown'} style={{ fontSize: 14 }} />

                    <div style={{ display: 'flex', alignItems: 'center' }}>
                        <div style={{ width: 12, height: 12, borderRadius: '50%', backgroundColor: color, marginRight: 8 }} title={task.Status || 'Not Started'} />
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#005a9e' }}>{truncate(task.Title, 30)}</div>
                    </div>

                    <div style={{ fontSize: 12, color: '#666' }} title={task.Task_x0020_Description}>
                        {truncate(task.Task_x0020_Description, 25)}
                    </div>

                    <div>
                        <span style={{
                            padding: '3px 8px',
                            backgroundColor: color,
                            color: 'white',
                            borderRadius: 3,
                            fontSize: 11,
                            fontWeight: 600
                        }}>
                            {task.Status || 'Not Started'}
                        </span>
                    </div>

                    <div style={{ fontWeight: 600 }}>{percentComplete}%</div>

                    <div>{formatDate(task.TaskStartDate)}</div>

                    <div style={{ fontWeight: task.TaskDueDate && new Date(task.TaskDueDate) < new Date() && task.Status !== 'Completed' ? 700 : 400, color: task.TaskDueDate && new Date(task.TaskDueDate) < new Date() && task.Status !== 'Completed' ? '#a80000' : 'inherit' }}>
                        {formatDate(task.TaskDueDate)}
                    </div>

                    <div style={{ color: task.Task_x0020_End_x0020_Date ? '#107c10' : '#999', fontWeight: task.Task_x0020_End_x0020_Date ? 600 : 400 }}>
                        {formatDate(task.Task_x0020_End_x0020_Date)}
                    </div>
                </div>
            );
        } else {
            // Level 2: Sub Task - Detailed Header
            const task = data.task as ISubTask;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const assignee = task && task.TaskAssignedTo ? (Array.isArray(task.TaskAssignedTo) ? task.TaskAssignedTo.map(u => u.Title).join(', ') : (task.TaskAssignedTo as any).Title) : 'Unassigned';

            const color = getStatusColor(task ? task.TaskStatus : 'Not Started');
            const createdDate = task ? task.Task_Created_Date || '' : '';

            return (
                <div style={{
                    display: 'grid',
                    gridTemplateColumns: '50px 200px 150px 120px 100px 90px 100px 100px 100px 100px',
                    gap: '8px',
                    alignItems: 'center',
                    padding: '10px 10px 10px 40px',
                    backgroundColor: '#fafafa',
                    borderBottom: '1px solid #edebe9',
                    cursor: 'pointer',
                    fontSize: 12
                }} onClick={() => props.onToggleCollapse!(group)}>
                    <Icon iconName={group.isCollapsed ? 'ChevronRight' : 'ChevronDown'} style={{ fontSize: 12 }} />

                    <div style={{ display: 'flex', alignItems: 'center' }}>
                        <div style={{ width: 10, height: 10, borderRadius: '50%', backgroundColor: color, marginRight: 6 }} title={task ? task.TaskStatus : 'N/A'} />
                        <div style={{ fontWeight: 600, fontSize: 13 }}>{truncate(group.name, 20)}</div>
                    </div>

                    <div style={{ fontSize: 11, color: '#666' }} title={task ? task.Task_Description : ''}>
                        {truncate(task ? task.Task_Description : '', 18)}
                    </div>

                    <div style={{ fontSize: 11 }}>{truncate(assignee, 15)}</div>

                    <div>{formatDate(createdDate)}</div>

                    <div style={{ fontSize: 11 }}>{task ? task.Category || '-' : '-'}</div>

                    <div>
                        <span style={{
                            padding: '2px 6px',
                            backgroundColor: color,
                            color: 'white',
                            borderRadius: 3,
                            fontSize: 10,
                            fontWeight: 600
                        }}>
                            {task ? task.TaskStatus || 'Not Started' : 'N/A'}
                        </span>
                    </div>

                    <div>{formatDate(createdDate)}</div>

                    <div style={{ fontWeight: task && task.TaskDueDate && new Date(task.TaskDueDate) < new Date() && task.TaskStatus !== 'Completed' ? 700 : 400, color: task && task.TaskDueDate && new Date(task.TaskDueDate) < new Date() && task.TaskStatus !== 'Completed' ? '#a80000' : 'inherit' }}>
                        {formatDate(task ? task.TaskDueDate : undefined)}
                    </div>

                    <div style={{ color: task && task.Task_End_Date ? '#107c10' : '#999' }}>
                        {formatDate(task ? task.Task_End_Date : undefined)}
                    </div>

                    <div style={{ display: 'flex', alignItems: 'center', backgroundColor: '#e1dfdd', padding: '2px 6px', borderRadius: 8, justifySelf: 'end' }}>
                        <span style={{ marginRight: 4, fontSize: 10 }}>📩</span>
                        <strong style={{ fontSize: 11 }}>{data.count}</strong>
                    </div>
                </div>
            );
        }
    }

    public render(): React.ReactElement<ICorrespondenceViewProps> {
        const { filteredItems, loading, filterParentTitle, filterChildTitle, filterFrom, filterTo, groups, showDialog, selectedItem, attachments } = this.state;

        const columns: IColumn[] = [
            {
                key: 'Attachment', name: '', minWidth: 24, maxWidth: 24, onRender: (item) => {
                    const isUnread = this.state.readIds.indexOf(item.Id) === -1;
                    const content = item.Attachments || (item.AttachmentFiles && item.AttachmentFiles.length > 0) ?
                        <FontIcon iconName="Attach" style={{ fontSize: 14, color: '#0078d4' }} /> : null;

                    return (
                        <div style={{ display: 'flex', alignItems: 'center' }}>
                            {isUnread && <div style={{ width: 6, height: 6, borderRadius: '50%', backgroundColor: '#0078d4', marginRight: 4 }} title="New Notification" />}
                            {content}
                        </div>
                    );
                }
            },
            {
                key: 'Date', name: 'Date', minWidth: 120, maxWidth: 140,
                onRender: (item) => {
                    const isUnread = this.state.readIds.indexOf(item.Id) === -1;
                    return <span style={{ fontWeight: isUnread ? 700 : 400, color: isUnread ? '#0078d4' : 'inherit' }}>
                        {new Date(item.Created).toLocaleDateString() + ' ' + new Date(item.Created).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                    </span>;
                }
            },
            {
                key: 'Subject', name: 'Subject', minWidth: 200,
                onRender: (item) => {
                    const isUnread = this.state.readIds.indexOf(item.Id) === -1;
                    return <span style={{ fontWeight: isUnread ? 700 : 600 }}>{item.Title}</span>
                }
            },
            {
                key: 'From', name: 'From', minWidth: 150, maxWidth: 200,
                onRender: (item) => item.FromAddress || item.Author?.Title
            },
            {
                key: 'To', name: 'To', minWidth: 150, maxWidth: 200,
                onRender: (item) => item.ToAddress
            },
            {
                key: 'Body', name: 'Message Preview', minWidth: 300,
                onRender: (item) => {
                    const raw = item.MessageBody || '';
                    const plain = raw.replace(/<[^>]+>/g, '');
                    const isUnread = this.state.readIds.indexOf(item.Id) === -1;
                    return <span style={{ color: isUnread ? '#323130' : '#666', fontWeight: isUnread ? 600 : 400 }}>{plain.substring(0, 60)}...</span>;
                }
            }
        ];

        if (loading) return <div>Loading correspondence report...</div>;

        return (
            <div style={{ padding: 15, backgroundColor: '#fff', minHeight: '80vh' }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 15 }}>
                    <h2 style={{ margin: 0, color: '#323130' }}>Correspondence Report</h2>
                </Stack>

                {/* Professional Filter Container */}
                <div style={{ backgroundColor: '#f8f8f8', padding: 20, borderRadius: 4, border: '1px solid #eaeaea', marginBottom: 20 }}>
                    <Stack tokens={{ childrenGap: 15 }}>
                        {/* Row 1: Context Filters */}
                        <Stack horizontal tokens={{ childrenGap: 20 }}>
                            <TextField
                                label="Main Task"
                                value={filterParentTitle}
                                onChange={(e, v) => this._onFilterChange('filterParentTitle', v || '')}
                                styles={{ root: { width: '50%' } }}
                                placeholder="Filter by Main Task Title"
                                iconProps={{ iconName: 'TaskGroup' }}
                            />
                            <TextField
                                label="Sub Task"
                                value={filterChildTitle}
                                onChange={(e, v) => this._onFilterChange('filterChildTitle', v || '')}
                                styles={{ root: { width: '50%' } }}
                                placeholder="Filter by Sub Task Title"
                                iconProps={{ iconName: 'TaskManager' }}
                            />
                        </Stack>

                        {/* Row 2: User Filters & Actions */}
                        <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end">
                            <TextField
                                label="From"
                                value={filterFrom}
                                onChange={(e, v) => this._onFilterChange('filterFrom', v || '')}
                                styles={{ root: { width: 200 } }}
                                placeholder="Name or Email"
                            />
                            <TextField
                                label="To"
                                value={filterTo}
                                onChange={(e, v) => this._onFilterChange('filterTo', v || '')}
                                styles={{ root: { width: 200 } }}
                                placeholder="Email"
                            />

                            <Stack.Item grow={1}>
                                <span />
                            </Stack.Item>

                            <DefaultButton
                                text="Clear Filters"
                                iconProps={{ iconName: 'Clear' }}
                                onClick={this._onClearFilters}
                            />
                            <PrimaryButton
                                text="Export to Excel"
                                iconProps={{ iconName: 'ExcelDocument' }}
                                onClick={this._onExport}
                            />
                        </Stack>
                    </Stack>
                </div>

                {/* Column Headers for Hierarchical View */}
                <div style={{ backgroundColor: '#0078d4', color: 'white', padding: '8px 10px', fontSize: 11, fontWeight: 600, marginBottom: 0 }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '30px 250px 200px 100px 100px 100px 100px 100px', gap: '10px' }}>
                        <div />
                        <div>Task / Subtask</div>
                        <div>Description</div>
                        <div>Status</div>
                        <div>% Complete</div>
                        <div>Start Date</div>
                        <div>Due Date</div>
                        <div>End Date</div>
                    </div>
                </div>

                <div style={{ fontStyle: 'italic', marginBottom: 0, padding: '8px 10px', fontSize: 11, color: '#605e5c', backgroundColor: '#f8f8f8', borderBottom: '1px solid #edebe9' }}>
                    * Hierarchy: Main Task &gt; Sub Task &gt; Correspondence. Click headers to expand/collapse groups.
                </div>

                <div style={{ height: '600px', overflow: 'auto', border: '1px solid #edebe9', borderTop: 'none', position: 'relative' }}>
                    <DetailsList
                        items={filteredItems}
                        groups={groups}
                        columns={columns}
                        selectionMode={SelectionMode.none}
                        layoutMode={DetailsListLayoutMode.fixedColumns} // Fixed props error
                        constrainMode={ConstrainMode.unconstrained}
                        groupProps={{
                            onRenderHeader: this._onRenderGroupHeader
                        }}
                        onItemInvoked={this._onItemInvoked}
                        compact={true}
                    />
                </div>

                {/* Dialog for Details */}
                <Dialog
                    hidden={!showDialog}
                    onDismiss={this._closeDialog}
                    dialogContentProps={{
                        type: DialogType.largeHeader,
                        title: selectedItem ? selectedItem.Title : 'Message Details',
                        subText: selectedItem ? `Date: ${new Date(selectedItem.Created).toLocaleString()}` : ''
                    }}
                    minWidth={600}
                >
                    {selectedItem && (
                        <div>
                            <p><strong>From:</strong> {selectedItem.FromAddress || selectedItem.Author?.Title} <br />
                                <strong>To:</strong> {selectedItem.ToAddress}</p>
                            <hr />
                            <div className="email-body-content"
                                style={{ maxHeight: 300, overflowY: 'auto', padding: 10, background: '#f9f9f9', border: '1px solid #eaeaea' }}
                                dangerouslySetInnerHTML={{ __html: selectedItem.MessageBody || '' }}
                            />
                            <hr />
                            <p><strong>Attachments:</strong></p>
                            {attachments.length === 0 ? <span>(None)</span> : (
                                <ul>
                                    {attachments.map((att) => (
                                        <li key={att.FileName || att.Name}>
                                            <a href={att.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">{att.FileName || att.Name}</a>
                                        </li>
                                    ))}
                                </ul>
                            )}
                        </div>
                    )}
                    <DialogFooter>
                        <DefaultButton onClick={this._closeDialog} text="Close" />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }
}
