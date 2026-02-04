
import * as React from 'react';
import { Pivot, PivotItem, DetailsList, SelectionMode, IColumn, IGroup, PrimaryButton, Panel, PanelType, TextField, DatePicker, Stack, DefaultButton, MessageBar, MessageBarType, Dropdown, IconButton, Checkbox, Persona, PersonaSize, ComboBox, IComboBoxOption, ScrollablePane, ScrollbarVisibility, Sticky, StickyPositionType, ConstrainMode, DetailsListLayoutMode, IDetailsHeaderProps, TooltipHost, Icon } from 'office-ui-fabric-react';
import { TaskDetail } from './TaskDetail';
import { AdminReports } from './AdminReports';

import styles from '../TaskTracking.module.scss';
import { taskService } from '../../../../services/sp-service';
import { IMainTask, ISubTask, LIST_MAIN_TASKS } from '../../../../services/interfaces';

export interface IAdminDashboardProps {
    initialViewTaskId?: number;
    initialParentTaskId?: number;
    initialChildTaskId?: number;
    initialTab?: string;
    onMainDeepLinkProcessed?: () => void;
    onChildDeepLinkProcessed?: () => void;
}

export const AdminDashboard: React.FunctionComponent<IAdminDashboardProps> = (props) => {
    const [tasks, setTasks] = React.useState<IMainTask[]>([]);
    const [allSubTasks, setAllSubTasks] = React.useState<ISubTask[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);
    const [error, setError] = React.useState<string | undefined>(undefined);

    const [isCreating, setIsCreating] = React.useState(false);
    const [newTitle, setNewTitle] = React.useState('');
    const [newDesc, setNewDesc] = React.useState('');
    const [newAssignKey, setNewAssignKey] = React.useState<string | number | undefined>(undefined);
    const [newDueDate, setNewDueDate] = React.useState<Date | undefined>(undefined);
    const [newYear, setNewYear] = React.useState<string>('');
    const [newMonth, setNewMonth] = React.useState<string>('');
    const [newDept, setNewDept] = React.useState('');
    const [newProject, setNewProject] = React.useState('');
    const [msg, setMsg] = React.useState<string | undefined>(undefined);
    const [attachFiles, setAttachFiles] = React.useState<File[]>([]);

    // User & Options State
    const [userOptions, setUserOptions] = React.useState<IComboBoxOption[]>([]);
    const [departmentOptions, setDepartmentOptions] = React.useState<IComboBoxOption[]>([]);

    // Filter state
    const [searchText, setSearchText] = React.useState('');
    const [statusFilter, setStatusFilter] = React.useState<string | undefined>(undefined);
    const [yearFilter, setYearFilter] = React.useState<string | undefined>(undefined);
    const [monthFilter, setMonthFilter] = React.useState<string | undefined>(undefined);
    const [userFilter, setUserFilter] = React.useState<string | undefined>(undefined);
    const [overdueFilter, setOverdueFilter] = React.useState<boolean>(false);

    // Sorting state
    const [sortedColumn, setSortedColumn] = React.useState<string | undefined>(undefined);
    const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

    const [selectedTaskForDetail, setSelectedTaskForDetail] = React.useState<IMainTask | null>(null);
    const [selectedStatusPopup, setSelectedStatusPopup] = React.useState<string | undefined>(undefined);
    const [clarificationMetadata, setClarificationMetadata] = React.useState<Map<number, { hasCorrespondence: boolean, isReply: boolean }>>(new Map());
    const prevReplyCountRef = React.useRef<number>(0);
    const processedMainIdRef = React.useRef<number | undefined>(undefined);

    React.useEffect(() => {
        loadData().catch(console.error);
    }, []);

    // Sound and Pulse trigger logic
    React.useEffect(() => {
        let replies = 0;
        clarificationMetadata.forEach(v => { if (v.isReply) replies++; });

        if (replies > prevReplyCountRef.current) {
            console.log("[Notification] New reply detected");
            // AudioService.playBell(); // Removed sound
        }
        prevReplyCountRef.current = replies;
    }, [clarificationMetadata]);

    React.useEffect(() => {
        const style = document.createElement('style');
        style.innerText = `
            @keyframes pulse {
                0% { transform: scale(1); opacity: 1; }
                50% { transform: scale(1.2); opacity: 0.8; }
                100% { transform: scale(1); opacity: 1; }
            }
        `;
        document.head.appendChild(style);
        return () => { document.head.removeChild(style); };
    }, []);

    // Handle deep linking for ViewTaskID / ParentTaskID
    React.useEffect(() => {
        const { initialViewTaskId, initialParentTaskId } = props;
        const targetId = initialViewTaskId || initialParentTaskId;

        if (loading || tasks.length === 0 || !targetId) return;

        // Prevent redundant processing
        if (processedMainIdRef.current === targetId) return;

        console.log('[AdminDashboard] Processing deep link for ID:', targetId);

        // Check if task exists in current list
        const matchingTasks = tasks.filter((t: IMainTask) => t.Id === targetId);
        const task = matchingTasks.length > 0 ? matchingTasks[0] : null;

        if (task) {
            console.log('[AdminDashboard] Task found in list, opening detail panel');
            setSelectedTaskForDetail(task);
            processedMainIdRef.current = targetId;
            if (props.onMainDeepLinkProcessed) props.onMainDeepLinkProcessed();
        } else {
            console.log('[AdminDashboard] Task not in list, fetching specifically');
            taskService.getMainTasksByIds([targetId])
                .then(fetchedTasks => {
                    if (fetchedTasks.length > 0) {
                        console.log('[AdminDashboard] Task fetched successfully, opening detail panel');
                        setSelectedTaskForDetail(fetchedTasks[0]);
                        processedMainIdRef.current = targetId;
                        if (props.onMainDeepLinkProcessed) props.onMainDeepLinkProcessed();
                    } else {
                        console.warn('[AdminDashboard] Task not found:', targetId);
                        processedMainIdRef.current = targetId;
                        if (props.onMainDeepLinkProcessed) props.onMainDeepLinkProcessed();
                    }
                })
                .catch(e => console.error('[AdminDashboard] Error fetching deep linked task:', e));
        }
    }, [loading, tasks.length, props.initialViewTaskId, props.initialParentTaskId, props.initialChildTaskId, props.initialTab]);


    const loadData = async () => {
        setLoading(true);
        setError(undefined);
        try {
            const data = await taskService.getAllMainTasks();
            setTasks(data);

            // Populate Department Options from Choice Field
            try {
                const choices = await taskService.getChoiceFieldOptions(LIST_MAIN_TASKS, 'Departments');
                setDepartmentOptions(choices.map(c => ({ key: c, text: c })));
            } catch (e) {
                console.warn('Could not load department choices', e);
            }

            // Load Users
            try {
                const users = await taskService.getSiteUsers();
                const uOptions: IComboBoxOption[] = users.map(u => ({
                    key: u.Id,
                    text: u.Title,
                    data: { email: u.Email }
                }));
                // Set both for form and filters to keep them in sync
                setUserOptions(uOptions);
            } catch (e) {
                console.warn('Could not load users', e);
            }

            // Load all subtasks for % Complete calculation
            try {
                const subtasks = await taskService.getAllSubTasks();
                setAllSubTasks(subtasks);
            } catch (e) {
                console.warn('Could not load subtasks for progress calculation', e);
            }

            // Fetch clarification metadata for Main Tasks (with rollup for subtasks)
            try {
                const ids = data.map(t => t.Id);
                const metadata = await taskService.getTaskCorrespondenceMetadata(ids, true, true);
                setClarificationMetadata(metadata);
            } catch (e) {
                console.warn('Could not load clarification metadata', e);
            }
        } catch (e: any) {
            console.error(e);
            setError("Error loading tasks: " + (e.message || e));
        } finally {
            setLoading(false);
        }
    };

    // Helper function to calculate % Complete based on subtasks
    const calculatePercentComplete = (mainTaskId: number, status: string): number => {
        const taskSubTasks = allSubTasks.filter(s => s.Admin_Job_ID === mainTaskId);

        if (status === 'Completed') return 100;

        if (taskSubTasks.length === 0) {
            // No subtasks - use status-based percentage
            if (status === 'In Progress') return 50;
            return 0;
        }

        const completedSubTasks = taskSubTasks.filter(s => s.TaskStatus === 'Completed').length;
        return Math.round((completedSubTasks / taskSubTasks.length) * 100);
    };

    // Get subtask info for display
    const getSubtaskInfo = (mainTaskId: number): { total: number; completed: number } => {
        const taskSubTasks = allSubTasks.filter(s => s.Admin_Job_ID === mainTaskId);
        const completedSubTasks = taskSubTasks.filter(s => s.TaskStatus === 'Completed').length;
        return { total: taskSubTasks.length, completed: completedSubTasks };
    };

    // Format date as dd-MMM-yyyy
    const formatDate = (date: string | Date | undefined): string => {
        if (!date) return '';
        const d = new Date(date as any);
        const dayNum = d.getDate();
        const day = dayNum < 10 ? '0' + dayNum : dayNum.toString();
        const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        const month = months[d.getMonth()];
        const year = d.getFullYear();
        return `${day}-${month}-${year}`;
    };

    // Helper to get assigned user info (handles both array and single object)
    // MUST be defined before filteredTasks which uses it
    const getAssignedUser = (task: IMainTask): { Title: string; EMail: string } => {
        const assignee = task.TaskAssignedTo;
        if (!assignee) return { Title: 'Unassigned', EMail: '' };

        // If it's an array, get first element
        if (Array.isArray(assignee) && assignee.length > 0) {
            return { Title: assignee[0].Title || 'Unassigned', EMail: (assignee[0] as any).EMail || '' };
        }

        // If it's an object (single person field)
        if (typeof assignee === 'object' && (assignee as any).Title) {
            return { Title: (assignee as any).Title, EMail: (assignee as any).EMail || '' };
        }

        return { Title: 'Unassigned', EMail: '' };
    };

    // Apply filters to tasks
    const filteredTasks = tasks.filter((t: IMainTask) => {
        // Search in Title or Description
        if (searchText) {
            const txt = searchText.toLowerCase();
            const title = (t.Title || '').toLowerCase();
            const desc = ((t as any).Task_x0020_Description || '').toLowerCase();
            if (title.indexOf(txt) === -1 && desc.indexOf(txt) === -1) return false;
        }
        // Status filter
        if (statusFilter && t.Status !== statusFilter) return false;
        // Year filter (uses SMTYear field)
        if (yearFilter) {
            const smtYear = (t as any).SMTYear || '';
            if (smtYear !== yearFilter) return false;
        }
        // Month filter (uses SMTMonth field)
        if (monthFilter) {
            const smtMonth = (t as any).SMTMonth || '';
            if (smtMonth !== monthFilter) return false;
        }
        // User filter
        if (userFilter) {
            const user = getAssignedUser(t).Title;
            if (user !== userFilter) return false;
        }
        // Overdue filter
        if (overdueFilter) {
            if (t.Status === 'Completed' || !t.TaskDueDate) return false;
            // Compare dates (ignore time for simplicity or keep it strict)
            if (new Date(t.TaskDueDate) >= new Date()) return false;
        }
        return true;
    });

    // Helper to get unique values from array
    const getUniqueValues = (arr: string[]): string[] => {
        const seen: { [key: string]: boolean } = {};
        return arr.filter((item) => {
            if (seen[item]) return false;
            seen[item] = true;
            return true;
        });
    };


    // Filter UI controls
    const statusList = getUniqueValues(tasks.map((t: IMainTask) => t.Status || '').filter((s: string) => s));
    const statusOptions = [{ key: '', text: 'All Status' }].concat(statusList.map((s: string) => ({ key: s, text: s })));

    const yearList = getUniqueValues(tasks.map((t: IMainTask) => (t as any).SMTYear || '').filter((y: string) => y));
    const yearOptions = [{ key: '', text: 'All Years' }].concat(yearList.map((y: string) => ({ key: y, text: y })));

    const monthList = getUniqueValues(tasks.map((t: IMainTask) => (t as any).SMTMonth || '').filter((m: string) => m));
    const monthOptions = [{ key: '', text: 'All Months' }].concat(monthList.map((m: string) => ({ key: m, text: m })));

    const userList = getUniqueValues(tasks.map((t: IMainTask) => getAssignedUser(t).Title));
    const userFilterOptions = [{ key: '', text: 'All Users' }].concat(userList.map((u: string) => ({ key: u, text: u })));


    // Check if any filter is active
    const hasActiveFilters = searchText || statusFilter || yearFilter || monthFilter || userFilter || overdueFilter;

    // Clear all filters
    const clearAllFilters = () => {
        setSearchText('');
        setStatusFilter(undefined);
        setYearFilter(undefined);
        setMonthFilter(undefined);
        setUserFilter(undefined);
        setOverdueFilter(false);
    };

    // Column click handler for sorting
    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        const newIsSortedDescending = column.key === sortedColumn ? !isSortedDescending : false;
        setSortedColumn(column.key);
        setIsSortedDescending(newIsSortedDescending);
    };

    // Apply sorting to filtered tasks
    const getSortedFilteredTasks = (): IMainTask[] => {
        let sorted = [...filteredTasks];
        if (sortedColumn) {
            sorted = sorted.sort((a, b) => {
                let aVal: any = '';
                let bVal: any = '';

                if (sortedColumn === 'assigned') {
                    aVal = getAssignedUser(a).Title;
                    bVal = getAssignedUser(b).Title;
                } else if (sortedColumn === 'year') {
                    aVal = (a as any).SMTYear || '';
                    bVal = (b as any).SMTYear || '';
                } else if (sortedColumn === 'month') {
                    aVal = (a as any).SMTMonth || '';
                    bVal = (b as any).SMTMonth || '';
                } else if (sortedColumn === 'desc') {
                    aVal = (a as any).Task_x0020_Description || '';
                    bVal = (b as any).Task_x0020_Description || '';
                } else {
                    aVal = (a as any)[sortedColumn] || (a as any)[sortedColumn.charAt(0).toUpperCase() + sortedColumn.slice(1)] || '';
                    bVal = (b as any)[sortedColumn] || (b as any)[sortedColumn.charAt(0).toUpperCase() + sortedColumn.slice(1)] || '';
                }

                if (aVal < bVal) return isSortedDescending ? 1 : -1;
                if (aVal > bVal) return isSortedDescending ? -1 : 1;
                return 0;
            });
        }
        return sorted;
    };

    const filterControls = (
        <div style={{ backgroundColor: '#f8f8f8', padding: '12px 16px', borderRadius: 8, marginBottom: 16, marginTop: 10, border: '1px solid #edebe9' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                <IconButton iconProps={{ iconName: 'Filter' }} title="Filters" disabled style={{ color: '#0078d4' }} />
                <span style={{ fontWeight: 600, color: '#323130' }}>Filters</span>
                {hasActiveFilters && (
                    <span style={{
                        backgroundColor: '#0078d4',
                        color: 'white',
                        fontSize: 11,
                        padding: '2px 8px',
                        borderRadius: 12,
                        marginLeft: 5
                    }}>
                        Active
                    </span>
                )}
            </div>
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap verticalAlign="end">
                <TextField
                    placeholder="Search tasks..."
                    value={searchText}
                    onChange={(_, v) => setSearchText(v || '')}
                    iconProps={{ iconName: 'Search' }}
                    styles={{
                        root: { width: 220 },
                        fieldGroup: { borderRadius: 4 }
                    }}
                />
                <Dropdown
                    placeholder="All Users"
                    selectedKey={userFilter || ''}
                    onChange={(_, opt) => setUserFilter(opt?.key as string || undefined)}
                    options={userFilterOptions}
                    styles={{
                        root: { width: 180 },
                        dropdown: { borderRadius: 4 }
                    }}
                />
                <Dropdown
                    placeholder="All Status"
                    selectedKey={statusFilter || ''}
                    onChange={(_, opt) => setStatusFilter(opt?.key as string || undefined)}
                    options={statusOptions}
                    styles={{
                        root: { width: 140 },
                        dropdown: { borderRadius: 4 }
                    }}
                />
                <Dropdown
                    placeholder="All Years"
                    selectedKey={yearFilter || ''}
                    onChange={(_, opt) => setYearFilter(opt?.key as string || undefined)}
                    options={yearOptions}
                    styles={{
                        root: { width: 110 },
                        dropdown: { borderRadius: 4 }
                    }}
                />
                <Dropdown
                    placeholder="All Months"
                    selectedKey={monthFilter || ''}
                    onChange={(_, opt) => setMonthFilter(opt?.key as string || undefined)}
                    options={monthOptions}
                    styles={{
                        root: { width: 130 },
                        dropdown: { borderRadius: 4 }
                    }}
                />
                <Checkbox
                    label="Overdue Only"
                    checked={overdueFilter}
                    onChange={(_, v) => setOverdueFilter(!!v)}
                    styles={{ root: { marginTop: 6, marginLeft: 8 } }}
                />
                {hasActiveFilters && (
                    <DefaultButton
                        iconProps={{ iconName: 'ClearFilter' }}
                        text="Clear All"
                        onClick={clearAllFilters}
                        styles={{
                            root: { borderRadius: 4, marginLeft: 8 }
                        }}
                    />
                )}
            </Stack>
        </div>
    );

    const getUserGroups = (): IGroup[] => {
        // Group by TaskAssignedTo/Title
        const groups: { [key: string]: IMainTask[] } = {};
        filteredTasks.forEach(t => {
            const assignedUser = getAssignedUser(t);
            const key = assignedUser.EMail ? `${assignedUser.Title} (${assignedUser.EMail})` : assignedUser.Title;

            if (!groups[key]) groups[key] = [];
            groups[key].push(t);
        });

        let startIndex = 0;
        return Object.keys(groups).sort().map(key => {
            const count = groups[key].length;
            const group = {
                key: key,
                name: key,
                startIndex: startIndex,
                count: count,
                level: 0
            };
            startIndex += count;
            return group;
        });
    };

    const getStatusGroups = (): IGroup[] => {
        const groups: { [key: string]: IMainTask[] } = {};
        filteredTasks.forEach(t => {
            const status = t.Status || 'Unknown';
            if (!groups[status]) groups[status] = [];
            groups[status].push(t);
        });

        let startIndex = 0;
        return Object.keys(groups).map(key => {
            const count = groups[key].length;
            const group = {
                key: key,
                name: key,
                startIndex: startIndex,
                count: count,
                level: 0
            };
            startIndex += count;
            return group;
        });
    };

    const getStatusClass = (status: string) => {
        const s = (status || '').replace(/\s+/g, '');
        if (!styles) return '';
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (styles as any)[`status_${s}`] || styles.status_NotStarted;
    };

    const columns: IColumn[] = [
        {
            key: 'action', name: '', minWidth: 40, maxWidth: 40, onRender: (item: IMainTask) => (
                <IconButton
                    iconProps={{ iconName: 'View' }}
                    title="View Details"
                    ariaLabel="View Details"
                    onClick={() => setSelectedTaskForDetail(item)}
                    styles={{ root: { height: 28 } }}
                />
            )
        },
        {
            key: 'year', name: 'Year', minWidth: 60, maxWidth: 70,
            onRender: (i) => (i as any).SMTYear || '',
            isSorted: sortedColumn === 'year', isSortedDescending, onColumnClick
        },
        {
            key: 'month', name: 'Month', minWidth: 80, maxWidth: 100,
            onRender: (i) => (i as any).SMTMonth || '',
            isSorted: sortedColumn === 'month', isSortedDescending, onColumnClick
        },
        {
            key: 'Title', name: 'Task', fieldName: 'Title', minWidth: 150, isResizable: true,
            isSorted: sortedColumn === 'Title', isSortedDescending, onColumnClick,
            onRender: (item: IMainTask) => {
                const metadata = clarificationMetadata.get(item.Id);
                return (
                    <Stack horizontal verticalAlign="center">
                        <span style={{ fontWeight: 600 }}>{item.Title}</span>
                        {metadata && (
                            <TooltipHost content={metadata.isReply ? "New reply received!" : "Correspondence history exists"}>
                                <Icon
                                    iconName={metadata.isReply ? "Ringer" : "Questionnaire"}
                                    style={{
                                        marginLeft: 8,
                                        color: metadata.isReply ? '#d13438' : '#0078d4',
                                        fontSize: 14,
                                        fontWeight: metadata.isReply ? 700 : 400,
                                        animation: metadata.isReply ? 'pulse 2s infinite' : 'none'
                                    }}
                                />
                            </TooltipHost>
                        )}
                    </Stack>
                );
            }
        },
        {
            key: 'desc', name: 'Description', minWidth: 120, isResizable: true,
            onRender: (i) => {
                const desc = (i as any).Task_x0020_Description || '';
                return <span title={desc}>{desc.length > 40 ? desc.substring(0, 40) + '...' : desc}</span>;
            },
            isSorted: sortedColumn === 'desc', isSortedDescending, onColumnClick
        },
        {
            key: 'assigned', name: 'Assigned To', minWidth: 180,
            onRender: (item: IMainTask) => {
                const user = getAssignedUser(item);
                // Use Office 365 profile picture URL
                const profilePicUrl = user.EMail
                    ? `/_layouts/15/userphoto.aspx?size=S&username=${encodeURIComponent(user.EMail)}`
                    : undefined;
                // Safely get initials
                const getInitials = (name: string): string => {
                    if (!name || name === 'Unassigned') return 'U';
                    const parts = name.trim().split(' ').filter(p => p.length > 0);
                    if (parts.length === 0) return 'U';
                    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
                    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
                };
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <Persona
                            text={user.Title}
                            size={PersonaSize.size24}
                            imageUrl={profilePicUrl}
                            imageInitials={getInitials(user.Title)}
                            styles={{
                                root: { minWidth: 'auto' },
                                details: { display: 'none' }
                            }}
                        />
                        <span style={{ fontSize: 12, fontWeight: 500 }}>{user.Title}</span>
                    </div>
                );
            },
            isSorted: sortedColumn === 'assigned', isSortedDescending, onColumnClick
        },
        {
            key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 110,
            isSorted: sortedColumn === 'Status', isSortedDescending, onColumnClick,
            onRender: (item) => (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <span className={`${styles.statusBadge} ${getStatusClass(item.Status)}`}>
                        {item.Status || 'Open'}
                    </span>
                    {item.Status === 'Completed' && (
                        <IconButton
                            iconProps={{ iconName: 'Refresh' }}
                            title="Reopen Task"
                            onClick={async (e) => {
                                e.stopPropagation();
                                try {
                                    await taskService.updateMainTaskStatus(item.Id, 'In Progress', 'Reopened from grid');
                                    await loadData();
                                } catch (err) {
                                    console.error(err);
                                }
                            }}
                            styles={{ root: { height: 16, width: 16, color: '#0078d4' } }}
                        />
                    )}
                </Stack>
            )
        },
        {
            key: 'progress', name: '% Complete', minWidth: 140,
            onRender: (item: IMainTask) => {
                const subtaskInfo = getSubtaskInfo(item.Id);
                const pct = calculatePercentComplete(item.Id, item.Status || 'Not Started');
                const progressColor = pct === 100 ? '#107c10' : pct >= 50 ? '#ffbf00' : pct > 0 ? '#00b0ff' : '#e0e0e0';

                return (
                    <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                        <div style={{ flexGrow: 1, maxWidth: 70 }}>
                            <div style={{ backgroundColor: '#e0e0e0', borderRadius: 4, height: 8, overflow: 'hidden' }}>
                                <div style={{ width: `${pct}%`, backgroundColor: progressColor, height: '100%', transition: 'width 0.3s' }} />
                            </div>
                        </div>
                        <span style={{ marginLeft: 6, fontSize: 11, fontWeight: 600, minWidth: 30 }}>{pct}%</span>
                        {subtaskInfo.total > 0 && (
                            <span style={{ marginLeft: 2, fontSize: 9, color: '#666' }}>({subtaskInfo.completed}/{subtaskInfo.total})</span>
                        )}
                    </div>
                );
            },
            isSorted: sortedColumn === 'progress', isSortedDescending, onColumnClick
        },
        {
            key: 'TaskDueDate', name: 'Due Date', minWidth: 100,
            onRender: (i: IMainTask) => {
                const isTaskOverdue = i.Status !== 'Completed' && i.TaskDueDate && new Date(i.TaskDueDate) < new Date();
                return (
                    <span style={{ color: isTaskOverdue ? '#a80000' : 'inherit', fontWeight: isTaskOverdue ? 600 : 'normal' }}>
                        {formatDate(i.TaskDueDate)}
                    </span>
                );
            },
            isSorted: sortedColumn === 'TaskDueDate', isSortedDescending, onColumnClick
        },
        {
            key: 'Task_x0020_End_x0020_Date', name: 'End Date', minWidth: 100,
            onRender: (i: IMainTask) => {
                const endDate = (i as any).Task_x0020_End_x0020_Date;
                if (endDate) {
                    return <span style={{ color: '#107c10', fontWeight: 500 }}>{formatDate(endDate)}</span>;
                }
                return <span style={{ color: '#999', fontStyle: 'italic' }}>-</span>;
            },
            isSorted: sortedColumn === 'Task_x0020_End_x0020_Date', isSortedDescending, onColumnClick
        },
    ];

    // Get sorted filtered tasks for grids
    const sortedFilteredTasks = getSortedFilteredTasks();

    // Sort by User for grouped view (maintaining additional column sort within groups)
    const tasksByUser = [...sortedFilteredTasks].sort((a, b) => {
        const userA = getAssignedUser(a);
        const keyA = userA.EMail ? `${userA.Title} (${userA.EMail})` : userA.Title;

        const userB = getAssignedUser(b);
        const keyB = userB.EMail ? `${userB.Title} (${userB.EMail})` : userB.Title;

        return keyA.localeCompare(keyB);
    });

    // Sort by Status for grouped view
    const tasksByStatus = [...sortedFilteredTasks].sort((a, b) => (a.Status || '').localeCompare(b.Status || ''));

    const handleCreateMainTask = async () => {
        try {
            if (!newTitle) { setMsg('Title is required'); return; }
            if (!newDesc) { setMsg('Description is required'); return; }
            if (!newYear) { setMsg('Year is required'); return; }
            if (!newMonth) { setMsg('Month is required'); return; }
            if (!newDueDate) { setMsg('Due Date is required'); return; }

            await taskService.createMainTask({
                Title: newTitle,
                Task_x0020_Description: newDesc,
                SMTYear: newYear,
                SMTMonth: newMonth,
                Departments: newDept,
                Project: newProject,
                TaskDueDate: newDueDate ? newDueDate.toISOString() : undefined,
                Status: 'Not Started',
                TaskAssignedToId: newAssignKey ? Number(newAssignKey) : undefined
            } as any, attachFiles);

            setIsCreating(false);
            setNewTitle(''); setNewDesc(''); setNewAssignKey(undefined); setAttachFiles([]); setNewYear(''); setNewMonth(''); setNewDept(''); setNewProject('');
            loadData().catch(console.error); // reload
        } catch (e: any) {
            setMsg(e.message || e);
        }
    };

    // Export to Excel function
    const exportToExcel = () => {
        // Build CSV content
        const headers = ['#', 'Task', 'Task Assigned', 'Email', 'Status', 'SMT Year', 'SMT Month', 'Due Date'];
        const rows = filteredTasks.map((t: IMainTask) => {
            const user = getAssignedUser(t);
            return [
                t.Id?.toString() || '',
                t.Title || '',
                user.Title,
                user.EMail,
                t.Status || '',
                (t as any).SMTYear || '',
                (t as any).SMTMonth || '',
                t.TaskDueDate ? new Date(t.TaskDueDate).toLocaleDateString() : ''
            ].map(v => `"${(v || '').toString().replace(/"/g, '""')}"`).join(',');
        });
        const csvContent = [headers.join(',')].concat(rows).join('\r\n');

        // Create and download file
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', 'TaskReport.csv');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    if (loading) {
        return <div style={{ padding: 20, textAlign: 'center' }}>Loading Admin Dashboard...</div>;
    }

    if (error) {
        return (
            <div style={{ padding: 20, color: '#a80000' }}>
                <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
            </div>
        );
    }

    return (
        <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                <h2 style={{ margin: 0 }}>Admin View</h2>
                <PrimaryButton text="Create New Main Task" iconProps={{ iconName: 'Add' }} onClick={() => setIsCreating(true)} />
            </div>

            {filterControls}

            <div style={{ marginBottom: 20 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                    <span><strong>Total Tasks:</strong> {filteredTasks.length}</span>
                    <DefaultButton text="Export to Excel" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportToExcel} />
                </div>
                <div style={{ display: 'flex', gap: 15, flexWrap: 'wrap' }}>
                    {statusList.map((status: string) => {
                        const count = filteredTasks.filter((t: IMainTask) => t.Status === status).length;
                        return (
                            <span
                                key={status}
                                className={`${styles.statusBadge} ${getStatusClass(status)}`}
                                style={{ padding: '4px 12px', cursor: 'pointer' }}
                                onClick={() => setSelectedStatusPopup(status)}
                            >
                                {status}: {count}
                            </span>
                        );
                    })}

                    {/* Overdue Badge */}
                    {(() => {
                        const overdueCount = filteredTasks.filter((t: IMainTask) => {
                            if (t.Status === 'Completed' || !t.TaskDueDate) return false;
                            return new Date(t.TaskDueDate) < new Date();
                        }).length;

                        if (overdueCount > 0) {
                            return (
                                <span
                                    className={styles.statusBadge}
                                    style={{ padding: '4px 12px', cursor: 'pointer', backgroundColor: '#a80000', color: 'white' }}
                                    onClick={() => setOverdueFilter(!overdueFilter)}
                                    title="Toggle Overdue Filter"
                                >
                                    Overdue: {overdueCount}
                                </span>
                            );
                        }
                        return null;
                    })()}
                </div>
            </div>
            <Pivot>
                <PivotItem headerText="By User">
                    <div style={{ height: '70vh', position: 'relative', marginTop: 10, minWidth: 0, overflowX: 'auto' }}>
                        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                            <DetailsList
                                items={tasksByUser}
                                groups={getUserGroups()}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                layoutMode={DetailsListLayoutMode.fixedColumns}
                                constrainMode={ConstrainMode.unconstrained}
                                onRenderDetailsHeader={
                                    (props: IDetailsHeaderProps | undefined, defaultRender?: (props: IDetailsHeaderProps) => JSX.Element | null): JSX.Element | null => {
                                        if (!props || !defaultRender) return null;
                                        return (
                                            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                                                {defaultRender(props)}
                                            </Sticky>
                                        );
                                    }
                                }
                            />
                        </ScrollablePane>
                    </div>
                </PivotItem>
                <PivotItem headerText="By Status">
                    <div style={{ height: '70vh', position: 'relative', marginTop: 10, minWidth: 0, overflowX: 'auto' }}>
                        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                            <DetailsList
                                items={tasksByStatus}
                                groups={getStatusGroups()}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                layoutMode={DetailsListLayoutMode.fixedColumns}
                                constrainMode={ConstrainMode.unconstrained}
                                onRenderDetailsHeader={
                                    (props: IDetailsHeaderProps | undefined, defaultRender?: (props: IDetailsHeaderProps) => JSX.Element | null): JSX.Element | null => {
                                        if (!props || !defaultRender) return null;
                                        return (
                                            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                                                {defaultRender(props)}
                                            </Sticky>
                                        );
                                    }
                                }
                            />
                        </ScrollablePane>
                    </div>
                </PivotItem>
                <PivotItem headerText="Reports & Analytics">
                    <AdminReports tasks={filteredTasks} />
                </PivotItem>
            </Pivot>

            <Panel
                isOpen={isCreating}
                onDismiss={() => setIsCreating(false)}
                headerText="Create New Main Task"
                type={PanelType.large}
            >
                <Stack tokens={{ childrenGap: 15 }}>
                    <TextField label="Task Title" value={newTitle} onChange={(e, v) => setNewTitle(v || '')} required />
                    <TextField label="Description" multiline rows={3} value={newDesc} onChange={(e, v) => setNewDesc(v || '')} required />

                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <Dropdown
                            label="Year"
                            placeholder="Select Year"
                            options={[{ key: '2024', text: '2024' }, { key: '2025', text: '2025' }, { key: '2026', text: '2026' }]}
                            selectedKey={newYear}
                            onChange={(_, opt) => setNewYear(opt?.key as string || '')}
                            required
                            styles={{ root: { width: '50%' } }}
                        />
                        <Dropdown
                            label="Month"
                            placeholder="Select Month"
                            options={[
                                'January', 'February', 'March', 'April', 'May', 'June',
                                'July', 'August', 'September', 'October', 'November', 'December'
                            ].map(m => ({ key: m, text: m }))}
                            selectedKey={newMonth}
                            onChange={(_, opt) => setNewMonth(opt?.key as string || '')}
                            required
                            styles={{ root: { width: '50%' } }}
                        />
                    </Stack>

                    <ComboBox
                        label="Assign To"
                        options={userOptions}
                        selectedKey={newAssignKey}
                        onChange={(e, opt) => {
                            if (opt) {
                                setNewAssignKey(opt.key as number);
                            } else {
                                setNewAssignKey(undefined);
                            }
                        }}
                        allowFreeform={true}
                        autoComplete="on"
                        useComboBoxAsMenuWidth
                        required
                    />

                    <ComboBox
                        label="Department"
                        options={departmentOptions}
                        selectedKey={newDept}
                        onChange={(_, opt) => setNewDept(opt?.key as string || opt?.text || '')}
                        allowFreeform={false}
                    />
                    <TextField
                        label="Project"
                        value={newProject}
                        onChange={(e, v) => setNewProject(v || '')}
                    />

                    <DatePicker label="Due Date" value={newDueDate} onSelectDate={(d) => setNewDueDate(d || undefined)} isRequired />

                    <div>
                        <label style={{ fontWeight: 600, display: 'block', marginBottom: 4 }}>Attachments</label>
                        <input
                            type="file"
                            multiple
                            onChange={(e) => {
                                const fileList = e.target.files;
                                if (fileList) {
                                    const files: File[] = [];
                                    for (let i = 0; i < fileList.length; i++) {
                                        files.push(fileList[i]);
                                    }
                                    setAttachFiles(files);
                                } else {
                                    setAttachFiles([]);
                                }
                            }}
                        />
                        {attachFiles.length > 0 && (
                            <div style={{ marginTop: 8, fontSize: 12, color: '#666' }}>
                                {attachFiles.map((f, i) => <div key={i}>{f.name}</div>)}
                            </div>
                        )}
                    </div>

                    {msg && <MessageBar messageBarType={MessageBarType.error}>{msg}</MessageBar>}

                    <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                        <PrimaryButton text="Create Task" onClick={handleCreateMainTask} />
                        <DefaultButton text="Cancel" onClick={() => setIsCreating(false)} />
                    </Stack>
                </Stack>
            </Panel>

            {/* Status Popup Panel */}
            <Panel
                isOpen={!!selectedStatusPopup}
                onDismiss={() => setSelectedStatusPopup(undefined)}
                headerText={`Tasks - ${selectedStatusPopup}`}
                type={PanelType.medium}
            >
                <div style={{ marginBottom: 10 }}>
                    <strong>Total: {filteredTasks.filter((t: IMainTask) => t.Status === selectedStatusPopup).length} tasks</strong>
                </div>
                <DetailsList
                    items={filteredTasks.filter((t: IMainTask) => t.Status === selectedStatusPopup)}
                    columns={[
                        { key: 'id', name: '#', fieldName: 'Id', minWidth: 30, maxWidth: 50 },
                        { key: 'title', name: 'Task', fieldName: 'Title', minWidth: 150, isResizable: true },
                        { key: 'assigned', name: 'Task Assigned', minWidth: 120, onRender: (i: IMainTask) => getAssignedUser(i).Title },
                        { key: 'month', name: 'SMT Month', minWidth: 80, onRender: (i: IMainTask) => (i as any).SMTMonth || '' },
                        { key: 'due', name: 'Due Date', minWidth: 90, onRender: (i: IMainTask) => i.TaskDueDate ? new Date(i.TaskDueDate).toLocaleDateString() : '' }
                    ]}
                    selectionMode={SelectionMode.none}
                    compact
                />
            </Panel>
            {/* Task Detail Panel (The "View" popup) */}
            <Panel
                isOpen={!!selectedTaskForDetail}
                onDismiss={() => {
                    setSelectedTaskForDetail(null);
                    processedMainIdRef.current = undefined; // Reset ref so it can re-open if clicked again
                }}
                type={PanelType.extraLarge}
                // Header is handled inside TaskDetail now or we can set it here? 
                // TaskDetail has its own header internal structure now but Panel needs a header text usually.
                // We'll leave headerText empty and let TaskDetail render the title.
                headerText=""
            >
                {selectedTaskForDetail && (
                    <TaskDetail
                        mainTask={selectedTaskForDetail}
                        initialChildTaskId={props.initialChildTaskId}
                        initialTab={props.initialTab}
                        onDeepLinkProcessed={props.onChildDeepLinkProcessed}
                    />
                )}
            </Panel>
        </div>
    );
};
