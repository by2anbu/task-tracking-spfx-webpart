
import * as React from 'react';
import {
    Stack,
    Text,
    SearchBox,
    Dropdown,
    IDropdownOption,
    IconButton,
    PrimaryButton,
    DefaultButton,
    Persona,
    PersonaSize,
    TooltipHost,
    Icon,
    Pivot,
    PivotItem,
    Panel,
    PanelType,
    TextField,
    MessageBar,
    MessageBarType
} from 'office-ui-fabric-react';
import styles from './SubtaskDashboard.module.scss';
import { taskService } from '../../../../services/sp-service';
import { ISubTask, IMainTask } from '../../../../services/interfaces';
import { sanitizeHtml } from '../../../../utils/sanitize';

export interface ISubtaskDashboardProps {
    userEmail: string;
}

interface IUiTask extends ISubTask {
    depth: number;
    hasChildren: boolean;
    isExpanded?: boolean;
    computedProgress: number;
    priorityColor?: string;
}

export const SubtaskDashboard: React.FunctionComponent<ISubtaskDashboardProps> = (props) => {
    const [allTasks, setAllTasks] = React.useState<ISubTask[]>([]);
    const [processedTasks, setProcessedTasks] = React.useState<IUiTask[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [viewMode, setViewMode] = React.useState<'table' | 'card' | 'gantt'>('table');
    const [expandedRows, setExpandedRows] = React.useState<Set<number>>(new Set());

    // Filters
    const [statusFilter, setStatusFilter] = React.useState<string>('all');
    const [categoryFilter, setCategoryFilter] = React.useState<string>('all');
    const [searchTerm, setSearchTerm] = React.useState<string>('');
    const [filteredTasks, setFilteredTasks] = React.useState<IUiTask[]>([]);
    const [sortField, setSortField] = React.useState<string | undefined>(undefined);
    const [sortDesc, setSortDesc] = React.useState<boolean>(false);

    // Edit Panel State
    const [selectedTask, setSelectedTask] = React.useState<IUiTask | undefined>(undefined);
    const [editStatus, setEditStatus] = React.useState<string>('');
    const [editRemarks, setEditRemarks] = React.useState<string>('');
    const [newFiles, setNewFiles] = React.useState<File[]>([]);
    const [saving, setSaving] = React.useState(false);
    const [message, setMessage] = React.useState<string | undefined>(undefined);

    // Clarification State
    const [clarificationHistory, setClarificationHistory] = React.useState<any[]>([]);
    const [clarificationMessage, setClarificationMessage] = React.useState('');
    const [loadingClarification, setLoadingClarification] = React.useState(false);
    const [clarificationMetadata, setClarificationMetadata] = React.useState<Map<number, { hasCorrespondence: boolean, isReply: boolean }>>(new Map());
    const chatContainerRef = React.useRef<HTMLDivElement>(null);
    const lastTaskIdRef = React.useRef<number | undefined>(undefined);
    const prevReplyCountRef = React.useRef<number>(0);

    React.useEffect(() => {
        if (chatContainerRef.current) {
            chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
        }
    }, [clarificationHistory]);

    React.useEffect(() => {
        loadData().catch(console.error);
    }, [props.userEmail]);

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

    React.useEffect(() => {
        applyFilters();
    }, [processedTasks, statusFilter, categoryFilter, searchTerm, sortField, sortDesc]);

    const getAssigneeName = (item: ISubTask): string => {
        const assigned = item.TaskAssignedTo;
        if (Array.isArray(assigned) && assigned.length > 0) {
            return assigned.map(u => u.Title).join(', ');
        }
        if (assigned && (assigned as any).Title) {
            return (assigned as any).Title;
        }
        return 'Unassigned';
    };

    const getAssigneeEmail = (item: ISubTask): string | undefined => {
        const assigned = item.TaskAssignedTo;
        if (Array.isArray(assigned) && assigned.length > 0) return (assigned[0] as any).EMail || (assigned[0] as any).Email;
        if (assigned) return (assigned as any).EMail || (assigned as any).Email;
        return undefined;
    };

    const applyFilters = () => {
        let filtered = [...processedTasks];

        if (statusFilter !== 'all') {
            filtered = filtered.filter(t => t.TaskStatus === statusFilter);
        }

        if (categoryFilter !== 'all') {
            filtered = filtered.filter(t => t.Category === categoryFilter);
        }

        if (searchTerm) {
            const lowerSearch = searchTerm.toLowerCase();
            filtered = filtered.filter(t =>
                (t.Task_Title || '').toLowerCase().indexOf(lowerSearch) > -1 ||
                getAssigneeName(t).toLowerCase().indexOf(lowerSearch) > -1
            );
        }

        // Sorting
        if (sortField) {
            filtered.sort((a: any, b: any) => {
                let valA = a[sortField];
                let valB = b[sortField];

                if (sortField === 'TaskDueDate' || sortField === 'TaskStartDate') {
                    valA = valA ? new Date(valA).getTime() : 0;
                    valB = valB ? new Date(valB).getTime() : 0;
                }

                if (valA < valB) return sortDesc ? 1 : -1;
                if (valA > valB) return sortDesc ? -1 : 1;
                return 0;
            });
        }

        setFilteredTasks(filtered);
    };

    const handleSort = (field: string) => {
        if (sortField === field) {
            setSortDesc(!sortDesc);
        } else {
            setSortField(field);
            setSortDesc(false);
        }
    };

    const loadData = async () => {
        setLoading(true);
        try {
            const tasks = await taskService.getSubTasksForUser(props.userEmail);
            setAllTasks(tasks);
            processHierarchy(tasks);

            // Fetch clarification metadata
            const ids = tasks.map(t => t.Id);
            const metadata = await taskService.getTaskCorrespondenceMetadata(ids);
            setClarificationMetadata(metadata);
        } catch (error) {
            console.error("Dashboard Load Error:", error);
        } finally {
            setLoading(false);
        }
    };

    const calculateProgress = (status: string): number => {
        switch (status) {
            case 'Completed': return 100;
            case 'In Progress': return 50;
            case 'On Hold': return 25;
            default: return 0;
        }
    };

    const processHierarchy = (items: ISubTask[]) => {
        const itemMap = new Map<number, ISubTask>();
        const childrenMap = new Map<number, number[]>();

        items.forEach(item => {
            itemMap.set(item.Id, item);
            const pId = item.ParentSubtaskId || 0;
            if (!childrenMap.has(pId)) childrenMap.set(pId, []);
            childrenMap.get(pId)!.push(item.Id);
        });

        const result: IUiTask[] = [];

        const traverse = (parentId: number, depth: number) => {
            const children = childrenMap.get(parentId) || [];
            children.forEach(childId => {
                const child = itemMap.get(childId);
                if (child) {
                    const hasChildren = childrenMap.has(child.Id);
                    result.push({
                        ...child,
                        depth,
                        hasChildren,
                        computedProgress: calculateProgress(child.TaskStatus)
                    });
                    traverse(child.Id, depth + 1);
                }
            });
        };

        traverse(0, 0);
        setProcessedTasks(result);
    };

    const toggleExpand = (id: number) => {
        const nextSet = new Set<number>();
        expandedRows.forEach(r => nextSet.add(r));

        if (nextSet.has(id)) nextSet.delete(id);
        else nextSet.add(id);
        setExpandedRows(nextSet);
    };

    const formatDate = (dateStr: string | undefined): string => {
        if (!dateStr) return '-';
        const date = new Date(dateStr);
        const day = date.getDate();
        const dayStr = day < 10 ? '0' + day : '' + day;
        const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        return `${dayStr}-${months[date.getMonth()]}-${date.getFullYear()}`;
    };

    const isOverdue = (dateStr: string, status: string): boolean => {
        if (status === 'Completed' || !dateStr) return false;
        return new Date(dateStr) < new Date();
    };

    const openEditPanel = async (task: IUiTask) => {
        setSelectedTask(task);
        setEditStatus(task.TaskStatus);
        setEditRemarks(task.User_Remarks || '');
        setNewFiles([]);
        setMessage(undefined);

        // Fetch clarification history
        if (lastTaskIdRef.current !== task.Id) {
            setClarificationHistory([]);
            lastTaskIdRef.current = task.Id;
        }

        setLoadingClarification(true);
        setClarificationMessage('');
        try {
            const history = await taskService.getCorrespondenceByTaskId(task.Admin_Job_ID, task.Id);
            setClarificationHistory(prev => {
                // If we reopen the SAME task, and SPO is lagging, keep the optimistic state
                if (history.length < prev.length && prev.length > 0) {
                    console.log("[Correspondence] Reopen: SPO lag detected, keeping state");
                    return prev;
                }
                return history;
            });
        } catch (e) {
            console.error("Error fetching clarification history:", e);
        } finally {
            setLoadingClarification(false);
        }
    };

    const closeEditPanel = () => {
        setSelectedTask(undefined);
    };

    const handleSaveTask = async () => {
        if (!selectedTask) return;
        setSaving(true);
        try {
            await taskService.updateSubTaskStatus(
                selectedTask.Id,
                selectedTask.Admin_Job_ID,
                editStatus,
                editRemarks
            );

            if (newFiles.length > 0) {
                // Use the list name from interfaces if available, or fallback to the known name "Task Tracking System User"
                await taskService.addAttachmentsToItem("Task Tracking System User", selectedTask.Id, newFiles);
            }

            setMessage("Task updated successfully!");
            setTimeout(() => {
                closeEditPanel();
                loadData().catch(console.error);
            }, 1000);
        } catch (error) {
            console.error(error);
            setMessage("Error updating task: " + error.message);
        } finally {
            setSaving(false);
        }
    };

    const handleSendClarification = async () => {
        if (!selectedTask || !clarificationMessage.trim()) return;
        setSaving(true);
        try {
            // Send email/correspondence
            // We'll send to the "Author" of the main task (the requester)
            const mainTask = await taskService.getMainTaskById(selectedTask.Admin_Job_ID);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const toEmail = (mainTask as any)?.Author?.EMail || '';

            const tempMsg = {
                FromAddress: props.userEmail,
                MessageBody: clarificationMessage,
                Created: new Date().toISOString(),
                Author: { Title: 'You' } // Temporary author for optimistic update
            };
            setClarificationHistory(prev => [...prev, tempMsg]);
            const currentMsg = clarificationMessage;
            setClarificationMessage('');

            await taskService.sendEmail(
                toEmail ? [toEmail] : [],
                `Clarification Needed: ${selectedTask.Task_Title}`,
                currentMsg,
                selectedTask.Admin_Job_ID,
                selectedTask.Id
            );

            // Fetch actual history after a short delay
            setTimeout(async () => {
                const history = await taskService.getCorrespondenceByTaskId(selectedTask.Admin_Job_ID, selectedTask.Id);
                setClarificationHistory(prev => {
                    // Stale sync protection: If SharePoint returns fewer items than we currently show 
                    // (which includes our optimistic temp message), it's still lagging. 
                    // Keep the local state until SPO catches up.
                    if (history.length < prev.length && prev.length > 0) {
                        console.log("[Correspondence] SPO lag detected, keeping optimistic state");
                        return prev;
                    }
                    return history;
                });
            }, 2000);

            // Update indicator
            setClarificationMetadata(prev => {
                const next = new Map<number, { hasCorrespondence: boolean, isReply: boolean }>();
                prev.forEach((v, k) => next.set(k, v));
                next.set(selectedTask.Id, { hasCorrespondence: true, isReply: false });
                return next;
            });
        } catch (e) {
            console.error("Error sending clarification:", e);
        } finally {
            setSaving(false);
        }
    };

    const exportToExcel = () => {
        const headers = ["ID", "Task Name", "Status", "Category", "Assigned To", "Due Date", "Progress"];
        const rows = filteredTasks.map(t => [
            t.Id,
            t.Task_Title,
            t.TaskStatus,
            t.Category || "General",
            getAssigneeName(t),
            formatDate(t.TaskDueDate),
            `${t.computedProgress}%`
        ]);

        const csvContent = "data:text/csv;charset=utf-8,"
            + headers.join(",") + "\n"
            + rows.map(e => e.join(",")).join("\n");

        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `Subtask_Export_${new Date().getTime()}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    const renderProgressBar = (progress: number) => (
        <div className={styles.progressWrapper}>
            <div className={styles.progressBar}>
                <div className={styles.progressFill} style={{ width: `${progress}%` }} />
            </div>
            <div className={styles.progressLabel}>
                <span>Completion</span>
                <span>{progress}%</span>
            </div>
        </div>
    );

    const renderTable = () => (
        <table className={styles.customTable}>
            <thead>
                <tr>
                    <th style={{ cursor: 'pointer' }} onClick={() => handleSort('Task_Title')}>
                        Task Name {sortField === 'Task_Title' && <Icon iconName={sortDesc ? 'SortDown' : 'SortUp'} />}
                    </th>
                    <th style={{ cursor: 'pointer' }} onClick={() => handleSort('TaskStatus')}>
                        Status {sortField === 'TaskStatus' && <Icon iconName={sortDesc ? 'SortDown' : 'SortUp'} />}
                    </th>
                    <th>Category</th>
                    <th>Assigned To</th>
                    <th style={{ cursor: 'pointer' }} onClick={() => handleSort('TaskDueDate')}>
                        Due Date {sortField === 'TaskDueDate' && <Icon iconName={sortDesc ? 'SortDown' : 'SortUp'} />}
                    </th>
                    <th style={{ cursor: 'pointer' }} onClick={() => handleSort('computedProgress')}>
                        Progress {sortField === 'computedProgress' && <Icon iconName={sortDesc ? 'SortDown' : 'SortUp'} />}
                    </th>
                </tr>
            </thead>
            <tbody>
                {filteredTasks.map(task => { // Use filteredTasks here
                    const overdue = isOverdue(task.TaskDueDate, task.TaskStatus);
                    return (
                        <tr
                            key={task.Id}
                            className={styles.animatedFadeIn}
                            onClick={() => openEditPanel(task)}
                            style={{ cursor: 'pointer' }}
                        >
                            <td style={{ paddingLeft: `${task.depth * 24 + 16}px` }}>
                                <Stack horizontal verticalAlign="center">
                                    {task.hasChildren && (
                                        <IconButton
                                            iconProps={{ iconName: expandedRows.has(task.Id) ? 'ChevronDown' : 'ChevronRight' }}
                                            onClick={() => toggleExpand(task.Id)}
                                            styles={{ root: { height: 24, width: 24 } }}
                                        />
                                    )}
                                    <Text variant="medium">
                                        <span style={{ fontWeight: 600 }}>{task.Task_Title}</span>
                                    </Text>
                                    {overdue && (
                                        <TooltipHost content="Overdue Task">
                                            <Icon iconName="ErrorBadge" className={styles.overdueAlert} style={{ marginLeft: 8 }} />
                                        </TooltipHost>
                                    )}
                                    {clarificationMetadata.has(task.Id) && (
                                        <TooltipHost content={clarificationMetadata.get(task.Id)?.isReply ? "New reply received!" : "Correspondence history exists"}>
                                            <Icon
                                                iconName={clarificationMetadata.get(task.Id)?.isReply ? "Ringer" : "Questionnaire"}
                                                style={{
                                                    marginLeft: 8,
                                                    color: clarificationMetadata.get(task.Id)?.isReply ? '#d13438' : '#0078d4',
                                                    fontSize: 14,
                                                    fontWeight: clarificationMetadata.get(task.Id)?.isReply ? 700 : 400,
                                                    animation: clarificationMetadata.get(task.Id)?.isReply ? 'pulse 2s infinite' : 'none'
                                                }}
                                            />
                                        </TooltipHost>
                                    )}
                                </Stack>
                            </td>
                            <td>
                                <span className={`${styles.statusBadge} ${(styles as any)[task.TaskStatus.replace(/\s/g, '')] || styles.notStarted}`}>
                                    {task.TaskStatus}
                                </span>
                            </td>
                            <td>
                                <Text variant="small">{task.Category || 'General'}</Text>
                            </td>
                            <td>
                                <Persona
                                    text={getAssigneeName(task)}
                                    imageUrl={getAssigneeEmail(task) ? `/_layouts/15/userphoto.aspx?size=S&username=${getAssigneeEmail(task)}` : undefined}
                                    size={PersonaSize.size24}
                                />
                            </td>
                            <td>
                                <Text variant="small" style={{ color: overdue ? '#ef4444' : 'inherit' }}>
                                    {formatDate(task.TaskDueDate)}
                                </Text>
                            </td>
                            <td>{renderProgressBar(task.computedProgress)}</td>
                        </tr>
                    );
                })}
            </tbody>
        </table>
    );

    const renderCards = () => (
        <div className={styles.cardGrid}>
            {filteredTasks.map(task => ( // Use filteredTasks here
                <div
                    key={task.Id}
                    className={`${styles.taskCard} ${styles.animatedFadeIn}`}
                    onClick={() => openEditPanel(task)}
                    style={{ cursor: 'pointer' }}
                >
                    <div className={styles.cardHeader}>
                        <Text variant="large">
                            <span style={{ fontWeight: 700 }}>{task.Task_Title}</span>
                        </Text>
                        <span className={`${styles.statusBadge} ${(styles as any)[task.TaskStatus.replace(/\s/g, '')] || styles.notStarted}`}>
                            {task.TaskStatus}
                        </span>
                        {clarificationMetadata.has(task.Id) && (
                            <TooltipHost content={clarificationMetadata.get(task.Id)?.isReply ? "New reply received!" : "Correspondence history exists"}>
                                <Icon
                                    iconName={clarificationMetadata.get(task.Id)?.isReply ? "Ringer" : "Questionnaire"}
                                    style={{
                                        marginLeft: 8,
                                        color: clarificationMetadata.get(task.Id)?.isReply ? '#d13438' : '#0078d4',
                                        fontSize: 14,
                                        fontWeight: clarificationMetadata.get(task.Id)?.isReply ? 700 : 400,
                                        animation: clarificationMetadata.get(task.Id)?.isReply ? 'pulse 2s infinite' : 'none'
                                    }}
                                />
                            </TooltipHost>
                        )}
                    </div>
                    <Text variant="small" block style={{ margin: '8px 0', color: '#64748b' }}>
                        {task.Task_Description}
                    </Text>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginTop: 12 }}>
                        <Persona
                            text={getAssigneeName(task)}
                            imageUrl={getAssigneeEmail(task) ? `/_layouts/15/userphoto.aspx?size=S&username=${getAssigneeEmail(task)}` : undefined}
                            size={PersonaSize.size24}
                        />
                        <Stack verticalAlign="end">
                            <Text variant="small" style={{ color: '#64748b' }}>{task.Category}</Text>
                            <Text variant="small" style={{ color: isOverdue(task.TaskDueDate, task.TaskStatus) ? '#ef4444' : '#64748b' }}>
                                Due: {formatDate(task.TaskDueDate)}
                            </Text>
                        </Stack>
                    </Stack>
                    {renderProgressBar(task.computedProgress)}
                </div>
            ))}
        </div>
    );

    // Get unique categories for dropdown
    const categories: IDropdownOption[] = [{ key: 'all', text: 'All Categories' }];
    const uniqueCats = new Set<string>();
    processedTasks.forEach(t => { if (t.Category) uniqueCats.add(t.Category); });
    uniqueCats.forEach(cat => categories.push({ key: cat, text: cat }));

    return (
        <div className={styles.subtaskDashboard}>
            <div className={styles.header}>
                <div>
                    <h1>Subtask Productivity Dashboard</h1>
                    <Text variant="medium">Real-time performance and tracking for all project sub-items.</Text>
                </div>
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                    <PrimaryButton iconProps={{ iconName: 'ExcelDocument' }} text="Export CSV" onClick={exportToExcel} />
                </Stack>
            </div>

            <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 16 }} className={styles.controls}>
                <Stack.Item grow>
                    <SearchBox
                        placeholder="Search tasks or assignees..."
                        onSearch={val => setSearchTerm(val)}
                        onChange={(_, val) => setSearchTerm(val || '')}
                    />
                </Stack.Item>
                <Dropdown
                    label="Status"
                    selectedKey={statusFilter}
                    options={[
                        { key: 'all', text: 'All Statuses' },
                        { key: 'Not Started', text: 'Not Started' },
                        { key: 'In Progress', text: 'In Progress' },
                        { key: 'Completed', text: 'Completed' },
                        { key: 'On Hold', text: 'On Hold' }
                    ]}
                    styles={{ root: { width: 180 } }}
                    onChange={(_, opt) => setStatusFilter(opt?.key as string)}
                />
                <Dropdown
                    label="Category (Priority)"
                    selectedKey={categoryFilter}
                    options={categories}
                    styles={{ root: { width: 180 } }}
                    onChange={(_, opt) => setCategoryFilter(opt?.key as string)}
                />
                <Pivot onLinkClick={(item) => setViewMode(item?.props.itemKey as any)}>
                    <PivotItem headerText="Table" itemKey="table" itemIcon="Table" />
                    <PivotItem headerText="Cards" itemKey="card" itemIcon="GridViewSmall" />
                    <PivotItem headerText="Gantt" itemKey="gantt" itemIcon="TimelineMatrixView" />
                </Pivot>
            </Stack>

            <div className={styles.viewContainer}>
                {loading ? (
                    <Text>Loading dashboard...</Text>
                ) : (
                    <>
                        {viewMode === 'table' && renderTable()}
                        {viewMode === 'card' && renderCards()}
                        {viewMode === 'gantt' && (
                            <div className={styles.ganttContainer}>
                                <Text variant="large">
                                    <span style={{ fontWeight: 700 }}>Project Timeline Overview</span>
                                </Text>
                                <div style={{ marginTop: 20 }}>
                                    {filteredTasks.map(t => {
                                        const overdue = isOverdue(t.TaskDueDate, t.TaskStatus);
                                        let barColor = '#3b82f6'; // Default Blue
                                        if (t.TaskStatus === 'Completed') barColor = '#10b981'; // Green
                                        else if (overdue) barColor = '#ef4444'; // Red
                                        else if (t.TaskStatus === 'Not Started') barColor = '#94a3b8'; // Grey

                                        // Calculate Overdue Days
                                        let overdueDays = 0;
                                        const dueDate = t.TaskDueDate ? new Date(t.TaskDueDate) : undefined;
                                        const endDate = t.Task_End_Date ? new Date(t.Task_End_Date) : (t.TaskStatus === 'Completed' ? (t.TaskDueDate ? new Date(t.TaskDueDate) : undefined) : new Date());

                                        if (dueDate && endDate && endDate > dueDate) {
                                            const diffTime = Math.abs(endDate.getTime() - dueDate.getTime());
                                            overdueDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                                        }

                                        return (
                                            <div key={t.Id} style={{ marginBottom: 16 }}>
                                                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                                                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                                                        <Text variant="small" style={{ fontWeight: 600 }}>{t.Task_Title}</Text>
                                                        {overdueDays > 0 && (
                                                            <span style={{
                                                                fontSize: '10px',
                                                                padding: '2px 6px',
                                                                borderRadius: '4px',
                                                                backgroundColor: t.TaskStatus === 'Completed' ? '#ffedd5' : '#fee2e2',
                                                                color: t.TaskStatus === 'Completed' ? '#9a3412' : '#b91c1c',
                                                                fontWeight: 600
                                                            }}>
                                                                {overdueDays} DAYS DELAYED
                                                            </span>
                                                        )}
                                                    </Stack>
                                                    <Text variant="small" style={{ color: '#64748b' }}>
                                                        {t.TaskStatus === 'Completed' ? (
                                                            <>Due: {formatDate(t.TaskDueDate)} | <span style={{ color: overdueDays > 0 ? '#b91c1c' : '#10b981', fontWeight: 600 }}>Finished: {formatDate(t.Task_End_Date || t.TaskDueDate)}</span></>
                                                        ) : (
                                                            <>Due: {formatDate(t.TaskDueDate)}</>
                                                        )}
                                                    </Text>
                                                </Stack>
                                                <div style={{ height: 12, background: '#f1f5f9', borderRadius: 6, position: 'relative', marginTop: 4, overflow: 'hidden' }}>
                                                    <div style={{
                                                        width: `${t.computedProgress}%`,
                                                        height: '100%',
                                                        background: barColor,
                                                        borderRadius: 6,
                                                        transition: 'width 1s ease-in-out, background-color 0.3s ease'
                                                    }} />
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                            </div>
                        )}
                    </>
                )}
            </div>

            {/* Edit Panel */}
            <Panel
                isOpen={!!selectedTask}
                onDismiss={closeEditPanel}
                type={PanelType.medium}
                headerText="Edit Subtask"
                closeButtonAriaLabel="Close"
            >
                {selectedTask && (
                    <Stack tokens={{ childrenGap: 20 }} style={{ padding: 10 }}>
                        <div style={{ borderBottom: '1px solid #edebe9', paddingBottom: 15 }}>
                            <div style={{ marginBottom: 10 }}>
                                <Text variant="small" style={{ color: '#605e5c', display: 'block' }}>Title</Text>
                                <Text variant="large" style={{ fontWeight: 600 }}>{selectedTask.Task_Title}</Text>
                            </div>
                            <div style={{ marginBottom: 10 }}>
                                <Text variant="small" style={{ color: '#605e5c', display: 'block' }}>Description</Text>
                                <Text variant="medium">{selectedTask.Task_Description || 'No description'}</Text>
                            </div>
                            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
                                <div>
                                    <Text variant="small" style={{ color: '#605e5c', display: 'block' }}>Assigned To</Text>
                                    <Persona
                                        text={getAssigneeName(selectedTask)}
                                        size={PersonaSize.size24}
                                        imageUrl={getAssigneeEmail(selectedTask) ? `/_layouts/15/userphoto.aspx?size=S&username=${getAssigneeEmail(selectedTask)}` : undefined}
                                    />
                                </div>
                                <div>
                                    <Text variant="small" style={{ color: '#605e5c', display: 'block' }}>Due Date</Text>
                                    <Text variant="medium" style={{ fontWeight: 600, color: isOverdue(selectedTask.TaskDueDate, selectedTask.TaskStatus) ? '#ef4444' : 'inherit' }}>
                                        {formatDate(selectedTask.TaskDueDate)}
                                    </Text>
                                </div>
                                <div>
                                    <Text variant="small" style={{ color: '#605e5c', display: 'block' }}>Category</Text>
                                    <Text variant="medium">{selectedTask.Category || 'General'}</Text>
                                </div>
                            </div>
                        </div>

                        <Dropdown
                            label="Status"
                            selectedKey={editStatus}
                            options={[
                                { key: 'Not Started', text: 'Not Started' },
                                { key: 'In Progress', text: 'In Progress' },
                                { key: 'Completed', text: 'Completed' },
                                { key: 'On Hold', text: 'On Hold' }
                            ]}
                            onChange={(_, opt) => setEditStatus(opt?.key as string)}
                        />

                        <TextField
                            label="Remarks"
                            multiline
                            rows={4}
                            value={editRemarks}
                            onChange={(_, val) => setEditRemarks(val || '')}
                        />

                        <div>
                            <Text variant="small" style={{ fontWeight: 600, display: 'block', marginBottom: 8 }}>Current Attachments:</Text>
                            {selectedTask.AttachmentFiles && selectedTask.AttachmentFiles.length > 0 ? (
                                <div style={{ marginBottom: 12 }}>
                                    {selectedTask.AttachmentFiles.map((file, idx) => (
                                        <div key={idx} style={{ marginBottom: 4 }}>
                                            <IconButton
                                                iconProps={{ iconName: 'Download' }}
                                                title="View Attachment"
                                                onClick={() => window.open(file.ServerRelativeUrl, '_blank')}
                                                styles={{ root: { height: 20, width: 20 } }}
                                            />
                                            <Text variant="small" style={{ marginLeft: 8 }}>{file.FileName}</Text>
                                        </div>
                                    ))}
                                </div>
                            ) : (
                                <Text variant="small" style={{ display: 'block', marginBottom: 12, color: '#605e5c' }}>No attachments</Text>
                            )}

                            <Text variant="small" style={{ fontWeight: 600, display: 'block', marginBottom: 8 }}>Add New Attachments:</Text>
                            <input
                                type="file"
                                multiple
                                onChange={(e) => {
                                    const files = e.target.files;
                                    if (files) {
                                        const fileArray: File[] = [];
                                        for (let i = 0; i < files.length; i++) fileArray.push(files[i]);
                                        setNewFiles(fileArray);
                                    }
                                }}
                            />
                            {newFiles.length > 0 && (
                                <div style={{ marginTop: 8 }}>
                                    {newFiles.map((f, i) => (
                                        <div key={i} style={{ fontSize: 12, color: '#0078d4' }}>
                                            <Icon iconName="Attach" style={{ marginRight: 4 }} /> {f.name}
                                        </div>
                                    ))}
                                </div>
                            )}
                        </div>

                        {/* Clarification Section */}
                        <div style={{ borderTop: '1px solid #edebe9', paddingTop: 20, marginTop: 10 }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 15 }}>
                                <Text variant="large" style={{ fontWeight: 600 }}>Task Correspondence</Text>
                                <IconButton
                                    iconProps={{ iconName: 'Refresh' }}
                                    title="Refresh history"
                                    onClick={async () => {
                                        if (!selectedTask) return;
                                        setLoadingClarification(true);
                                        const h = await taskService.getCorrespondenceByTaskId(selectedTask.Admin_Job_ID, selectedTask.Id);
                                        setClarificationHistory(h);
                                        setLoadingClarification(false);
                                    }}
                                    disabled={loadingClarification}
                                />
                            </div>

                            <div ref={chatContainerRef} style={{
                                maxHeight: 250,
                                overflowY: 'auto',
                                background: '#f8fafc',
                                padding: 12,
                                borderRadius: 8,
                                border: '1px solid #e2e8f0',
                                marginBottom: 15
                            }}>
                                {loadingClarification ? (
                                    <Text variant="small" style={{ textAlign: 'center', display: 'block', padding: 20 }}>Loading history...</Text>
                                ) : clarificationHistory.length === 0 ? (
                                    <Text variant="small" style={{ textAlign: 'center', display: 'block', padding: 20, color: '#64748b' }}>No history found.</Text>
                                ) : (
                                    <Stack tokens={{ childrenGap: 10 }}>
                                        {clarificationHistory.map((msg, i) => (
                                            <div key={i} style={{
                                                padding: '8px 12px',
                                                background: msg.FromAddress === props.userEmail ? '#eff6ff' : 'white',
                                                border: '1px solid #e2e8f0',
                                                borderRadius: 8,
                                                alignSelf: msg.FromAddress === props.userEmail ? 'flex-end' : 'flex-start',
                                                maxWidth: '90%'
                                            }}>
                                                <div style={{ fontSize: 10, color: '#64748b', marginBottom: 4, display: 'flex', justifyContent: 'space-between' }}>
                                                    <strong>{msg.Author?.Title || msg.FromAddress}</strong>
                                                    <span style={{ marginLeft: 8 }}>{new Date(msg.Created).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                                                </div>
                                                <div dangerouslySetInnerHTML={{ __html: sanitizeHtml(msg.MessageBody) }} style={{ fontSize: 12 }} />
                                            </div>
                                        ))}
                                    </Stack>
                                )}
                            </div>

                            <TextField
                                label="Ask a question or respond"
                                multiline
                                rows={2}
                                value={clarificationMessage}
                                onChange={(_, v) => setClarificationMessage(v || '')}
                                placeholder="Need more info? Ask here..."
                                styles={{ root: { marginBottom: 10 } }}
                            />
                            <PrimaryButton
                                text="Send Message"
                                onClick={handleSendClarification}
                                disabled={saving || !clarificationMessage.trim()}
                                style={{ width: 'fit-content' }}
                            />
                        </div>

                        {message && (
                            <MessageBar messageBarType={message.indexOf('Error') > -1 ? MessageBarType.error : MessageBarType.success}>
                                {message}
                            </MessageBar>
                        )}

                        <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                            <PrimaryButton text="Save Changes" onClick={handleSaveTask} disabled={saving} />
                            <DefaultButton text="Cancel" onClick={closeEditPanel} disabled={saving} />
                        </Stack>
                    </Stack>
                )}
            </Panel>
        </div>
    );
};

