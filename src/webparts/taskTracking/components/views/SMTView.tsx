import * as React from 'react';
import styles from '../TaskTracking.module.scss';
import { DetailsList, SelectionMode, IColumn, Panel, PanelType, TextField, PrimaryButton, DefaultButton, DatePicker, Stack, MessageBar, MessageBarType, Dropdown, Checkbox, Pivot, PivotItem, ScrollablePane, ScrollbarVisibility, IconButton, ComboBox, IComboBoxOption, Sticky, StickyPositionType, DetailsListLayoutMode, ConstrainMode, IDetailsHeaderProps } from 'office-ui-fabric-react';
import { taskService } from '../../../../services/sp-service';
import { IMainTask, ISubTask, LIST_SUB_TASKS, LIST_MAIN_TASKS } from '../../../../services/interfaces';
import { TaskDetail } from './TaskDetail';

export interface ISMTViewProps {
    userEmail: string;
    initialParentTaskId?: number;
    initialChildTaskId?: number;
    initialViewTaskId?: number;
    initialTab?: string;
    onMainDeepLinkProcessed?: () => void;
    onChildDeepLinkProcessed?: () => void;
}

export const SMTView: React.FunctionComponent<ISMTViewProps> = (props) => {
    const [mainTasks, setMainTasks] = React.useState<IMainTask[]>([]);
    const [allSubTasks, setAllSubTasks] = React.useState<ISubTask[]>([]);
    const [selectedTask, setSelectedTask] = React.useState<IMainTask | null>(null);
    const [selectedSubtask, setSelectedSubtask] = React.useState<ISubTask | null>(null);
    const [loading, setLoading] = React.useState<boolean>(true);
    const processedMainIdRef = React.useRef<number | undefined>(undefined);

    // Filter State
    const [statusFilter, setStatusFilter] = React.useState<string | undefined>(undefined);
    const [categoryFilter, setCategoryFilter] = React.useState<string | undefined>(undefined);
    const [overdueFilter, setOverdueFilter] = React.useState<boolean>(false);

    // Sort State
    const [sortedColumn, setSortedColumn] = React.useState<string | undefined>(undefined);
    const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

    // User & Options State
    const [userOptions, setUserOptions] = React.useState<IComboBoxOption[]>([]);
    const [departmentOptions, setDepartmentOptions] = React.useState<IComboBoxOption[]>([]);

    // All main tasks for looking up parent task names
    const [allMainTasks, setAllMainTasks] = React.useState<IMainTask[]>([]);

    React.useEffect(() => {
        loadData().catch(console.error);
    }, [props.userEmail]);

    // Handle URL parameters for deep linking
    React.useEffect(() => {
        if (loading || (mainTasks.length === 0 && allMainTasks.length === 0)) return;

        const { initialParentTaskId, initialChildTaskId, initialViewTaskId } = props;
        const targetId = initialViewTaskId || initialParentTaskId;

        if (!targetId || processedMainIdRef.current === targetId) return;

        console.log('[SMTView] Processing deep link for ID:', targetId);

        const matchingTasks = allMainTasks.filter((t: IMainTask) => t.Id === targetId);
        const mainTask = matchingTasks.length > 0 ? matchingTasks[0] : null;

        if (mainTask) {
            console.log('[SMTView] Opening detail panel for task:', targetId);
            setSelectedTask(mainTask);
            processedMainIdRef.current = targetId;
            if (props.onMainDeepLinkProcessed) props.onMainDeepLinkProcessed();
        } else {
            console.warn('[SMTView] Task not found for ID:', targetId);
        }
    }, [mainTasks, allMainTasks, props.initialParentTaskId, props.initialChildTaskId, props.initialViewTaskId, props.initialTab]);

    const loadData = async () => {
        setLoading(true);
        try {
            // Load user's main tasks first
            const tasks = await taskService.getMainTasksForUser(props.userEmail);
            setMainTasks(tasks);
            console.log('[SMTView] User main tasks loaded:', tasks.length, 'tasks');
            console.log('[SMTView] Main task IDs:', tasks.map(t => t.Id));

            // Try to load all main tasks for lookups
            let allTasks: IMainTask[] = [];
            try {
                allTasks = await taskService.getAllMainTasks();
            } catch (e) {
                console.warn('Could not load all main tasks', e);
                allTasks = tasks; // Fallback to user's tasks
            }
            setAllMainTasks(allTasks);

            // If ViewTaskID is present and not in allTasks, fetch it specifically
            if (props.initialViewTaskId && !allTasks.some(t => t.Id === props.initialViewTaskId)) {
                try {
                    const viewedTasks = await taskService.getMainTasksByIds([props.initialViewTaskId]);
                    if (viewedTasks.length > 0) {
                        setAllMainTasks(prev => [...prev, ...viewedTasks]);
                    }
                } catch (e) {
                    console.warn('Could not fetch specific viewed task', e);
                }
            }

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
                setUserOptions(uOptions);
            } catch (e) {
                console.warn('Could not load users', e);
            }

            // Try to load subtasks - fallback to empty if fails
            let mySubTasks: ISubTask[] = [];
            try {
                const allSubTasksData = await taskService.getAllSubTasks();
                console.log('[SMTView] All subtasks loaded:', allSubTasksData.length);

                // Filter subtasks to only show those belonging to user's main tasks
                const userMainTaskIds = tasks.map((t: IMainTask) => t.Id);
                mySubTasks = allSubTasksData.filter((s: ISubTask) =>
                    userMainTaskIds.indexOf(s.Admin_Job_ID) !== -1
                );
                console.log('[SMTView] Filtered subtasks for user:', mySubTasks.length);

                // Debug: log subtask Admin_Job_IDs
                if (allSubTasksData.length > 0) {
                    console.log('[SMTView] Sample subtask Admin_Job_IDs:', allSubTasksData.slice(0, 5).map(s => s.Admin_Job_ID));
                }
            } catch (e) {
                console.warn('Could not load subtasks', e);
            }
            setAllSubTasks(mySubTasks);
        } catch (e) {
            console.error('Error loading data', e);
        } finally {
            setLoading(false);
        }
    };

    // Get parent task title by ID
    const getParentTaskTitle = (adminJobId: number): string => {
        const matches = allMainTasks.filter((t: IMainTask) => t.Id === adminJobId);
        const parent = matches.length > 0 ? matches[0] : null;
        return parent ? parent.Title : `Task #${adminJobId}`;
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

    // Check if overdue
    const isOverdue = (item: IMainTask): boolean => {
        if (item.Status === 'Completed' || !item.TaskDueDate) return false;
        return new Date(item.TaskDueDate) < new Date();
    };

    const isSubtaskOverdue = (item: ISubTask): boolean => {
        if (item.TaskStatus === 'Completed' || !item.TaskDueDate) return false;
        return new Date(item.TaskDueDate) < new Date();
    };

    // Unique values for filters
    const getUniqueValues = (arr: string[]): string[] => {
        const seen: { [key: string]: boolean } = {};
        return arr.filter((item) => {
            if (seen[item]) return false;
            seen[item] = true;
            return true;
        });
    };

    const statusList = getUniqueValues(mainTasks.map(t => t.Status || 'Not Started'));
    const statusOptions = [{ key: '', text: 'All Status' }].concat(statusList.map(s => ({ key: s, text: s })));

    const categoryList = getUniqueValues(allSubTasks.map(t => t.Category || 'Unknown').filter(c => c));
    const categoryOptions = [{ key: '', text: 'All Categories' }].concat(categoryList.map(c => ({ key: c, text: c })));

    // Apply filters
    const getFilteredTasks = (): IMainTask[] => {
        let filtered = mainTasks;
        if (statusFilter) {
            filtered = filtered.filter(t => (t.Status || 'Not Started') === statusFilter);
        }
        if (overdueFilter) {
            filtered = filtered.filter(t => isOverdue(t));
        }
        // Sorting
        if (sortedColumn) {
            filtered = [...filtered].sort((a, b) => {
                const aVal = (a as any)[sortedColumn] || '';
                const bVal = (b as any)[sortedColumn] || '';
                if (aVal < bVal) return isSortedDescending ? 1 : -1;
                if (aVal > bVal) return isSortedDescending ? -1 : 1;
                return 0;
            });
        }
        return filtered;
    };

    const filteredTasks = getFilteredTasks();
    const overdueCount = mainTasks.filter(t => isOverdue(t)).length;

    // Subtask filters
    const [subStatusFilter, setSubStatusFilter] = React.useState<string | undefined>(undefined);
    const [subCategoryFilter, setSubCategoryFilter] = React.useState<string | undefined>(undefined);
    const [subMasterFilter, setSubMasterFilter] = React.useState<string | undefined>(undefined);
    const [subAssignedFilter, setSubAssignedFilter] = React.useState<string | undefined>(undefined);
    const [subOverdueFilter, setSubOverdueFilter] = React.useState<boolean>(false);

    // Helper to get assignee name from subtask
    const getSubtaskAssigneeName = (item: ISubTask): string => {
        const assigned = item.TaskAssignedTo;
        if (Array.isArray(assigned) && assigned.length > 0) {
            return assigned.map(u => u.Title).join(', ');
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if (assigned && (assigned as any).Title) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return (assigned as any).Title;
        }
        return 'Unassigned';
    };

    const getFilteredSubTasks = (): ISubTask[] => {
        let filtered = allSubTasks;
        if (subStatusFilter) {
            filtered = filtered.filter(t => (t.TaskStatus || 'Not Started') === subStatusFilter);
        }
        if (subCategoryFilter) {
            filtered = filtered.filter(t => (t.Category || 'Unknown') === subCategoryFilter);
        }
        if (subMasterFilter) {
            filtered = filtered.filter(t => getParentTaskTitle(t.Admin_Job_ID) === subMasterFilter);
        }
        if (subAssignedFilter) {
            filtered = filtered.filter(t => getSubtaskAssigneeName(t) === subAssignedFilter);
        }
        if (subOverdueFilter) {
            filtered = filtered.filter(t => isSubtaskOverdue(t));
        }
        return filtered;
    };

    const filteredSubTasks = getFilteredSubTasks();
    const subOverdueCount = allSubTasks.filter(t => isSubtaskOverdue(t)).length;

    const subStatusList = getUniqueValues(allSubTasks.map(t => t.TaskStatus || 'Not Started'));
    const subStatusOptions = [{ key: '', text: 'All Status' }].concat(subStatusList.map(s => ({ key: s, text: s })));

    const subCategoryList = getUniqueValues(allSubTasks.map(t => t.Category || 'Unknown').filter(c => c));
    const subCategoryOptions = [{ key: '', text: 'All Categories' }].concat(subCategoryList.map(c => ({ key: c, text: c })));

    // Master Task filter options (from user's main tasks)
    const subMasterList = getUniqueValues(allSubTasks.map(t => getParentTaskTitle(t.Admin_Job_ID)));
    const subMasterOptions = [{ key: '', text: 'All Master Tasks' }].concat(subMasterList.map(m => ({ key: m, text: m })));

    // Assigned To filter options
    const subAssignedList = getUniqueValues(allSubTasks.map(t => getSubtaskAssigneeName(t)));
    const subAssignedOptions = [{ key: '', text: 'All Assignees' }].concat(subAssignedList.map(a => ({ key: a, text: a })));

    const resetFilters = () => {
        setStatusFilter(undefined);
        setCategoryFilter(undefined);
        setOverdueFilter(false);
    };

    const resetSubFilters = () => {
        setSubStatusFilter(undefined);
        setSubCategoryFilter(undefined);
        setSubMasterFilter(undefined);
        setSubAssignedFilter(undefined);
        setSubOverdueFilter(false);
    };

    // Export Main Tasks to Excel
    const exportMainTasksToExcel = (): void => {
        try {
            const headers = ['ID', 'Task', 'Description', 'Status', '% Complete', 'Due Date', 'Start Date', 'End Date'];
            const rows = filteredTasks.map(t => {
                const subtaskInfo = getSubtaskInfo(t.Id);
                let pct = 0;
                if (t.Status === 'Completed') {
                    pct = 100;
                } else if (subtaskInfo.total > 0) {
                    pct = calculatePercentComplete(t.Id);
                }
                return [
                    t.Id,
                    t.Title || '',
                    t.Task_x0020_Description || '',
                    t.Status || 'Not Started',
                    pct + '%',
                    formatDate(t.TaskDueDate),
                    formatDate((t as any).TaskStartDate),
                    formatDate((t as any).Task_x0020_End_x0020_Date)
                ].map(v => `"${(v || '').toString().replace(/"/g, '""')}"`).join(',');
            });
            const csvContent = [headers.join(',')].concat(rows).join('\r\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', `MainTasks_${new Date().toISOString().split('T')[0]}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (e) {
            console.error('Export failed:', e);
        }
    };

    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        const newIsSortedDescending = column.key === sortedColumn ? !isSortedDescending : false;
        setSortedColumn(column.key);
        setIsSortedDescending(newIsSortedDescending);
    };

    const getStatusClass = (status: string) => {
        const s = (status || '').replace(/\s+/g, '');
        if (!styles) return '';
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (styles as any)[`status_${s}`] || styles.status_NotStarted;
    };

    // Helper function to calculate % Complete for a main task based on its subtasks
    const calculatePercentComplete = (mainTaskId: number): number => {
        const taskSubTasks = allSubTasks.filter(s => s.Admin_Job_ID === mainTaskId);
        if (taskSubTasks.length === 0) {
            // No subtasks - calculate based on status
            return 0;
        }

        const completedSubTasks = taskSubTasks.filter(s => s.TaskStatus === 'Completed').length;
        return Math.round((completedSubTasks / taskSubTasks.length) * 100);
    };

    // Function to get subtask count info
    const getSubtaskInfo = (mainTaskId: number): { total: number; completed: number } => {
        const taskSubTasks = allSubTasks.filter(s => s.Admin_Job_ID === mainTaskId);
        const completedSubTasks = taskSubTasks.filter(s => s.TaskStatus === 'Completed').length;
        return { total: taskSubTasks.length, completed: completedSubTasks };
    };

    const columns: IColumn[] = [
        {
            key: 'view', name: '', minWidth: 40, maxWidth: 40,
            onRender: (item: IMainTask) => (
                <IconButton
                    iconProps={{ iconName: 'View' }}
                    title="View Details"
                    ariaLabel="View Details"
                    onClick={(e) => { e.stopPropagation(); setSelectedTask(item); }}
                    styles={{ root: { height: 28 } }}
                />
            )
        },
        { key: 'Id', name: '#', fieldName: 'Id', minWidth: 40, maxWidth: 60, isSorted: sortedColumn === 'Id', isSortedDescending, onColumnClick },
        { key: 'Title', name: 'Task', fieldName: 'Title', minWidth: 200, maxWidth: 300, isResizable: true, isSorted: sortedColumn === 'Title', isSortedDescending, onColumnClick },
        { key: 'Task_x0020_Description', name: 'Description', minWidth: 200, maxWidth: 350, isResizable: true, onRender: (i) => (i as any).Task_x0020_Description || '', isSorted: sortedColumn === 'Task_x0020_Description', isSortedDescending, onColumnClick },
        {
            key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, isSorted: sortedColumn === 'Status', isSortedDescending, onColumnClick,
            onRender: (item: IMainTask) => (
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
            key: 'progress', name: '% Complete', minWidth: 150,
            onRender: (item: IMainTask) => {
                const subtaskInfo = getSubtaskInfo(item.Id);
                let pct = 0;

                if (item.Status === 'Completed') {
                    pct = 100;
                } else if (subtaskInfo.total > 0) {
                    // Calculate based on completed subtasks
                    pct = calculatePercentComplete(item.Id);
                } else if (item.Status === 'In Progress') {
                    pct = 50; // Default for In Progress with no subtasks
                }
                // Not Started with no subtasks = 0%

                const progressColor = pct === 100 ? '#107c10' : pct >= 50 ? '#ffbf00' : pct > 0 ? '#00b0ff' : '#e0e0e0';

                return (
                    <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                        <div style={{ flexGrow: 1, maxWidth: 80 }}>
                            <div style={{ backgroundColor: '#e0e0e0', borderRadius: 4, height: 8, overflow: 'hidden' }}>
                                <div style={{ width: `${pct}%`, backgroundColor: progressColor, height: '100%', transition: 'width 0.3s' }} />
                            </div>
                        </div>
                        <span style={{ marginLeft: 8, fontSize: 11, fontWeight: 600, minWidth: 35 }}>{pct}%</span>
                        {subtaskInfo.total > 0 ? (
                            <span style={{ marginLeft: 4, fontSize: 10, color: '#666' }}>({subtaskInfo.completed}/{subtaskInfo.total})</span>
                        ) : (
                            <span style={{ marginLeft: 4, fontSize: 10, color: '#999', fontStyle: 'italic' }}>(No subtasks)</span>
                        )}
                    </div>
                );
            }
        },
        {
            key: 'TaskDueDate', name: 'Due Date', minWidth: 100,
            onRender: (i: IMainTask) => {
                const dateStr = formatDate(i.TaskDueDate);
                const isTaskOverdue = isOverdue(i);
                return (
                    <span style={{ color: isTaskOverdue ? '#a80000' : 'inherit', fontWeight: isTaskOverdue ? 600 : 'normal' }}>
                        {dateStr}
                    </span>
                );
            },
            isSorted: sortedColumn === 'TaskDueDate', isSortedDescending, onColumnClick
        },
        {
            key: 'TaskStartDate', name: 'Start Date', minWidth: 100,
            isSorted: sortedColumn === 'TaskStartDate', isSortedDescending, onColumnClick
        },
        { key: 'Departments', name: 'Department', fieldName: 'Departments', minWidth: 100, isResizable: true, isSorted: sortedColumn === 'Departments', isSortedDescending, onColumnClick },
        { key: 'Project', name: 'Project', fieldName: 'Project', minWidth: 100, isResizable: true, isSorted: sortedColumn === 'Project', isSortedDescending, onColumnClick },
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

    const subColumns: IColumn[] = [
        { key: 'Admin_Job_ID', name: 'Master Task', minWidth: 150, isResizable: true, onRender: (i: ISubTask) => getParentTaskTitle(i.Admin_Job_ID) },
        { key: 'Task_Title', name: 'Subtask', fieldName: 'Task_Title', minWidth: 120, isResizable: true },
        { key: 'Task_Description', name: 'Description', fieldName: 'Task_Description', minWidth: 150, isResizable: true },
        { key: 'TaskAssignedTo', name: 'Assigned To', minWidth: 120, isResizable: true, onRender: (i: ISubTask) => getSubtaskAssigneeName(i) },
        {
            key: 'TaskStatus', name: 'Status', minWidth: 100, onRender: (i: ISubTask) => (
                <span className={`${styles.statusBadge} ${getStatusClass(i.TaskStatus || '')}`}>
                    {i.TaskStatus || 'Not Started'}
                </span>
            )
        },
        { key: 'Category', name: 'Category', fieldName: 'Category', minWidth: 80, isResizable: true },
        { key: 'TaskDueDate', name: 'Due Date', minWidth: 100, onRender: (i: ISubTask) => formatDate(i.TaskDueDate) },
        { key: 'Task_End_Date', name: 'End Date', minWidth: 100, onRender: (i: ISubTask) => formatDate(i.Task_End_Date) },
        { key: 'User_Remarks', name: 'Remarks', minWidth: 100, isResizable: true, onRender: (i: ISubTask) => i.User_Remarks || '' },
    ];

    // Export subtasks to CSV
    const exportSubtasksToExcel = () => {
        try {
            const headers = ['Master Task', 'Subtask', 'Description', 'Assigned To', 'Status', 'Category', 'Due Date', 'End Date', 'Remarks'];
            const rows = filteredSubTasks.map(t => {
                return [
                    getParentTaskTitle(t.Admin_Job_ID),
                    t.Task_Title || '',
                    t.Task_Description || '',
                    getSubtaskAssigneeName(t),
                    t.TaskStatus || '',
                    t.Category || '',
                    formatDate(t.TaskDueDate),
                    formatDate(t.Task_End_Date),
                    t.User_Remarks || ''
                ].map(v => `"${(v || '').toString().replace(/"/g, '""')}"`).join(',');
            });
            const csvContent = [headers.join(',')].concat(rows).join('\r\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', 'MySubtasks.csv');
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (e) {
            console.error(e);
        }
    };

    // Create task form state
    const [isCreating, setIsCreating] = React.useState(false);
    const [newTitle, setNewTitle] = React.useState('');
    const [newDesc, setNewDesc] = React.useState('');
    const [newAssignKey, setNewAssignKey] = React.useState<string | number | undefined>(undefined);
    const [newDueDate, setNewDueDate] = React.useState<Date | undefined>(undefined);
    const [newDept, setNewDept] = React.useState('');
    const [newProject, setNewProject] = React.useState('');
    const [newYear, setNewYear] = React.useState<string>('');
    const [newMonth, setNewMonth] = React.useState<string>('');
    const [msg, setMsg] = React.useState<string | undefined>(undefined);

    const handleCreateMainTask = async () => {
        try {
            if (!newTitle) {
                setMsg('Title is required.');
                return;
            }
            if (!newDueDate) {
                setMsg('Due Date is required.');
                return;
            }
            if (!newAssignKey) {
                setMsg('Assigned To is required.');
                return;
            }
            if (!newDept) {
                setMsg('Department is required.');
                return;
            }
            if (!newProject) {
                setMsg('Project is required.');
                return;
            }
            if (!newYear) {
                setMsg('Year is required.');
                return;
            }
            if (!newMonth) {
                setMsg('Month is required.');
                return;
            }

            await taskService.createMainTask({
                Title: newTitle,
                Task_x0020_Description: newDesc,
                TaskDueDate: newDueDate.toISOString(),
                Status: 'Not Started',
                TaskAssignedToId: Number(newAssignKey),
                Departments: newDept,
                Project: newProject,
                SMTYear: newYear,
                SMTMonth: newMonth
            } as any);
            setIsCreating(false);
            setNewTitle(''); setNewDesc(''); setNewAssignKey(undefined); setNewDueDate(undefined); setNewDept(''); setNewProject(''); setNewYear(''); setNewMonth('');
            loadData().catch(console.error);
        } catch (e: any) {
            setMsg(e.message || e);
        }
    };

    if (loading) return <div>Loading...</div>;

    return (
        <div>
            <div style={{ marginBottom: 10 }}>
                <h3 style={{ margin: 0 }}>My Tasks</h3>
            </div>

            <Pivot>
                {/* Main Tasks Tab */}
                <PivotItem headerText="Main Tasks">
                    {/* Dashboard Badges */}
                    <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 10, marginTop: 10 }}>
                        <span style={{ padding: '4px 12px', backgroundColor: '#0078d4', color: 'white', borderRadius: 4 }}>
                            Total: {mainTasks.length}
                        </span>
                        {statusList.map(status => {
                            const count = mainTasks.filter(t => (t.Status || 'Not Started') === status).length;
                            const colors: { [key: string]: string } = {
                                'Not Started': '#6c757d',
                                'In Progress': '#0078d4',
                                'Completed': '#107c10',
                                'On Hold': '#ff8c00'
                            };
                            return (
                                <span
                                    key={status}
                                    style={{ padding: '4px 12px', backgroundColor: colors[status] || '#6c757d', color: 'white', borderRadius: 4, cursor: 'pointer', opacity: statusFilter === status ? 1 : 0.8 }}
                                    onClick={() => setStatusFilter(statusFilter === status ? undefined : status)}
                                >
                                    {status}: {count}
                                </span>
                            );
                        })}
                        {overdueCount > 0 && (
                            <span style={{ padding: '4px 12px', backgroundColor: '#a80000', color: 'white', borderRadius: 4, cursor: 'pointer', opacity: overdueFilter ? 1 : 0.8 }} onClick={() => setOverdueFilter(!overdueFilter)}>
                                Overdue: {overdueCount}
                            </span>
                        )}
                    </div>

                    {/* Filters */}
                    <Stack horizontal tokens={{ childrenGap: 10 }} wrap styles={{ root: { marginBottom: 10 } }} verticalAlign="end">
                        <Dropdown
                            label="Status"
                            placeholder="All Status"
                            selectedKey={statusFilter || ''}
                            onChange={(_, opt) => setStatusFilter(opt?.key as string || undefined)}
                            options={statusOptions}
                            styles={{ root: { width: 140 } }}
                        />
                        <Checkbox
                            label="Overdue Only"
                            checked={overdueFilter}
                            onChange={(_, v) => setOverdueFilter(!!v)}
                            styles={{ root: { marginTop: 6 } }}
                        />
                        <DefaultButton iconProps={{ iconName: 'Clear' }} text="Reset" onClick={resetFilters} />
                        <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export to Excel" onClick={exportMainTasksToExcel} />
                        <PrimaryButton iconProps={{ iconName: 'Add' }} text="Add New Task" onClick={() => setIsCreating(true)} />
                    </Stack>

                    {/* Grid */}
                    {/* Grid */}
                    <div style={{ position: 'relative', height: '70vh', border: '1px solid #edebe9', borderRadius: 4, marginTop: 10, minWidth: 0, overflowX: 'auto' }}>
                        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                            <DetailsList
                                items={filteredTasks}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                compact={true}
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

                {/* Child Tasks Tab */}
                <PivotItem headerText="My Subtasks">
                    {/* Dashboard Badges - Show both total and filtered */}
                    <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 10, marginTop: 10, alignItems: 'center' }}>
                        <span style={{ padding: '4px 12px', backgroundColor: '#323130', color: 'white', borderRadius: 4 }}>
                            Showing: {filteredSubTasks.length} / {allSubTasks.length}
                        </span>
                        {['Not Started', 'In Progress', 'Completed', 'On Hold'].map(status => {
                            const totalCount = allSubTasks.filter(t => (t.TaskStatus || 'Not Started') === status).length;
                            const filteredCount = filteredSubTasks.filter(t => (t.TaskStatus || 'Not Started') === status).length;
                            const colors: { [key: string]: string } = {
                                'Not Started': '#6c757d',
                                'In Progress': '#0078d4',
                                'Completed': '#107c10',
                                'On Hold': '#ff8c00'
                            };
                            const isActive = subStatusFilter === status;
                            return (
                                <span
                                    key={status}
                                    style={{
                                        padding: '4px 12px',
                                        backgroundColor: isActive ? colors[status] : colors[status],
                                        color: 'white',
                                        borderRadius: 4,
                                        cursor: 'pointer',
                                        opacity: isActive ? 1 : 0.7,
                                        border: isActive ? '2px solid #000' : '2px solid transparent',
                                        fontWeight: isActive ? 'bold' : 'normal',
                                        transition: 'all 0.2s'
                                    }}
                                    onClick={() => setSubStatusFilter(isActive ? undefined : status)}
                                    title={`Click to ${isActive ? 'clear' : 'filter by'} ${status}`}
                                >
                                    {status}: {filteredCount} / {totalCount}
                                </span>
                            );
                        })}
                        <span
                            style={{
                                padding: '4px 12px',
                                backgroundColor: '#a80000',
                                color: 'white',
                                borderRadius: 4,
                                cursor: 'pointer',
                                opacity: subOverdueFilter ? 1 : 0.7,
                                border: subOverdueFilter ? '2px solid #000' : '2px solid transparent',
                                fontWeight: subOverdueFilter ? 'bold' : 'normal',
                                transition: 'all 0.2s'
                            }}
                            onClick={() => setSubOverdueFilter(!subOverdueFilter)}
                            title={`Click to ${subOverdueFilter ? 'clear' : 'filter by'} Overdue`}
                        >
                            Overdue: {filteredSubTasks.filter(t => isSubtaskOverdue(t)).length} / {subOverdueCount}
                        </span>
                    </div>

                    {/* Filters */}
                    <Stack horizontal tokens={{ childrenGap: 10 }} wrap styles={{ root: { marginBottom: 10 } }} verticalAlign="end">
                        <Dropdown
                            label="Master Task"
                            placeholder="All Master Tasks"
                            selectedKey={subMasterFilter || ''}
                            onChange={(_, opt) => setSubMasterFilter(opt?.key as string || undefined)}
                            options={subMasterOptions}
                            styles={{ root: { width: 180 } }}
                        />
                        <Dropdown
                            label="Assigned To"
                            placeholder="All Assignees"
                            selectedKey={subAssignedFilter || ''}
                            onChange={(_, opt) => setSubAssignedFilter(opt?.key as string || undefined)}
                            options={subAssignedOptions}
                            styles={{ root: { width: 140 } }}
                        />
                        <Dropdown
                            label="Status"
                            placeholder="All Status"
                            selectedKey={subStatusFilter || ''}
                            onChange={(_, opt) => setSubStatusFilter(opt?.key as string || undefined)}
                            options={subStatusOptions}
                            styles={{ root: { width: 140 } }}
                        />
                        <Dropdown
                            label="Category"
                            placeholder="All Categories"
                            selectedKey={subCategoryFilter || ''}
                            onChange={(_, opt) => setSubCategoryFilter(opt?.key as string || undefined)}
                            options={subCategoryOptions}
                            styles={{ root: { width: 140 } }}
                        />
                        <Checkbox
                            label="Overdue Only"
                            checked={subOverdueFilter}
                            onChange={(_, v) => setSubOverdueFilter(!!v)}
                            styles={{ root: { marginTop: 6 } }}
                        />
                        <DefaultButton iconProps={{ iconName: 'Clear' }} text="Reset" onClick={resetSubFilters} />
                        <PrimaryButton iconProps={{ iconName: 'ExcelDocument' }} text="Export to Excel" onClick={exportSubtasksToExcel} />
                    </Stack>

                    {/* Grid */}
                    {/* Grid */}
                    <div style={{ position: 'relative', height: '70vh', border: '1px solid #edebe9', borderRadius: 4, marginTop: 10, minWidth: 0, overflowX: 'auto' }}>
                        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                            <DetailsList
                                items={filteredSubTasks}
                                columns={subColumns}
                                selectionMode={SelectionMode.single}
                                compact={true}
                                onActiveItemChanged={(item: ISubTask) => item && setSelectedSubtask(item)}
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
                    {filteredSubTasks.length === 0 && <div style={{ padding: 20, textAlign: 'center', color: '#666' }}>No subtasks found.</div>}
                </PivotItem>
            </Pivot>

            {/* Subtask Detail Panel */}
            <Panel
                isOpen={!!selectedSubtask}
                onDismiss={() => setSelectedSubtask(null)}
                type={PanelType.medium}
                headerText="Subtask Details"
            >
                {selectedSubtask && (
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { padding: 10 } }}>
                        <div style={{ marginBottom: 10 }}>
                            <span style={{
                                padding: '4px 12px',
                                backgroundColor: selectedSubtask.TaskStatus === 'Completed' ? '#107c10' :
                                    selectedSubtask.TaskStatus === 'In Progress' ? '#0078d4' :
                                        selectedSubtask.TaskStatus === 'On Hold' ? '#ff8c00' : '#6c757d',
                                color: 'white',
                                borderRadius: 4,
                                fontSize: 12
                            }}>
                                {selectedSubtask.TaskStatus || 'Not Started'}
                            </span>
                        </div>

                        <div>
                            <strong>Master Task:</strong> {getParentTaskTitle(selectedSubtask.Admin_Job_ID)}
                        </div>
                        <div>
                            <strong>Subtask Title:</strong> {selectedSubtask.Task_Title}
                        </div>
                        <div>
                            <strong>Description:</strong> {selectedSubtask.Task_Description || 'N/A'}
                        </div>
                        <div>
                            <strong>Assigned To:</strong> {getSubtaskAssigneeName(selectedSubtask)}
                        </div>
                        <div>
                            <strong>Category:</strong> {selectedSubtask.Category || 'N/A'}
                        </div>
                        <div>
                            <strong>Due Date:</strong> {formatDate(selectedSubtask.TaskDueDate) || 'N/A'}
                        </div>
                        <div>
                            <strong>End Date:</strong> {formatDate(selectedSubtask.Task_End_Date) || 'N/A'}
                        </div>
                        <div>
                            <strong>User Remarks:</strong>
                            <div style={{ marginTop: 5, padding: 10, backgroundColor: '#f3f2f1', borderRadius: 4 }}>
                                {selectedSubtask.User_Remarks || 'No remarks'}
                            </div>
                        </div>

                        {/* Attachments */}
                        <div>
                            <strong>Attachments:</strong>
                            {(() => {
                                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                const files = (selectedSubtask as any).AttachmentFiles;
                                if (!files || files.length === 0) return <div style={{ color: '#999', marginTop: 5 }}>No attachments</div>;
                                return (
                                    <ul style={{ margin: '5px 0', paddingLeft: 20 }}>
                                        {/* eslint-disable-next-line @typescript-eslint/no-explicit-any */}
                                        {files.map((f: any, idx: number) => (
                                            <li key={idx}>
                                                <a href={f.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                                                    {f.FileName}
                                                </a>
                                            </li>
                                        ))}
                                    </ul>
                                );
                            })()}
                        </div>

                        <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                            <DefaultButton text="Close" onClick={() => setSelectedSubtask(null)} />
                        </Stack>
                    </Stack>
                )}
            </Panel>

            {/* Task Detail Panel */}
            <Panel
                isOpen={!!selectedTask}
                onDismiss={() => {
                    setSelectedTask(null);
                    processedMainIdRef.current = undefined; // Reset ref so it can re-open if clicked again
                }}
                type={PanelType.extraLarge}
                headerText=""
            >
                {selectedTask && (
                    <TaskDetail
                        mainTask={selectedTask}
                        initialChildTaskId={props.initialChildTaskId}
                        initialTab={props.initialTab}
                        onDeepLinkProcessed={props.onChildDeepLinkProcessed}
                    />
                )}
            </Panel>

            {/* Create Task Panel */}
            <Panel
                isOpen={isCreating}
                onDismiss={() => setIsCreating(false)}
                headerText="Create New Main Task"
                type={PanelType.large}
            >
                <Stack tokens={{ childrenGap: 15 }}>
                    <TextField label="Task Title" value={newTitle} onChange={(e, v) => setNewTitle(v || '')} required />
                    <TextField label="Description" multiline rows={3} value={newDesc} onChange={(e, v) => setNewDesc(v || '')} />

                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <Dropdown
                            label="Year"
                            options={[{ key: '2024', text: '2024' }, { key: '2025', text: '2025' }, { key: '2026', text: '2026' }]}
                            selectedKey={newYear}
                            onChange={(_, opt) => setNewYear(opt?.key as string || '')}
                            styles={{ root: { width: '50%' } }}
                        />
                        <Dropdown
                            label="Month"
                            options={['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'].map(m => ({ key: m, text: m }))}
                            selectedKey={newMonth}
                            onChange={(_, opt) => setNewMonth(opt?.key as string || '')}
                            styles={{ root: { width: '50%' } }}
                        />
                    </Stack>

                    <ComboBox
                        label="Assign To"
                        options={userOptions}
                        selectedKey={newAssignKey}
                        onChange={(_, opt) => setNewAssignKey(opt?.key as string | number | undefined)}
                        allowFreeform={true}
                        autoComplete="on"
                        useComboBoxAsMenuWidth
                    />

                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <ComboBox
                            label="Departments"
                            options={departmentOptions}
                            selectedKey={newDept}
                            onChange={(_, opt) => setNewDept(opt?.key as string || opt?.text || '')} // Allow freeform for new depts
                            allowFreeform={true}
                            styles={{ root: { width: '50%' } }}
                        />
                        <TextField
                            label="Project"
                            value={newProject}
                            onChange={(e, v) => setNewProject(v || '')}
                            styles={{ root: { width: '50%' } }}
                        />
                    </Stack>

                    <DatePicker label="Due Date" value={newDueDate} onSelectDate={(d) => setNewDueDate(d || undefined)} />

                    {msg && <MessageBar messageBarType={MessageBarType.error}>{msg}</MessageBar>}
                    <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                        <PrimaryButton text="Create Task" onClick={handleCreateMainTask} />
                        <DefaultButton text="Cancel" onClick={() => setIsCreating(false)} />
                    </Stack>
                </Stack>
            </Panel>
        </div>
    );
};
