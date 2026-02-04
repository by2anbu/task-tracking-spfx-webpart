import * as React from 'react';
import { DetailsList, SelectionMode, TextField, PrimaryButton, DefaultButton, DatePicker, Stack, Separator, MessageBar, MessageBarType, ProgressIndicator, Dropdown, Modal, IDropdownOption, Persona, PersonaSize, ComboBox, IComboBox, IComboBoxOption, SelectableOptionMenuItemType, Pivot, PivotItem, PivotLinkFormat, PivotLinkSize, ActivityItem, Icon, Dialog, DialogType, DialogFooter, IconButton, ScrollablePane, ScrollbarVisibility, Sticky, StickyPositionType, ConstrainMode, DetailsListLayoutMode, IDetailsHeaderProps, TooltipHost } from 'office-ui-fabric-react';
import { taskService } from '../../../../services/sp-service';
import { IMainTask, ISubTask, LIST_MAIN_TASKS, LIST_SUB_TASKS } from '../../../../services/interfaces';
import { WorkflowDesigner } from './WorkflowDesigner';
import styles from './TaskDetail.module.scss';
import { sanitizeHtml } from '../../../../utils/sanitize';

export interface ITaskDetailProps {
    mainTask: IMainTask;
    initialChildTaskId?: number;
    initialTab?: string;
    onDeepLinkProcessed?: () => void;
    onTaskUpdated?: () => void; // Callback to refresh parent data after completing task
}

// Extended Interface for UI
interface IUiSubTask extends ISubTask {
    depth: number;
    hasChildren: boolean;
    isExpanded?: boolean; // For future collapsible support
}

export const TaskDetail: React.FunctionComponent<ITaskDetailProps> = ({ mainTask, initialChildTaskId, initialTab, onDeepLinkProcessed, onTaskUpdated }) => {
    const [subTasks, setSubTasks] = React.useState<ISubTask[]>([]);
    const [uiSubTasks, setUiSubTasks] = React.useState<IUiSubTask[]>([]); // Processed list
    const [isAdding, setIsAdding] = React.useState<boolean>(false);
    const [addingToParentId, setAddingToParentId] = React.useState<number | undefined>(undefined); // Track if adding to a subtask
    const [selectedSubtask, setSelectedSubtask] = React.useState<ISubTask | null>(null);
    const [assignedFilter, setAssignedFilter] = React.useState<string | undefined>(undefined);
    const [categoryFilter, setCategoryFilter] = React.useState<string | undefined>(undefined);
    const [sortedColumn, setSortedColumn] = React.useState<string | undefined>(undefined);
    const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

    // Form State
    const [newTitle, setNewTitle] = React.useState<string>('');
    const [newDesc, setNewDesc] = React.useState<string>('');

    // ComboBox Selection
    const [selectedUserKey, setSelectedUserKey] = React.useState<string | number | undefined>(undefined);
    const [userOptions, setUserOptions] = React.useState<IComboBoxOption[]>([]);

    const [newCategory, setNewCategory] = React.useState<string>('');
    const [categoryFormOptions, setCategoryFormOptions] = React.useState<IComboBoxOption[]>([]);
    const [dueDate, setDueDate] = React.useState<Date | undefined>(new Date());
    const [message, setMessage] = React.useState<string | undefined>(undefined);

    // Attachments
    const [attachments, setAttachments] = React.useState<any[]>([]);

    // Edit Panel State
    const [editStatus, setEditStatus] = React.useState<string>('');
    const [editRemarks, setEditRemarks] = React.useState<string>('');
    const [editDueDate, setEditDueDate] = React.useState<Date | undefined>(undefined);
    const [editAssigneeId, setEditAssigneeId] = React.useState<number | undefined>(undefined);
    const [editFiles, setEditFiles] = React.useState<FileList | null>(null);

    // Subtask Change Reason
    const [changePrompt, setChangePrompt] = React.useState<{ visible: boolean; field: 'DueDate' | 'Assignee'; newValue: any } | null>(null);
    const [changeRemark, setChangeRemark] = React.useState('');

    // Email / Correspondence State
    const [correspondence, setCorrespondence] = React.useState<any[]>([]);
    const [emailSubject, setEmailSubject] = React.useState<string>('');
    const [emailBody, setEmailBody] = React.useState<string>('');
    const [emailToKey, setEmailToKey] = React.useState<string | number | undefined>(undefined);
    const [emailFile, setEmailFile] = React.useState<File | undefined>(undefined);
    const [isSending, setIsSending] = React.useState<boolean>(false);
    const [subtaskCorrespondence, setSubtaskCorrespondence] = React.useState<any[]>([]);
    const [subtaskAttachments, setSubtaskAttachments] = React.useState<any[]>([]);
    const [isSubtaskSending, setIsSubtaskSending] = React.useState<boolean>(false);

    // Query to Assignee State
    const [showQueryDialog, setShowQueryDialog] = React.useState<boolean>(false);
    const [queryMessage, setQueryMessage] = React.useState<string>('');
    const [isSendingQuery, setIsSendingQuery] = React.useState<boolean>(false);

    // Complete Dialog State
    const [showCompleteDialog, setShowCompleteDialog] = React.useState<boolean>(false);
    const [completeRemarks, setCompleteRemarks] = React.useState<string>('');
    const [isCompleting, setIsCompleting] = React.useState<boolean>(false);

    // Track if the main task is completed (to hide buttons), end date, and user remarks
    const [currentMainTaskStatus, setCurrentMainTaskStatus] = React.useState<string>(mainTask.Status || 'Not Started');
    const [currentEndDate, setCurrentEndDate] = React.useState<string | undefined>(mainTask.Task_x0020_End_x0020_Date);
    const [currentUserRemarks, setCurrentUserRemarks] = React.useState<string>(mainTask.UserRemarks || '');

    // Sync status, end date, and user remarks with mainTask prop changes
    React.useEffect(() => {
        setCurrentMainTaskStatus(mainTask.Status || 'Not Started');
        setCurrentEndDate(mainTask.Task_x0020_End_x0020_Date);
        setCurrentUserRemarks(mainTask.UserRemarks || '');
    }, [mainTask.Status, mainTask.Task_x0020_End_x0020_Date, mainTask.UserRemarks]);

    const [subtaskInitialTab, setSubtaskInitialTab] = React.useState<string>('Details');

    // Ref to track if we have already handled this deep link to prevent re-opening on manual clicks
    const processedChildTaskIdRef = React.useRef<number | undefined>(undefined);

    // Effect to handle deep linking for subtask
    React.useEffect(() => {
        // Only proceed if we have an ID, we haven't processed it yet (or it changed), and data is loaded
        if (initialChildTaskId && subTasks.length > 0 && processedChildTaskIdRef.current !== initialChildTaskId) {

            // Check if this subtask exists in the current list
            const targetSubtask = subTasks.filter(t => t.Id === initialChildTaskId)[0];

            if (targetSubtask) {
                console.log('[TaskDetail] Deep linking to subtask:', initialChildTaskId);

                // Determine tab
                const targetTab = (initialTab === 'Correspondence') ? 'Email History' : 'Details';
                handleSubtaskClick(targetSubtask, targetTab);

                // Mark as processed so we don't open it again if the user clicks the main task row manually
                processedChildTaskIdRef.current = initialChildTaskId;

                // Notify parent to clear state (so props are reset)
                if (onDeepLinkProcessed) {
                    onDeepLinkProcessed();
                }
            }
        }
    }, [initialChildTaskId, subTasks, initialTab]);

    // ... (rest of code)

    const handleSubtaskClick = (item: ISubTask, targetTabOrIndex?: string | number): void => {
        let targetTab = typeof targetTabOrIndex === 'string' ? targetTabOrIndex : 'Activity';

        // Map legacy or ambiguous tabs
        if (targetTab === 'Details' || targetTab === 'Update') targetTab = 'Update';
        if (targetTab === 'Correspondence' || targetTab === 'Activity') targetTab = 'Activity';

        setSubtaskInitialTab(targetTab);

        setSelectedSubtask(item);
        setEditStatus(item.TaskStatus || 'Not Started');
        setEditRemarks(item.User_Remarks || '');
        setEditDueDate(item.TaskDueDate ? new Date(item.TaskDueDate) : undefined);
        setEditAssigneeId(item.TaskAssignedTo?.Id || item.TaskAssignedToId);
        setEditFiles(null);

        // Reset email form
        setEmailSubject(`Regarding: ${item.Task_Title}`);
        setEmailBody('');
        setEmailToKey(undefined);
        setEmailFile(undefined);
        setSubtaskCorrespondence([]);
        setSubtaskAttachments([]);
        loadSubtaskData(item.Id);
    };

    const loadSubtaskData = async (subTaskId: number) => {
        try {
            const [history, files] = await Promise.all([
                taskService.getTaskCorrespondence(mainTask.Id, subTaskId),
                taskService.getAttachments(LIST_SUB_TASKS, subTaskId)
            ]);
            setSubtaskCorrespondence(history);
            setSubtaskAttachments(files);
        } catch (e) {
            console.warn("Could not load subtask data", e);
        }
    };

    // Effect to update status display when subtasks change (to reflect "In Progress" if subtasks added)
    React.useEffect(() => {
        if (subTasks.length > 0 && currentMainTaskStatus === 'Not Started') {
            // If we have subtasks but status is still "Not Started", update to "In Progress"
            setCurrentMainTaskStatus('In Progress');
        }
    }, [subTasks.length]);

    React.useEffect(() => {
        loadData().catch(console.error);
    }, [mainTask]);

    // Expand/Collapse State
    const [expandedRows, setExpandedRows] = React.useState<Set<number>>(new Set());

    const toggleExpand = (itemId: number): void => {
        const newExpanded = new Set<number>();
        expandedRows.forEach(item => newExpanded.add(item));

        if (newExpanded.has(itemId)) {
            newExpanded.delete(itemId);
        } else {
            newExpanded.add(itemId);
        }
        setExpandedRows(newExpanded);
    };

    const processHierarchy = (items: ISubTask[]): IUiSubTask[] => {
        const itemMap = new Map<number, ISubTask>();
        const childrenMap = new Map<number, number[]>();

        items.forEach(item => {
            itemMap.set(item.Id, item);
            const pId = item.ParentSubtaskId || 0;
            if (!childrenMap.has(pId)) {
                childrenMap.set(pId, []);
            }
            childrenMap.get(pId)!.push(item.Id);
        });

        const result: IUiSubTask[] = [];

        const traverse = (parentId: number, depth: number) => {
            const children = childrenMap.get(parentId) || [];
            // Sort children by ID or Created date? Original was flat sort. 
            // We'll keep them in order of appearance in original list if possible or just ID
            // Simple sort by ID for stability
            children.sort((a, b) => a - b);

            children.forEach(childId => {
                const child = itemMap.get(childId);
                if (child) {
                    result.push({
                        ...child,
                        depth: depth,
                        hasChildren: childrenMap.has(child.Id) && childrenMap.get(child.Id)!.length > 0
                    });
                    traverse(child.Id, depth + 1);
                }
            });
        };

        traverse(0, 0); // Start with top level (ParentSubtaskId = 0 or undefined)

        // If filters are active, we might need to fallback to flat list or handle differently
        // For now, if filters are active, hierarchy might be confusing. 
        // We will Apply filters AFTER hierarchy flattening? No, if we filter, we break hierarchy.
        // Strategy: If filter active, show flat list. If no filter, show hierarchy.
        return result;
    };

    const loadData = async (): Promise<void> => {
        const data = await taskService.getSubTasksForMainTask(mainTask.Id);
        setSubTasks(data);
        // Initial hierarchy processing
        setUiSubTasks(processHierarchy(data));

        try {
            const files = await taskService.getAttachments(LIST_MAIN_TASKS, mainTask.Id);
            setAttachments(files);

            // Load users for ComboBox
            const users = await taskService.getSiteUsers();
            const options: IComboBoxOption[] = users.map(u => ({
                key: u.Id,
                text: u.Title,
                data: { email: u.Email }
            }));
            setUserOptions(options);

            // Load Categories
            try {
                const cats = await taskService.getChoiceFieldOptions('Task Tracking System User', 'Category');
                setCategoryFormOptions(cats.map(c => ({ key: c, text: c })));
            } catch (e) {
                console.warn('Could not load categories', e);
            }

            // Load Correspondence for Main Task initially
            await loadCorrespondence(undefined);

        } catch (e) {
            console.error("Error loading data", e);
        }
    };

    const loadCorrespondence = async (subTaskId?: number) => {
        try {
            const history = await taskService.getTaskCorrespondence(mainTask.Id, subTaskId);
            if (subTaskId) {
                setSubtaskCorrespondence(history);
            } else {
                setCorrespondence(history);
            }
        } catch (e) {
            console.warn("Could not load correspondence", e);
        }
    };

    const handleSubtaskCorrespondence = async (): Promise<void> => {
        if (!selectedSubtask || !emailBody) return;
        setIsSubtaskSending(true);
        try {
            await taskService.createTaskCorrespondence({
                parentTaskId: mainTask.Id,
                childTaskId: selectedSubtask.Id,
                subject: `Subtask Remark: ${selectedSubtask.Task_Title}`,
                messageBody: emailBody,
                toAddress: (selectedSubtask.TaskAssignedTo as any)?.EMail || '',
                fromAddress: await taskService.getCurrentUserEmail()
            });
            setEmailBody('');
            await loadSubtaskData(selectedSubtask.Id);
            setMessage("Remark posted successfully!");
        } catch (e: any) {
            setMessage("Error posting remark: " + e.message);
        } finally {
            setIsSubtaskSending(false);
        }
    };

    const handleSendQueryToAssignee = async (): Promise<void> => {
        if (!selectedSubtask || !queryMessage.trim()) {
            setMessage("Please enter a query message.");
            return;
        }

        const assigneeEmail = (selectedSubtask.TaskAssignedTo as any)?.EMail;
        if (!assigneeEmail) {
            setMessage("No assignee found for this task.");
            return;
        }

        setIsSendingQuery(true);
        try {
            await taskService.createTaskCorrespondence({
                parentTaskId: mainTask.Id,
                childTaskId: selectedSubtask.Id,
                subject: `Query: ${selectedSubtask.Task_Title}`,
                messageBody: `📩 Query from ${await taskService.getCurrentUserEmail()}:\n\n${queryMessage}`,
                toAddress: assigneeEmail,
                fromAddress: await taskService.getCurrentUserEmail()
            });
            setQueryMessage('');
            setShowQueryDialog(false);
            await loadSubtaskData(selectedSubtask.Id);
            setMessage("Query sent successfully to assignee!");
        } catch (e: any) {
            setMessage("Error sending query: " + e.message);
        } finally {
            setIsSendingQuery(false);
        }
    };

    const getProgressStats = (): { status: string; percent: number } => {
        // If task is marked as completed, show 100%
        if (currentMainTaskStatus === 'Completed') {
            return { status: 'Completed', percent: 100 };
        }
        if (!subTasks || subTasks.length === 0) {
            return { status: currentMainTaskStatus || 'Not Started', percent: 0 };
        }
        const total = subTasks.length;
        const completed = subTasks.filter(t => t.TaskStatus === 'Completed').length;
        if (completed === total) {
            return { status: 'Completed', percent: 100 };
        }
        const percent = Math.round((completed / total) * 100);
        return { status: 'In Progress', percent: percent };
    };

    // Handle completing a main task (for tasks without subtasks)
    const handleCompleteTask = async (): Promise<void> => {
        setIsCompleting(true);
        try {
            await taskService.updateMainTaskStatus(mainTask.Id, 'Completed', completeRemarks);
            setCurrentMainTaskStatus('Completed');
            setCurrentEndDate(new Date().toISOString()); // Set end date to now
            setCurrentUserRemarks(completeRemarks); // Update remarks in header
            setShowCompleteDialog(false);
            setCompleteRemarks('');
            setMessage('Task completed successfully!');
            // Notify parent to refresh data
            if (onTaskUpdated) {
                onTaskUpdated();
            }
        } catch (err: any) {
            setMessage('Error completing task: ' + (err.message || err));
        } finally {
            setIsCompleting(false);
        }
    };

    const stats = getProgressStats();

    const calculateWeightage = (): string => {
        if (subTasks.length === 0) return "0%";
        return (100 / subTasks.length).toFixed(1) + "%";
    };

    const getAssigneeName = (item: ISubTask): string => {
        const assignee = item.TaskAssignedTo;
        if (Array.isArray(assignee) && assignee.length > 0) return assignee[0].Title;
        if (assignee && (assignee as any).Title) return (assignee as any).Title;
        return 'Unassigned';
    };

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

    const getFilteredTasks = (): IUiSubTask[] => {
        // If sorting or filtering is active, we might not want hierarchy.
        // Assuming we want to search broadly.
        let filtered: ISubTask[] = subTasks;

        if (assignedFilter) {
            filtered = filtered.filter(t => getAssigneeName(t) === assignedFilter);
        }
        if (categoryFilter) {
            filtered = filtered.filter(t => (t.Category || 'Unknown') === categoryFilter);
        }

        // If data is filtered, hierarchies are broken. Just show flat.
        // But if NO filters, we show hierarchy.
        // But if NO filters, we show hierarchy.
        if (!assignedFilter && !categoryFilter && !sortedColumn) {
            // Apply Expansion Logic
            const visibleItems: IUiSubTask[] = [];
            const visibleParents = new Set<number>();
            visibleParents.add(0); // Root is always visible (items with ParentSubtaskId=0)

            // Our processHierarchy returns a flat list in DFS order. 
            // We can iterate and decide to show based on parent visibility.

            for (const item of uiSubTasks) {
                const parentId = item.ParentSubtaskId || 0;

                // If parent is visible and expanded, then this child is potentially visible.
                // Wait, logic is: IF parent is visible AND parent is expanded, THEN child is visible.
                // But "visibleParents" tracks nodes that are effectively shown.

                if (visibleParents.has(parentId)) {
                    visibleItems.push(item);
                    // If this item is expanded, its children can be seen
                    if (expandedRows.has(item.Id)) {
                        visibleParents.add(item.Id);
                    }
                }
            }
            return visibleItems;
        }

        // If we have filters/sort, return flat mapped to UI task with depth 0
        let result = filtered.map(t => ({ ...t, depth: 0, hasChildren: false }));

        if (sortedColumn) {
            result = [...result].sort((a, b) => {
                const aVal = (a as any)[sortedColumn] || '';
                const bVal = (b as any)[sortedColumn] || '';
                if (aVal < bVal) return isSortedDescending ? 1 : -1;
                if (aVal > bVal) return isSortedDescending ? -1 : 1;
                return 0;
            });
        }
        return result;
    };

    const filteredTasks = getFilteredTasks();
    const uniqueAssignees = subTasks.map(t => getAssigneeName(t)).filter((v, i, a) => a.indexOf(v) === i).map(u => ({ key: u, text: u }));
    const uniqueCategories = subTasks.map(t => t.Category || 'Unknown').filter((v, i, a) => a.indexOf(v) === i).map(c => ({ key: c, text: c }));

    const categoryOptions: IDropdownOption[] = [
        { key: 'Development', text: 'Development' },
        { key: 'Testing', text: 'Testing' },
        { key: 'Documentation', text: 'Documentation' },
        { key: 'Design', text: 'Design' },
        { key: 'Support', text: 'Support' },
        { key: 'Requirement Analysis', text: 'Requirement Analysis' },
        { key: 'Deployment', text: 'Deployment' }
    ];

    // --- ComboBox Rendering ---
    const onRenderUserOption = (option: IComboBoxOption): JSX.Element => {
        if (option.itemType === SelectableOptionMenuItemType.Header || option.itemType === SelectableOptionMenuItemType.Divider) {
            return <span>{option.text}</span>;
        }
        return (
            <div style={{ display: 'flex', alignItems: 'center' }}>
                <Persona
                    size={PersonaSize.size24}
                    text={option.text}
                    secondaryText={option.data?.email}
                    imageInitials={option.text.charAt(0)}
                    styles={{ root: { marginRight: 8 } }}
                />
            </div>
        );
    };

    const onUserComboChange = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string): void => {
        if (option) {
            setSelectedUserKey(option.key);
        } else if (value) {
            setSelectedUserKey(undefined);
        } else {
            setSelectedUserKey(undefined);
        }
    };

    const onEmailToComboChange = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string): void => {
        if (option) {
            setEmailToKey(option.key);
        } else {
            setEmailToKey(undefined);
        }
    };
    // ---------------------------

    const exportToExcel = (): void => {
        try {
            const headers = ['Title', 'Description', 'Assigned To', 'Status', 'Due Date', 'End Date', 'Category', 'Remarks'];
            const rows = filteredTasks.map(t => {
                return [
                    t.Task_Title || '',
                    t.Task_Description || '',
                    getAssigneeName(t),
                    t.TaskStatus || '',
                    formatDate(t.TaskDueDate),
                    formatDate(t.Task_End_Date),
                    t.Category || '',
                    t.User_Remarks || ''
                ].map(v => `"${(v || '').toString().replace(/"/g, '""')}"`).join(',');
            });
            const csvContent = [headers.join(',')].concat(rows).join('\r\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', `Subtasks_${mainTask.Id}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (e) {
            console.error(e);
            setMessage('Export failed: ' + ((e as any).message || e));
        }
    };

    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: any): void => {
        const newIsSortedDescending = column.key === sortedColumn ? !isSortedDescending : false;
        setSortedColumn(column.key);
        setIsSortedDescending(newIsSortedDescending);
    };


    const handleSaveSubtask = async (): Promise<void> => {
        if (!selectedSubtask) return;
        try {
            setMessage(undefined);
            await taskService.updateSubTaskStatus(selectedSubtask.Id, mainTask.Id, editStatus, editRemarks, false, editDueDate?.toISOString());
            if (editFiles && editFiles.length > 0) {
                const fileArray: File[] = [];
                for (let i = 0; i < editFiles.length; i++) {
                    fileArray.push(editFiles[i]);
                }
                await taskService.addAttachmentsToItem(LIST_SUB_TASKS, selectedSubtask.Id, fileArray);
            }
            await loadData();
            setSelectedSubtask(null);
            setMessage('Subtask updated successfully!');
        } catch (err) {
            setMessage('Error updating subtask: ' + ((err as any).message || err));
        }
    };

    const isCompleted = selectedSubtask?.TaskStatus === 'Completed';
    const statusOptions = [
        { key: 'Not Started', text: 'Not Started' },
        { key: 'In Progress', text: 'In Progress' },
        { key: 'Completed', text: 'Completed' },
        { key: 'On Hold', text: 'On Hold' }
    ];

    const columns = [
        {
            key: 'Task_Description', name: 'Task Description & Hierarchy', fieldName: 'Task_Description', minWidth: 250, maxWidth: 450, isResizable: true,
            isSorted: sortedColumn === 'Task_Description', isSortedDescending, onColumnClick,
            onRender: (item: IUiSubTask) => {
                const isExpanded = expandedRows.has(item.Id);
                const indent = item.depth * 24;

                return (
                    <div className={styles.hierarchyRow} style={{ paddingLeft: indent }}>
                        {item.depth > 0 && (
                            <div className={styles.connector} style={{ left: -(24 / 2) }} />
                        )}

                        <div style={{ display: 'flex', alignItems: 'center' }}>
                            {item.hasChildren ? (
                                <IconButton
                                    iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                                    styles={{ root: { height: 24, width: 24, marginRight: 4, color: '#0078d4' } }}
                                    onClick={(e) => { e.stopPropagation(); toggleExpand(item.Id); }}
                                />
                            ) : (
                                <span style={{ width: 24, display: 'inline-block', marginRight: 4 }}>
                                    {item.depth > 0 && <span className={styles['depth-indicator']}>•</span>}
                                </span>
                            )}

                            <Stack>
                                <span className={styles.title} title={item.Task_Description}>{item.Task_Description}</span>
                                {item.depth > 0 && <span style={{ fontSize: 10, color: '#666' }}>ID: {item.Id}</span>}
                            </Stack>

                            {item.TaskStatus !== 'Completed' && currentMainTaskStatus !== 'Completed' && (
                                <IconButton
                                    iconProps={{ iconName: 'Add' }}
                                    title="Add Sub-subtask"
                                    styles={{ root: { height: 24, width: 24, marginLeft: 10, color: '#107c10' } }}
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        setAddingToParentId(item.Id);
                                        setIsAdding(true);
                                        setNewTitle('');
                                        setNewDesc('');
                                        setNewCategory(item.Category || '');
                                        setSelectedUserKey(undefined);
                                        if (!expandedRows.has(item.Id)) toggleExpand(item.Id);
                                    }}
                                />
                            )}
                        </div>
                    </div>
                );
            }
        },
        {
            key: 'assigned', name: 'Assignee', minWidth: 130, maxWidth: 180, isResizable: true,
            onRender: (i: ISubTask) => {
                const name = getAssigneeName(i);
                return (
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Persona size={PersonaSize.size24} text={name} hidePersonaDetails />
                        <span style={{ fontSize: 12 }}>{name}</span>
                    </Stack>
                );
            },
            isSorted: sortedColumn === 'assigned', isSortedDescending, onColumnClick
        },
        {
            key: 'TaskStatus', name: 'Status', fieldName: 'TaskStatus', minWidth: 100, maxWidth: 120, isResizable: true,
            isSorted: sortedColumn === 'TaskStatus', isSortedDescending, onColumnClick,
            onRender: (item: ISubTask) => {
                const status = item.TaskStatus || 'Not Started';
                const statusClass = status === 'Completed' ? styles.completed :
                    status === 'In Progress' ? styles.inProgress :
                        status === 'On Hold' ? styles.onHold : styles.notStarted;
                return <span className={`${styles.statusPill} ${statusClass}`}>{status}</span>;
            }
        },
        { key: 'weight', name: 'Weight', minWidth: 60, maxWidth: 80, onRender: (): string => calculateWeightage() },
        {
            key: 'TaskDueDate', name: 'Due Date', fieldName: 'TaskDueDate', minWidth: 100, maxWidth: 120, onRender: (item: ISubTask) => {
                const isOverdue = item.TaskStatus !== 'Completed' && item.TaskDueDate && new Date(item.TaskDueDate) < new Date();
                return <span style={{ color: isOverdue ? '#d13438' : 'inherit', fontWeight: isOverdue ? 600 : 400 }}>{formatDate(item.TaskDueDate)}</span>;
            }, isSorted: sortedColumn === 'TaskDueDate', isSortedDescending, onColumnClick
        },
        {
            key: 'attachments', name: 'Files', minWidth: 60, maxWidth: 100, onRender: (item: ISubTask) => {
                const files = (item as any).AttachmentFiles;
                if (!files || files.length === 0) return null;
                return (
                    <TooltipHost content={`${files.length} attachment(s)`}>
                        <IconButton iconProps={{ iconName: 'Attach' }} style={{ height: 24 }} />
                    </TooltipHost>
                );
            }
        }
    ];

    return (
        <div className={styles.taskDetailContainer}>
            <Pivot defaultSelectedKey={initialTab || "Overview"}>
                <PivotItem headerText="Overview" itemKey="Overview">
                    <div className={styles.glassCard}>
                        <div className={styles.headerArea}>
                            <Stack>
                                <h2>{mainTask.Title}</h2>
                                <span style={{ fontSize: 12, color: '#666', marginTop: 4 }}>Main Task ID: {mainTask.Id}</span>
                            </Stack>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                {currentMainTaskStatus === 'Completed' && (
                                    <TooltipHost content="Reopen this task">
                                        <IconButton
                                            iconProps={{ iconName: 'Refresh' }}
                                            onClick={async () => {
                                                try {
                                                    await taskService.updateMainTaskStatus(mainTask.Id, 'In Progress', 'Task Reopened');
                                                    setCurrentMainTaskStatus('In Progress');
                                                    setCurrentEndDate(undefined);
                                                    setMessage('Task reopened successfully.');
                                                    if (onTaskUpdated) onTaskUpdated();
                                                } catch (e: any) {
                                                    setMessage('Error reopening task: ' + (e.message || e));
                                                }
                                            }}
                                            styles={{ root: { backgroundColor: '#f3f2f1', borderRadius: '50%' } }}
                                        />
                                    </TooltipHost>
                                )}
                                <span className={styles.statusBadge} style={{
                                    background: currentMainTaskStatus === 'Completed' ? '#107c10' :
                                        currentMainTaskStatus === 'In Progress' ? '#0078d4' :
                                            currentMainTaskStatus === 'On Hold' ? '#d13438' : '#7a7a7a'
                                }}>
                                    {currentMainTaskStatus}
                                </span>
                            </div>
                        </div>

                        <div style={{ marginBottom: 20 }}>
                            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                                <span style={{ fontSize: 13, fontWeight: 600 }}>Overall Progress</span>
                                <span style={{ fontSize: 13, fontWeight: 700, color: '#0078d4' }}>{stats.percent}%</span>
                            </Stack>
                            <ProgressIndicator
                                percentComplete={stats.percent / 100}
                                barHeight={10}
                                styles={{ itemProgress: { padding: '8px 0' } }}
                            />
                        </div>

                        <div className={styles.infoGrid}>
                            <div className={styles.infoItem}>
                                <span className={styles.label}>Business Unit</span>
                                <span className={styles.value}>{mainTask.BusinessUnit || 'N/A'}</span>
                            </div>
                            <div className={styles.infoItem}>
                                <span className={styles.label}>Department</span>
                                <span className={styles.value}>{mainTask.Departments || 'N/A'}</span>
                            </div>
                            <div className={styles.infoItem}>
                                <span className={styles.label}>Project</span>
                                <span className={styles.value}>{mainTask.Project || 'N/A'}</span>
                            </div>
                            <div className={styles.infoItem}>
                                <span className={styles.label}>Due Date</span>
                                <span className={styles.value}>{formatDate(mainTask.TaskDueDate) || 'N/A'}</span>
                            </div>
                            {currentEndDate && (
                                <div className={styles.infoItem}>
                                    <span className={styles.label}>Completed Date</span>
                                    <span className={styles.value} style={{ color: '#107c10', fontWeight: 600 }}>{formatDate(currentEndDate)}</span>
                                </div>
                            )}
                        </div>

                        <div className={styles.glassCard} style={{ background: 'rgba(255,255,255,0.4)', padding: 15, marginBottom: 15 }}>
                            <span className={styles.label} style={{ marginBottom: 8 }}>Description</span>
                            <div style={{ whiteSpace: 'pre-wrap', fontSize: 14, lineHeight: 1.6 }}>{mainTask.Task_x0020_Description || 'No description provided.'}</div>
                        </div>

                        {currentUserRemarks && (
                            <div className={styles.glassCard} style={{ background: 'rgba(255,185,0,0.05)', padding: 15, borderLeft: '4px solid #ffb900', marginBottom: 15 }}>
                                <span className={styles.label}>Latest Remarks</span>
                                <div style={{ fontStyle: 'italic', marginTop: 5, fontSize: 13 }}>{currentUserRemarks}</div>
                            </div>
                        )}

                        {attachments && attachments.length > 0 && (
                            <div style={{ marginTop: 10 }}>
                                <span className={styles.label}>Attachments</span>
                                <Stack horizontal wrap tokens={{ childrenGap: 10 }} style={{ marginTop: 8 }}>
                                    {attachments.map((file: any, index: number) => (
                                        <DefaultButton
                                            key={index}
                                            href={file.ServerRelativeUrl}
                                            target="_blank"
                                            text={file.FileName}
                                            iconProps={{ iconName: 'Attach' }}
                                            styles={{ root: { borderRadius: 20, background: 'rgba(255,255,255,0.6)', border: '1px solid #ddd' } }}
                                        />
                                    ))}
                                </Stack>
                            </div>
                        )}
                    </div>

                    <div className={styles.glassCard}>
                        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 20 } }}>
                            <Stack>
                                <h3 style={{ margin: 0 }}>Subtasks & Hierarchy</h3>
                                <span style={{ fontSize: 12, color: '#666' }}>Showing {filteredTasks.length} tasks</span>
                            </Stack>
                            <Stack horizontal tokens={{ childrenGap: 10 }}>
                                <PrimaryButton text="Export" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportToExcel} styles={{ root: { borderRadius: 20 } }} />
                                {currentMainTaskStatus !== 'Completed' && (
                                    <PrimaryButton
                                        text="New Subtask"
                                        iconProps={{ iconName: 'Add' }}
                                        onClick={() => { setIsAdding(!isAdding); setAddingToParentId(undefined); }}
                                        styles={{ root: { borderRadius: 20, background: '#107c10', border: 'none' } }}
                                    />
                                )}
                            </Stack>
                        </Stack>

                        <Stack horizontal tokens={{ childrenGap: 15 }} styles={{ root: { marginBottom: 20 } }} verticalAlign="end">
                            <Dropdown
                                label="Assignee"
                                placeholder="Filter User"
                                options={[{ key: 'All', text: 'All Users' }, ...uniqueAssignees]}
                                selectedKey={assignedFilter || 'All'}
                                onChange={(_, opt) => setAssignedFilter(opt?.key === 'All' ? undefined : opt?.key as string)}
                                styles={{ root: { width: 150 } }}
                            />
                            <Dropdown
                                label="Category"
                                placeholder="Filter Category"
                                options={[{ key: 'All', text: 'All' }, ...uniqueCategories]}
                                selectedKey={categoryFilter || 'All'}
                                onChange={(_, opt) => setCategoryFilter(opt?.key === 'All' ? undefined : opt?.key as string)}
                                styles={{ root: { width: 150 } }}
                            />
                            <DefaultButton iconProps={{ iconName: 'Clear' }} onClick={() => { setAssignedFilter(undefined); setCategoryFilter(undefined); }} styles={{ root: { borderRadius: 4, height: 32 } }} />
                        </Stack>

                        {isAdding && (
                            <div className={styles.glassCard} style={{ background: 'rgba(0, 120, 212, 0.05)', border: '1px dashed #0078d4', marginBottom: 20 }}>
                                <Stack tokens={{ childrenGap: 15 }}>
                                    <h4 style={{ margin: 0 }}>{addingToParentId ? `Add Sub-subtask to ID: ${addingToParentId}` : 'Add New Subtask'}</h4>
                                    <TextField label="Title" value={newTitle} onChange={(e, v) => setNewTitle(v || '')} required />
                                    <ComboBox
                                        label="Category"
                                        options={categoryFormOptions}
                                        selectedKey={newCategory}
                                        onChange={(_, opt) => setNewCategory(opt?.key as string || opt?.text || '')}
                                        allowFreeform={true}
                                        autoComplete="on"
                                    />
                                    <TextField label="Description" multiline rows={2} value={newDesc} onChange={(e, v) => setNewDesc(v || '')} required />
                                    <ComboBox
                                        label="Assign To"
                                        options={userOptions}
                                        selectedKey={selectedUserKey}
                                        onChange={onUserComboChange}
                                        onRenderOption={onRenderUserOption}
                                        autoComplete="on"
                                        required
                                    />
                                    <DatePicker label="Due Date" value={dueDate} onSelectDate={(d) => setDueDate(d || undefined)} isRequired />
                                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                                        <PrimaryButton text="Create Subtask" onClick={async () => {
                                            if (!newTitle || !selectedUserKey || !dueDate) {
                                                setMessage("Please fill required fields.");
                                                return;
                                            }
                                            try {
                                                await taskService.createSubTask({
                                                    Admin_Job_ID: mainTask.Id,
                                                    Task_Title: newTitle,
                                                    Task_Description: newDesc,
                                                    Category: newCategory,
                                                    TaskDueDate: dueDate.toISOString(),
                                                    TaskStatus: 'Not Started',
                                                    TaskAssignedToId: selectedUserKey as number,
                                                    ParentSubtaskId: addingToParentId
                                                } as any);
                                                setIsAdding(false);
                                                setNewTitle(''); setNewDesc(''); setNewCategory(''); setSelectedUserKey(undefined);
                                                loadData();
                                            } catch (e: any) {
                                                setMessage(e.message);
                                            }
                                        }} />
                                        <DefaultButton text="Cancel" onClick={() => setIsAdding(false)} />
                                    </Stack>
                                </Stack>
                            </div>
                        )}

                        {message && <MessageBar messageBarType={MessageBarType.info} onDismiss={() => setMessage(undefined)} styles={{ root: { marginBottom: 15 } }}>{message}</MessageBar>}

                        <div className={styles.customList}>
                            <DetailsList
                                items={filteredTasks}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                layoutMode={DetailsListLayoutMode.justified}
                                onItemInvoked={handleSubtaskClick}
                            />
                        </div>
                    </div>
                </PivotItem>

                <PivotItem headerText="Workflow" itemKey="Workflow" itemIcon="Org">
                    <div style={{ height: '800px', marginTop: 10 }}>
                        <WorkflowDesigner mainTaskId={mainTask.Id} readonly={true} />
                    </div>
                </PivotItem>

                <PivotItem headerText="Activity Log" itemKey="Activity" itemIcon="History">
                    <div className={styles.glassCard} style={{ marginTop: 20 }}>
                        <Stack tokens={{ childrenGap: 15 }}>
                            {correspondence.length === 0 ? (
                                <div style={{ textAlign: 'center', padding: 40, color: '#666' }}>
                                    <Icon iconName="Message" style={{ fontSize: 32, color: '#ddd', display: 'block', marginBottom: 10 }} />
                                    No activity history available for this task.
                                </div>
                            ) : (
                                correspondence.map((item: any, idx: number) => (
                                    <div key={idx} style={{
                                        marginBottom: 15,
                                        padding: 15,
                                        background: '#f8fbff',
                                        borderRadius: '16px 16px 16px 4px',
                                        border: '1px solid #e1f0fb',
                                        boxShadow: '0 2px 5px rgba(0,0,0,0.03)'
                                    }}>
                                        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8, borderBottom: '1px solid #edf2f6', paddingBottom: 5 }}>
                                            <span style={{ fontWeight: 600, color: '#0078d4', fontSize: 13 }}>
                                                {item.Author?.Title || item.FromAddress || 'System'}
                                            </span>
                                            <span style={{ color: '#999', fontSize: 11 }}>
                                                {formatDate(item.Created)}
                                            </span>
                                        </div>
                                        <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 5 }}>{item.Title}</div>
                                        <div dangerouslySetInnerHTML={{ __html: sanitizeHtml(item.MessageBody) }} style={{ fontSize: 13, color: '#333', lineHeight: '1.5' }} />
                                    </div>
                                ))
                            )}
                        </Stack>
                    </div>
                </PivotItem>
            </Pivot>

            {/* Subtask Details / Edit Modal (Centered) */}
            <Modal
                isOpen={!!selectedSubtask}
                onDismiss={() => setSelectedSubtask(null)}
                isBlocking={false}
                containerClassName={`${styles.glassCard} ${styles.animateScaleIn}`}
                styles={{ main: { maxWidth: 850, width: '92%', borderRadius: 20, padding: 0, background: 'rgba(255, 255, 255, 0.8)', backdropFilter: 'blur(20px)' } }}
            >
                {selectedSubtask && (
                    <div style={{ padding: '30px 40px', maxHeight: '90vh', overflowY: 'auto' }}>
                        {/* Custom Header Bar */}
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 25 }} className={styles.animateFadeIn}>
                            <div>
                                <h2 style={{ margin: 0, color: '#0078d4', fontSize: 28, fontWeight: 800, letterSpacing: '-0.5px' }}>{isCompleted ? "Subtask Details" : "Update Subtask"}</h2>
                                <div style={{ fontSize: 13, color: '#666', marginTop: 4 }}>Manage and track your subtask progress below</div>
                            </div>
                            <IconButton
                                iconProps={{ iconName: 'Cancel' }}
                                onClick={() => setSelectedSubtask(null)}
                                styles={{ root: { background: '#f5f5f5', borderRadius: '50%', padding: 20 } }}
                            />
                        </div>

                        <Stack tokens={{ childrenGap: 24 }}>
                            {/* Detailed Info Card */}
                            <div className={`${styles.premiumCard} ${styles.animateSlideUp} ${styles.stagger1}`}>
                                {/* Main Task Context */}
                                <div style={{ marginBottom: 20 }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                                        <div style={{ background: 'rgba(0,120,212,0.1)', padding: 6, borderRadius: 6 }}>
                                            <Icon iconName="ProjectCollection" style={{ fontSize: 14, color: '#0078d4' }} />
                                        </div>
                                        <span style={{ fontSize: 11, fontWeight: 700, color: '#0078d4', textTransform: 'uppercase', letterSpacing: '1px' }}>Parent Task Context</span>
                                    </div>
                                    <h4 style={{ margin: 0, color: '#323130', fontSize: 16, fontWeight: 600 }}>{mainTask.Title}</h4>
                                    <div style={{ fontSize: 12, color: '#666', marginTop: 6, lineHeight: 1.6, background: 'rgba(255,255,255,0.4)', padding: '8px 12px', borderRadius: 8 }}>{mainTask.Task_x0020_Description}</div>
                                </div>

                                <Separator styles={{ root: { margin: '20px 0', opacity: 0.3 } }} />

                                {/* Subtask Specific Details */}
                                <div>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                                        <div style={{ background: 'rgba(0,90,158,0.1)', padding: 6, borderRadius: 6 }}>
                                            <Icon iconName="TaskGroup" style={{ fontSize: 14, color: '#005a9e' }} />
                                        </div>
                                        <span style={{ fontSize: 11, fontWeight: 700, color: '#005a9e', textTransform: 'uppercase', letterSpacing: '1px' }}>Current Subtask</span>
                                    </div>
                                    <h3 style={{ margin: '0 0 12px 0', color: '#005a9e', fontSize: 22, fontWeight: 700 }}>{selectedSubtask.Task_Title}</h3>
                                    <div style={{ fontSize: 14, lineHeight: 1.6, color: '#444', background: 'rgba(255,255,255,0.3)', padding: 15, borderRadius: 10, border: '1px solid rgba(0,0,0,0.02)' }}>{selectedSubtask.Task_Description}</div>

                                    <div className={styles.infoGrid} style={{ marginTop: 25, gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 20 }}>
                                        <div className={styles.infoItem} style={{ background: 'white', boxShadow: '0 2px 8px rgba(0,0,0,0.02)' }}>
                                            <span className={styles.label}>Assignee</span>
                                            <ComboBox
                                                options={userOptions}
                                                selectedKey={editAssigneeId}
                                                onChange={(_, opt) => {
                                                    if (opt && opt.key !== editAssigneeId) {
                                                        setChangePrompt({ visible: true, field: 'Assignee', newValue: opt.key });
                                                    }
                                                }}
                                                onRenderOption={onRenderUserOption}
                                                autoComplete="on"
                                                allowFreeform={true}
                                                disabled={isCompleted}
                                                styles={{ root: { width: '100%', border: 'none', borderBottom: '1px solid #ccc' } }}
                                            />
                                        </div>
                                        <div className={styles.infoItem} style={{ background: 'white', boxShadow: '0 2px 8px rgba(0,0,0,0.02)' }}>
                                            <span className={styles.label}>Target Date</span>
                                            <DatePicker
                                                value={editDueDate}
                                                onSelectDate={(date) => {
                                                    if (date && date.toISOString() !== editDueDate?.toISOString()) {
                                                        setChangePrompt({ visible: true, field: 'DueDate', newValue: date.toISOString() });
                                                    }
                                                }}
                                                underlined
                                                disabled={isCompleted}
                                                styles={{ root: { width: '100%', marginTop: 4 } }}
                                                formatDate={(date) => formatDate(date)}
                                            />
                                        </div>
                                        <div className={styles.infoItem} style={{ background: 'white', boxShadow: '0 2px 8px rgba(0,0,0,0.02)' }}>
                                            <span className={styles.label}>Workstream</span>
                                            <span className={styles.value}>{selectedSubtask.Category || 'General Work'}</span>
                                        </div>
                                        <div className={styles.infoItem} style={{ background: 'white', boxShadow: '0 2px 8px rgba(0,0,0,0.02)' }}>
                                            <span className={styles.label}>Status Tracking</span>
                                            <div className={styles.statusPill} style={{
                                                padding: '6px 14px',
                                                borderRadius: 16,
                                                marginTop: 4,
                                                background: selectedSubtask.TaskStatus === 'Completed' ? 'linear-gradient(135deg, #107c10, #32a032)' :
                                                    selectedSubtask.TaskStatus === 'In Progress' ? 'linear-gradient(135deg, #0078d4, #2b88d8)' :
                                                        selectedSubtask.TaskStatus === 'On Hold' ? 'linear-gradient(135deg, #d13438, #e06666)' : '#999'
                                            }}>
                                                {selectedSubtask.TaskStatus}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <div className={`${styles.animateSlideUp} ${styles.stagger2}`}>
                                <Pivot
                                    selectedKey={subtaskInitialTab}
                                    onLinkClick={(item) => item && setSubtaskInitialTab(item.props.itemKey || 'Activity')}
                                    styles={{
                                        root: { borderBottom: '1px solid #eee', marginBottom: 20 },
                                        link: { height: 45, fontSize: 16, fontWeight: 600 },
                                        linkIsSelected: { color: '#0078d4', borderBottom: '3px solid #0078d4' }
                                    }}
                                >
                                    <PivotItem headerText="Activity Feed" itemKey="Activity" itemIcon="Message">
                                        <Stack tokens={{ childrenGap: 24 }} style={{ marginTop: 25 }} className={styles.animateFadeIn}>
                                            {/* Send Query Button */}
                                            {selectedSubtask && (selectedSubtask.TaskAssignedTo as any)?.EMail && (
                                                <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 10 }}>
                                                    <PrimaryButton
                                                        text="Send Query to Assignee"
                                                        iconProps={{ iconName: 'Send' }}
                                                        onClick={() => setShowQueryDialog(true)}
                                                        styles={{
                                                            root: {
                                                                borderRadius: 20,
                                                                height: 40,
                                                                padding: '0 20px',
                                                                background: 'linear-gradient(135deg, #0078d4, #2b88d8)',
                                                                border: 'none',
                                                                boxShadow: '0 4px 12px rgba(0, 120, 212, 0.3)'
                                                            }
                                                        }}
                                                    />
                                                </div>
                                            )}

                                            <div style={{
                                                maxHeight: 450,
                                                overflowY: 'auto',
                                                padding: '20px 10px',
                                                background: 'rgba(240,244,248,0.4)',
                                                borderRadius: 20,
                                                border: '1px solid rgba(0,0,0,0.03)'
                                            }}>
                                                {subtaskCorrespondence.length === 0 ? (
                                                    <div style={{ textAlign: 'center', padding: 60, color: '#999' }}>
                                                        <div style={{ background: '#f0f0f0', width: 60, height: 60, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 15px' }}>
                                                            <Icon iconName="Message" style={{ fontSize: 24 }} />
                                                        </div>
                                                        <div style={{ fontWeight: 600 }}>No communication history</div>
                                                        <div style={{ fontSize: 12 }}>Be the first to post a remark for this subtask</div>
                                                    </div>
                                                ) : (
                                                    <Stack tokens={{ childrenGap: 20 }}>
                                                        {subtaskCorrespondence.map((msg: any, idx: number) => (
                                                            <div key={idx} className={`${styles.chatBubbleCentered} ${styles.animateScaleIn}`} style={{ animationDelay: `${idx * 0.05}s` }}>
                                                                <div style={{
                                                                    display: 'flex',
                                                                    justifyContent: 'space-between',
                                                                    marginBottom: 10,
                                                                    fontSize: 12
                                                                }}>
                                                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                                                        <Persona size={PersonaSize.size24} text={msg.Author?.Title || msg.FromAddress} />
                                                                        <strong style={{ color: '#0078d4' }}>{msg.Author?.Title || msg.FromAddress}</strong>
                                                                    </div>
                                                                    <span style={{ color: '#999', fontSize: 11 }}>{formatDate(msg.Created)}</span>
                                                                </div>
                                                                <div dangerouslySetInnerHTML={{ __html: sanitizeHtml(msg.MessageBody) }} style={{ fontSize: 14, color: '#333', lineHeight: '1.7' }} />
                                                            </div>
                                                        ))}
                                                    </Stack>
                                                )}
                                            </div>

                                            {/* Composer Area */}
                                            <div className={styles.premiumCard} style={{ background: 'white', padding: 15, borderRadius: 20 }}>
                                                <TextField
                                                    placeholder="Type your question or progress remark here..."
                                                    multiline
                                                    rows={2}
                                                    value={emailBody}
                                                    onChange={(_, v) => setEmailBody(v || '')}
                                                    borderless
                                                    styles={{ root: { marginBottom: 10 } }}
                                                />
                                                <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                                                    <PrimaryButton
                                                        text={isSubtaskSending ? "Processing..." : "Submit Remark"}
                                                        iconProps={{ iconName: 'Send' }}
                                                        onClick={handleSubtaskCorrespondence}
                                                        disabled={isSubtaskSending || !emailBody.trim()}
                                                        styles={{ root: { borderRadius: 12, height: 40, padding: '0 25px' } }}
                                                    />
                                                </div>
                                            </div>
                                        </Stack>
                                    </PivotItem>

                                    <PivotItem headerText="Update Status" itemKey="Update" itemIcon="Edit">
                                        <Stack tokens={{ childrenGap: 24 }} style={{ marginTop: 25 }} className={styles.animateFadeIn}>
                                            <div className={styles.premiumCard} style={{ background: 'white' }}>
                                                <Stack tokens={{ childrenGap: 20 }}>
                                                    <Dropdown
                                                        label="Current Progress Status"
                                                        selectedKey={editStatus}
                                                        options={statusOptions}
                                                        onChange={(_, opt) => setEditStatus(opt?.key as string)}
                                                        disabled={isCompleted}
                                                    />
                                                    <TextField
                                                        label="Executive Remarks / Progress Detailed"
                                                        multiline
                                                        rows={4}
                                                        value={editRemarks}
                                                        onChange={(_, v) => setEditRemarks(v || '')}
                                                        placeholder="Provide a detailed update on the work performed..."
                                                        disabled={isCompleted}
                                                    />
                                                </Stack>
                                            </div>

                                            <div className={styles.premiumCard} style={{ background: 'white' }}>
                                                <h4 style={{ marginTop: 0, marginBottom: 15, display: 'flex', alignItems: 'center', gap: 10, color: '#333' }}>
                                                    <Icon iconName="Attach" style={{ color: '#0078d4' }} /> Subtask Artifacts & Attachments
                                                </h4>
                                                <Stack tokens={{ childrenGap: 12 }}>
                                                    {subtaskAttachments.length === 0 ? (
                                                        <div style={{ padding: '20px', textAlign: 'center', background: '#f9f9f9', borderRadius: 12, border: '1px dashed #ccc' }}>
                                                            <div style={{ fontSize: 13, color: '#888' }}>No files attached to this subtask.</div>
                                                        </div>
                                                    ) : (
                                                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: 10 }}>
                                                            {subtaskAttachments.map((file, idx) => (
                                                                <div key={idx} style={{
                                                                    display: 'flex',
                                                                    alignItems: 'center',
                                                                    gap: 12,
                                                                    padding: '12px',
                                                                    background: '#f8fbff',
                                                                    borderRadius: 12,
                                                                    border: '1px solid #e1f0fb',
                                                                    transition: 'all 0.2s ease'
                                                                }} className={styles.animateScaleIn}>
                                                                    <div style={{ background: '#0078d4', padding: 8, borderRadius: 8 }}>
                                                                        <Icon iconName="Document" style={{ color: 'white', fontSize: 14 }} />
                                                                    </div>
                                                                    <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer" style={{ fontSize: 13, color: '#333', textDecoration: 'none', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                                                        {file.FileName}
                                                                    </a>
                                                                </div>
                                                            ))}
                                                        </div>
                                                    )}
                                                    {!isCompleted && (
                                                        <div style={{ marginTop: 10 }}>
                                                            <input
                                                                type="file"
                                                                multiple
                                                                id="subtask-file-upload"
                                                                style={{ display: 'none' }}
                                                                onChange={(e) => setEditFiles(e.target.files)}
                                                            />
                                                            <DefaultButton
                                                                text={editFiles ? `${editFiles.length} folders/files selected` : "Upload New Artifacts"}
                                                                iconProps={{ iconName: 'Add' }}
                                                                onClick={() => document.getElementById('subtask-file-upload')?.click()}
                                                                styles={{ root: { borderRadius: 12, height: 40, border: '2px dashed #0078d4', color: '#0078d4' } }}
                                                            />
                                                        </div>
                                                    )}
                                                </Stack>
                                            </div>

                                        </Stack>
                                    </PivotItem>
                                </Pivot>

                                {!isCompleted && (
                                    <div style={{ display: 'flex', justifyContent: 'center', marginTop: 30, paddingBottom: 20 }}>
                                        <PrimaryButton
                                            text="Save Official Changes"
                                            iconProps={{ iconName: 'Save' }}
                                            onClick={handleSaveSubtask}
                                            styles={{ root: { borderRadius: 30, height: 50, padding: '0 40px', fontSize: 16, boxShadow: '0 8px 20px rgba(0, 120, 212, 0.3)' } }}
                                        />
                                    </div>
                                )}
                            </div>
                        </Stack>
                    </div>
                )}
            </Modal>

            <Dialog
                hidden={!showCompleteDialog}
                onDismiss={() => setShowCompleteDialog(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Complete Main Task',
                    subText: 'Are you sure all subtasks are finished and you want to mark this main task as COMPLETED?'
                }}
                modalProps={{ isBlocking: true, className: styles.glassCard }}
            >
                <TextField
                    label="Completion Remarks"
                    multiline
                    rows={4}
                    value={completeRemarks}
                    onChange={(_, v) => setCompleteRemarks(v || '')}
                    placeholder="Final notes for this task..."
                />
                <DialogFooter>
                    <PrimaryButton onClick={handleCompleteTask} text="Confirm Completion" styles={{ root: { background: '#107c10', border: 'none' } }} />
                    <DefaultButton onClick={() => setShowCompleteDialog(false)} text="Cancel" />
                </DialogFooter>
            </Dialog>

            <Dialog
                hidden={!changePrompt?.visible}
                onDismiss={() => { setChangePrompt(null); setChangeRemark(''); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: `Reason for ${changePrompt?.field === 'Assignee' ? 'Change of Assignee' : 'Change of Target Date'}`,
                    subText: `Please provide a remark explaining the change for the official log.`
                }}
                modalProps={{ isBlocking: true }}
            >
                <TextField
                    label="Change Remark"
                    multiline
                    rows={3}
                    value={changeRemark}
                    onChange={(_, v) => setChangeRemark(v || '')}
                    placeholder="E.g. Extended scope, User on leave, etc."
                />
                <DialogFooter>
                    <PrimaryButton
                        text="Update Field"
                        disabled={!changeRemark.trim()}
                        onClick={async () => {
                            if (!selectedSubtask || !changePrompt) return;
                            try {
                                setMessage(undefined);
                                await taskService.updateSubTaskField({
                                    subTaskId: selectedSubtask.Id,
                                    mainTaskId: mainTask.Id,
                                    title: selectedSubtask.Task_Title,
                                    field: changePrompt.field,
                                    newValue: changePrompt.newValue,
                                    remark: changeRemark
                                });

                                // Update local state immediately
                                if (changePrompt.field === 'Assignee') setEditAssigneeId(changePrompt.newValue as number);
                                else if (changePrompt.field === 'DueDate') setEditDueDate(new Date(changePrompt.newValue));

                                setChangePrompt(null);
                                setChangeRemark('');
                                setMessage(`${changePrompt.field} updated and notification sent!`);
                                await loadData();
                            } catch (e: any) {
                                setMessage("Error: " + e.message);
                            }
                        }}
                    />
                    <DefaultButton text="Cancel" onClick={() => { setChangePrompt(null); setChangeRemark(''); }} />
                </DialogFooter>
            </Dialog>

            {/* Send Query to Assignee Dialog */}
            <Dialog
                hidden={!showQueryDialog}
                onDismiss={() => { setShowQueryDialog(false); setQueryMessage(''); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: '📩 Send Query to Assignee',
                    subText: selectedSubtask ? `Send a query to ${(selectedSubtask.TaskAssignedTo as any)?.Title || 'the assignee'} regarding this task` : ''
                }}
                modalProps={{
                    isBlocking: true,
                    styles: { main: { maxWidth: 600, borderRadius: 12 } }
                }}
            >
                {selectedSubtask && (
                    <Stack tokens={{ childrenGap: 15 }}>
                        {/* Assignee Info */}
                        <div style={{
                            background: 'rgba(0, 120, 212, 0.05)',
                            padding: 15,
                            borderRadius: 10,
                            border: '1px solid rgba(0, 120, 212, 0.2)'
                        }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 }}>
                                <Icon iconName="Contact" style={{ color: '#0078d4', fontSize: 16 }} />
                                <span style={{ fontWeight: 600, color: '#0078d4' }}>Assignee</span>
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                                <Persona
                                    size={PersonaSize.size32}
                                    text={(selectedSubtask.TaskAssignedTo as any)?.Title || 'Unknown'}
                                    secondaryText={(selectedSubtask.TaskAssignedTo as any)?.EMail || ''}
                                />
                            </div>
                        </div>

                        {/* Task Info */}
                        <div style={{
                            background: 'rgba(0, 90, 158, 0.05)',
                            padding: 12,
                            borderRadius: 8,
                            borderLeft: '3px solid #005a9e'
                        }}>
                            <div style={{ fontSize: 11, color: '#666', marginBottom: 4 }}>REGARDING TASK</div>
                            <div style={{ fontWeight: 600, color: '#005a9e' }}>{selectedSubtask.Task_Title}</div>
                        </div>

                        {/* Query Message */}
                        <TextField
                            label="Your Query"
                            multiline
                            rows={5}
                            value={queryMessage}
                            onChange={(_, v) => setQueryMessage(v || '')}
                            placeholder="Type your question or clarification request here..."
                            required
                            styles={{
                                field: {
                                    fontSize: 14,
                                    lineHeight: 1.6
                                }
                            }}
                        />

                        <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 8 } }}>
                            <div style={{ fontSize: 12 }}>
                                <Icon iconName="Info" style={{ marginRight: 5 }} />
                                The assignee will receive a notification with your query and can respond via the Activity Feed.
                            </div>
                        </MessageBar>
                    </Stack>
                )}
                <DialogFooter>
                    <PrimaryButton
                        text={isSendingQuery ? "Sending..." : "Send Query"}
                        iconProps={{ iconName: 'Send' }}
                        onClick={handleSendQueryToAssignee}
                        disabled={isSendingQuery || !queryMessage.trim()}
                        styles={{
                            root: {
                                borderRadius: 20,
                                background: 'linear-gradient(135deg, #0078d4, #2b88d8)',
                                border: 'none'
                            }
                        }}
                    />
                    <DefaultButton
                        text="Cancel"
                        onClick={() => { setShowQueryDialog(false); setQueryMessage(''); }}
                        styles={{ root: { borderRadius: 20 } }}
                    />
                </DialogFooter>
            </Dialog>
        </div>
    );
};
