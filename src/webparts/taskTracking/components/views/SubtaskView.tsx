/* eslint-disable max-lines */
import * as React from 'react';
import { DetailsList, SelectionMode, Dropdown, IDropdownOption, ConstrainMode, Panel, PanelType, TextField, PrimaryButton, DefaultButton, Stack, MessageBar, MessageBarType, IconButton, Checkbox, IColumn, ScrollablePane, ScrollbarVisibility, DatePicker, ComboBox, IComboBox, IComboBoxOption, SelectableOptionMenuItemType, Persona, PersonaSize, Sticky, StickyPositionType, DetailsListLayoutMode, IDetailsHeaderProps, Dialog, DialogFooter, DialogType, Icon, SearchBox } from 'office-ui-fabric-react';
import useSpeechRecognition from '../../hooks/useSpeechRecognition';
import { VoiceControl } from '../common/VoiceControl';
import { taskService } from '../../../../services/sp-service';
import { ISubTask, IMainTask, ITaskCorrespondence } from '../../../../services/interfaces';
import { LIST_SUB_TASKS } from '../../../../services/interfaces';

export interface ISubtaskViewProps {
    userEmail: string;
    initialChildTaskId?: number;
    onDeepLinkProcessed?: () => void;
}

// Extended Interface for UI
interface IUiSubTask extends ISubTask {
    depth: number;
    hasChildren: boolean;
    isExpanded?: boolean;
}

export const SubtaskView: React.FunctionComponent<ISubtaskViewProps> = (props) => {
    const [tasks, setTasks] = React.useState<ISubTask[]>([]);
    const [uiSubTasks, setUiSubTasks] = React.useState<IUiSubTask[]>([]); // Processed list
    const [mainTaskMap, setMainTaskMap] = React.useState<Map<number, IMainTask>>(new Map()); // STORE MAIN TASKS
    const [loading, setLoading] = React.useState<boolean>(true);

    // Filter State
    const [statusFilter, setStatusFilter] = React.useState<string | undefined>(undefined);
    const [categoryFilter, setCategoryFilter] = React.useState<string | undefined>(undefined);
    const [overdueFilter, setOverdueFilter] = React.useState<boolean>(false);
    const [showCompleted, setShowCompleted] = React.useState<boolean>(false); // NEW: Default to false
    const [searchTerm, setSearchTerm] = React.useState<string>('');

    // Sort State
    const [sortedColumn, setSortedColumn] = React.useState<string | undefined>(undefined);
    const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

    // Edit Subtask Panel State
    const [selectedSubtask, setSelectedSubtask] = React.useState<ISubTask | null>(null);
    const [editStatus, setEditStatus] = React.useState<string>('');
    const [editRemarks, setEditRemarks] = React.useState<string>('');
    const [newAttachFiles, setNewAttachFiles] = React.useState<File[]>([]);
    const [saving, setSaving] = React.useState<boolean>(false);
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

    // Create Subtask State
    const [isAdding, setIsAdding] = React.useState<boolean>(false);
    const [addingToParentId, setAddingToParentId] = React.useState<number | undefined>(undefined);
    const [newTitle, setNewTitle] = React.useState<string>('');
    const [newDesc, setNewDesc] = React.useState<string>('');
    const [newCategory, setNewCategory] = React.useState<string>('');
    const [dueDate, setDueDate] = React.useState<Date | undefined>(new Date());
    const [selectedUserKey, setSelectedUserKey] = React.useState<string | number | undefined>(undefined);
    const [userOptions, setUserOptions] = React.useState<IComboBoxOption[]>([]);
    const [createAttachments, setCreateAttachments] = React.useState<File[]>([]);

    // Expand/Collapse State
    const [expandedRows, setExpandedRows] = React.useState<Set<number>>(new Set());

    // Incomplete Children Dialog State
    const [showIncompleteDialog, setShowIncompleteDialog] = React.useState<boolean>(false);
    const [incompleteTasks, setIncompleteTasks] = React.useState<ISubTask[]>([]);
    const [pendingStatusChange, setPendingStatusChange] = React.useState<{ item: ISubTask; option: IDropdownOption; remarks?: string } | null>(null);

    // Voice Control Hook
    const { isListening, transcript, startListening, stopListening, hasRecognitionSupport, error: voiceError } = useSpeechRecognition();

    React.useEffect(() => {
        if (voiceError) {
            setMessage(`Voice Error: ${voiceError} `);
        }
    }, [voiceError]);

    // Voice Command Processor
    React.useEffect(() => {
        if (!transcript) return;
        const lowerTranscript = transcript.toLowerCase().trim();

        console.log("Voice Command:", lowerTranscript);

        // Command: "Search for..." or just "Search..."
        if (lowerTranscript.indexOf('search for ') === 0) {
            const term = lowerTranscript.replace('search for ', '').trim();
            if (term) {
                setSearchTerm(term);
                setMessage(`Searching for: "${term}"`);
            }
        }
        else if (lowerTranscript.indexOf('search ') === 0) {
            const term = lowerTranscript.replace('search ', '').trim();
            if (term) {
                setSearchTerm(term);
                setMessage(`Searching for: "${term}"`);
            }
        }

        // Command: "Create Task" or "New Task"
        else if (lowerTranscript.indexOf('create task') !== -1 || lowerTranscript.indexOf('new task') !== -1) {
            setIsAdding(true);
            setMessage('Opening new task form...');
        }

        // Command: "Clear Filters" or "Reset"
        else if (lowerTranscript.indexOf('clear filters') !== -1 || lowerTranscript.indexOf('reset') !== -1) {
            resetFilters();
            setMessage('Filters reset.');
        }

        // Command: "Show Completed"
        else if (lowerTranscript.indexOf('show completed') !== -1) {
            setShowCompleted(true);
            setMessage('Showing completed tasks.');
        }

        // Command: "Hide Completed"
        else if (lowerTranscript.indexOf('hide completed') !== -1) {
            setShowCompleted(false);
            setMessage('Hiding completed tasks.');
        }

        // Command: "My Tasks" (Filter)
        else if (lowerTranscript.indexOf('my tasks') !== -1) {
            resetFilters();
            setMessage('Showing default view.');
        }

        // Fallback: Default to Search if it looks like a query (more than 1 char) and didn't match commands?
        // OR just tell the user what we heard.
        else {
            // Enhancing UX: If it's a single word that isn't a command, maybe they just want to search it?
            // Let's explicitly tell them if it's not recognized.
            setMessage(`Heard: "${transcript}"(Command not recognized)`);

            // Optional: If they just say a word, maybe we search for it? 
            // setSearchTerm(lowerTranscript); // Let's try explicit search for now to avoid accidental jumps.
        }

    }, [transcript]);

    React.useEffect(() => {
        loadTasks().catch(console.error);
    }, [props.userEmail]);


    // Deep Link Logic
    React.useEffect(() => {
        if (loading || tasks.length === 0) return;

        const { initialChildTaskId, onDeepLinkProcessed } = props;

        if (initialChildTaskId) {
            const targets = tasks.filter((t: ISubTask) => t.Id === initialChildTaskId);
            const target = targets.length > 0 ? targets[0] : null;
            if (target) {
                console.log('[SubtaskView] Deep linking to subtask:', initialChildTaskId);
                openEditPanel(target);

                // Clear parent state immediately after opening
                if (onDeepLinkProcessed) onDeepLinkProcessed();
            } else {
                console.warn('[SubtaskView] Target subtask not found in user list:', initialChildTaskId);
            }
        }
    }, [tasks, loading, props.initialChildTaskId]);

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
        const processedIds = new Set<number>();

        const traverse = (parentId: number, depth: number) => {
            const children = childrenMap.get(parentId) || [];
            children.sort((a, b) => a - b); // Sort by ID

            children.forEach(childId => {
                // Only process if this child exists in our user's list (items)
                // If parent isn't in list (e.g. assigned to someone else), this child logic handles it?
                // Actually, if we start from 0, we find Top Level items.
                // What if a child is in 'items' but its parent is NOT?
                // It means 'item.ParentSubtaskId' points to an ID that is not in 'items'.
                // Then that item should be treated as top level for this view.

                const child = itemMap.get(childId);
                if (child) {
                    result.push({
                        ...child,
                        depth: depth,
                        hasChildren: childrenMap.has(child.Id) && childrenMap.get(child.Id)!.length > 0
                    });
                    processedIds.add(childId);
                    traverse(child.Id, depth + 1);
                }
            });
        };

        // Standard traversal for items linked to Root (0) or undefined
        traverse(0, 0);

        // Handling "Orphaned" items (where parent exists in DB but not in this user's view)
        // We iterate specifically to find items we haven't processed yet
        items.forEach(item => {
            if (!processedIds.has(item.Id)) {
                // This item wasn't reached from 0. 
                // Either it's a root item we missed (shouldn't happen if pId=0 is mapped to 0)
                // OR it has a parent that is NOT in this filtered list.
                // We treat it as a new root.

                // We restart traversal from here, but treat this node as depth 0
                // We need to verify we don't double process if there's a cycle? (Assumed acyclic)

                // Add this item
                result.push({
                    ...item,
                    depth: 0,
                    hasChildren: childrenMap.has(item.Id) && childrenMap.get(item.Id)!.length > 0
                });
                processedIds.add(item.Id);

                // And process its children
                traverse(item.Id, 1);
            }
        });

        return result;
    };



    const [categoryFormOptions, setCategoryFormOptions] = React.useState<IComboBoxOption[]>([]);

    React.useEffect(() => {
        const loadCategories = async () => {
            try {
                // Fetch Category choices from 'Task Tracking System User' list
                const choices = await taskService.getChoiceFieldOptions('Task Tracking System User', 'Category');
                setCategoryFormOptions(choices.map(c => ({ key: c, text: c })));
            } catch (e) {
                console.warn('Could not load Review/Category choices', e);
            }
        };
        loadCategories();
    }, []);

    const loadTasks = async () => {
        setLoading(true);
        try {
            // 1. Get tasks directly assigned to me
            const myTasks = await taskService.getSubTasksForUser(props.userEmail);

            // 2. Identify the Main Task IDs I am involved in
            const rawIds = myTasks.map(t => t.Admin_Job_ID).filter(id => id);
            const mainTaskIds = rawIds.filter((item, pos) => rawIds.indexOf(item) === pos);

            let allRelatedTasks: ISubTask[] = [];

            if (mainTaskIds.length > 0) {
                // 3. Fetch ALL subtasks for these Main Tasks (to see peers and children assigned to others)
                allRelatedTasks = await taskService.getSubTasksByMainTaskIds(mainTaskIds);

                // 3b. Fetch Main Tasks details for validation (Due Date)
                try {
                    const mainTasks = await taskService.getMainTasksByIds(mainTaskIds);
                    const mMap = new Map<number, IMainTask>();
                    mainTasks.forEach(m => mMap.set(m.Id, m));
                    setMainTaskMap(mMap);
                } catch (e) {
                    console.warn("Could not fetch main tasks details", e);
                }
            } else {
                // Fallback if no tasks found
                allRelatedTasks = myTasks;
            }

            // 4. FILTER: Keep only tasks that are assigned to Me OR are descendants of tasks assigned to Me.
            const myEmail = props.userEmail.toLowerCase();
            const keepIds = new Set<number>();

            // Build adjacency list
            const childrenMap = new Map<number, number[]>();
            allRelatedTasks.forEach(t => {
                const pId = t.ParentSubtaskId || 0;
                if (!childrenMap.has(pId)) childrenMap.set(pId, []);
                childrenMap.get(pId)!.push(t.Id);
            });

            // Helper to collect all descendants recursively
            const collectDescendants = (parentId: number) => {
                const children = childrenMap.get(parentId);
                if (children) {
                    children.forEach(childId => {
                        if (!keepIds.has(childId)) {
                            keepIds.add(childId);
                            collectDescendants(childId);
                        }
                    });
                }
            };

            // Identify "My Tasks" and start collecting
            allRelatedTasks.forEach(t => {
                const assigned = t.TaskAssignedTo;
                let isMine = false;

                // Check if assigned to me
                if (Array.isArray(assigned)) {
                    isMine = assigned.some(u => u.EMail?.toLowerCase() === myEmail);
                } else if (assigned && (assigned as any).EMail) {
                    isMine = (assigned as any).EMail?.toLowerCase() === myEmail;
                }

                if (isMine) {
                    keepIds.add(t.Id);
                    collectDescendants(t.Id);
                }
            });

            // Final Filter
            const filteredHierarchy = allRelatedTasks.filter(t => keepIds.has(t.Id));

            setTasks(filteredHierarchy);
            setUiSubTasks(processHierarchy(filteredHierarchy));

            // Fetch clarification indicators
            const ids = filteredHierarchy.map(t => t.Id);
            const metadata = await taskService.getTaskCorrespondenceMetadata(ids);
            setClarificationMetadata(metadata);

            // Load users for ComboBox
            const users = await taskService.getSiteUsers();
            const options: IComboBoxOption[] = users.map(u => ({
                key: u.Id,
                text: u.Title,
                data: { email: u.Email }
            }));
            setUserOptions(options);
        } catch (error) {
            console.error(error);
        } finally {
            setLoading(false);
        }
    };

    // Format date as dd-MMM-yyyy (e.g., 07-DEC-2025)
    const formatDate = (date: string | Date | undefined): string => {
        if (!date) return '';
        const d = new Date(date as any);
        const dayNum = d.getDate();
        const day = dayNum < 10 ? '0' + dayNum : dayNum.toString();
        const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        const month = months[d.getMonth()];
        const year = d.getFullYear();
        return `${day} -${month} -${year} `;
    };

    // Check if a task is overdue
    const isOverdue = (item: ISubTask): boolean => {
        if (item.TaskStatus === 'Completed' || !item.TaskDueDate) return false;

        const due = new Date(item.TaskDueDate);
        due.setHours(23, 59, 59, 999); // Due date includes the whole day

        const now = new Date();
        now.setHours(0, 0, 0, 0); // Compare against start of today? 
        // Logic: specific date (e.g. 17th) is overdue ONLY if today is 18th.
        // So due < now (where now is "just now") might fail if due is 17th midnight.
        // Let's settle: Overdue means Due Date < Today (Start of today). 
        // If due is 17th, and today is 17th -> Not overdue.
        // If due is 16th, and today is 17th -> Overdue.

        // Resetting due to midnight for safe comparison
        const dueDay = new Date(item.TaskDueDate);
        dueDay.setHours(0, 0, 0, 0);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        return dueDay < today;
    };

    // Get unique values for filter dropdowns
    const getUniqueValues = (arr: string[]): string[] => {
        const seen: { [key: string]: boolean } = {};
        return arr.filter((item) => {
            if (seen[item]) return false;
            seen[item] = true;
            return true;
        });
    };

    const statusList = getUniqueValues(tasks.map(t => t.TaskStatus || 'Not Started'));
    const statusOptions = [{ key: '', text: 'All Status' }].concat(statusList.map(s => ({ key: s, text: s })));

    const categoryList = getUniqueValues(tasks.map(t => t.Category || 'Unknown').filter(c => c));
    const categoryOptions = [{ key: '', text: 'All Categories' }].concat(categoryList.map(c => ({ key: c, text: c })));

    // Apply filters and sorting
    // Apply filters and sorting
    const toggleExpand = (id: number) => {
        const newExpanded = new Set<number>();
        expandedRows.forEach((r) => newExpanded.add(r));

        if (newExpanded.has(id)) {
            newExpanded.delete(id);
        } else {
            newExpanded.add(id);
        }
        setExpandedRows(newExpanded);
    };

    // Apply filters and sorting
    const getAssigneeName = (item: ISubTask): string => {
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

    const getAssigneeEmail = (item: ISubTask): string | undefined => {
        const assigned = item.TaskAssignedTo;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if (Array.isArray(assigned) && assigned.length > 0) return (assigned[0] as any).EMail || (assigned[0] as any).Email;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if (assigned) return (assigned as any).EMail || (assigned as any).Email;
        return undefined;
    };

    // Recursively get all descendants of a set of parent IDs
    const getDescendants = (parentIds: Set<number>, allTasks: ISubTask[]): ISubTask[] => {
        const directChildren = allTasks.filter(t => t.ParentSubtaskId && parentIds.has(t.ParentSubtaskId));
        if (directChildren.length === 0) return [];

        const childIds = new Set(directChildren.map(c => c.Id));
        return [...directChildren, ...getDescendants(childIds, allTasks)];
    };

    // Recursively get all ancestors of a set of task IDs
    const getAncestors = (taskIds: Set<number>, allTasks: ISubTask[]): ISubTask[] => {
        const ancestors: ISubTask[] = [];
        const taskMap = new Map<number, ISubTask>();
        allTasks.forEach(t => taskMap.set(t.Id, t));

        const queue: number[] = [];
        taskIds.forEach(id => queue.push(id));
        const visited = new Set<number>();

        while (queue.length > 0) {
            const currentId = queue.shift()!;
            if (visited.has(currentId)) continue;
            visited.add(currentId);

            const task = taskMap.get(currentId);
            if (task && task.ParentSubtaskId) {
                const parent = taskMap.get(task.ParentSubtaskId);
                if (parent) {
                    ancestors.push(parent);
                    queue.push(parent.Id);
                }
            }
        }
        return ancestors;
    };


    // Helper to get ALL tasks matching current filters (ignoring collapse state)
    const getRawFilteredTasks = (): ISubTask[] => {
        let filtered: ISubTask[] = tasks;

        if (statusFilter) filtered = filtered.filter(t => (t.TaskStatus || 'Not Started') === statusFilter);
        if (categoryFilter) filtered = filtered.filter(t => (t.Category || 'Unknown') === categoryFilter);
        if (overdueFilter) filtered = filtered.filter(t => isOverdue(t));

        // Default Filter: Hide Completed unless explicitly shown OR filtered by "Completed" status
        if (!showCompleted && statusFilter !== 'Completed') {
            filtered = filtered.filter(t => t.TaskStatus !== 'Completed');
        }

        // Filter by Search Term (Name, Desc, Assignee)
        if (searchTerm) {
            const lowerTerm = searchTerm.toLowerCase();
            const matches: ISubTask[] = [];
            const matchIds = new Set<number>();

            // 1. Find direct matches
            for (const t of filtered) {
                const assigneeName = getAssigneeName(t).toLowerCase();
                const isMatch = (
                    (t.Task_Title && t.Task_Title.toLowerCase().indexOf(lowerTerm) !== -1) ||
                    (t.Task_Description && t.Task_Description.toLowerCase().indexOf(lowerTerm) !== -1) ||
                    assigneeName.indexOf(lowerTerm) !== -1
                );
                if (isMatch) {
                    matches.push(t);
                    matchIds.add(t.Id);
                }
            }

            // 2. Find descendants of matches
            const descendants = getDescendants(matchIds, tasks); // Search against ALL tasks to find children

            // Combine and unique
            const uniqueResults = new Map<number, ISubTask>();
            matches.forEach(m => uniqueResults.set(m.Id, m));
            descendants.forEach(d => uniqueResults.set(d.Id, d));

            filtered = [];
            uniqueResults.forEach(task => filtered.push(task));
        }
        return filtered;
    };

    const getFilteredTasks = (): IUiSubTask[] => {
        const filtered = getRawFilteredTasks();

        // 1. If NO filters/sort, return standard hierarchy loop
        if (!statusFilter && !categoryFilter && !overdueFilter && !searchTerm && !sortedColumn) {

            const visibleItems: IUiSubTask[] = [];
            const visibleParents = new Set<number>();
            visibleParents.add(0); // Virtual root

            for (const item of uiSubTasks) {
                // VISIBILITY CHECK 1: Respect "Show Completed" in Hierarchy
                if (!showCompleted && item.TaskStatus === 'Completed') continue;

                const parentId = item.ParentSubtaskId || 0;

                // VISIBILITY CHECK 2: Hierarchy Expansion
                if (item.depth === 0) {
                    visibleItems.push(item);
                    if (expandedRows.has(item.Id)) {
                        visibleParents.add(item.Id);
                    }
                } else {
                    if (visibleParents.has(parentId)) {
                        visibleItems.push(item);
                        if (expandedRows.has(item.Id)) {
                            visibleParents.add(item.Id);
                        }
                    }
                }
            }
            return visibleItems;
        }

        // 2. If SEARCH is active (and no Sort), Preserve Hierarchy!
        if (searchTerm && !sortedColumn) {
            // 'filtered' contains all matches + descendants, but flat.
            // We want to show them in their Tree order with valid depth.

            // Add Ancestors so the tree structure is visible
            const matchIds = new Set(filtered.map(t => t.Id));
            const ancestors = getAncestors(matchIds, tasks);

            const validIds = new Set(filtered.map(t => t.Id));
            ancestors.forEach(a => validIds.add(a.Id));

            // Filter uiSubTasks (the master tree) to keep only valid items
            return uiSubTasks.filter(t => validIds.has(t.Id)).map(t => ({
                ...t,
                // We keep depth!
                // Ensure expansion so user sees the matches
                isExpanded: true
            }));
        }

        // 3. Otherwise (Status Filter Only, or Sort active) -> FLAT VIEW
        let result = filtered.map(t => {
            // Check if this task has ANY children in the full dataset
            const hasKids = tasks.some(child => child.ParentSubtaskId === t.Id);
            return { ...t, depth: 0, hasChildren: hasKids };
        });

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
    const overdueCount = tasks.filter(t => isOverdue(t)).length;

    const onStatusChange = async (item: ISubTask, option: IDropdownOption) => {
        if (item.TaskStatus === 'Completed') return; // Prevent changing completed tasks

        // Validate Admin_Job_ID exists
        if (!item.Admin_Job_ID) {
            setMessage('Error: Cannot update status - missing parent task reference. Please refresh.');
            return;
        }

        try {
            // Optimistic update
            const newTasks = tasks.map(t => t.Id === item.Id ? { ...t, TaskStatus: option.key as string } : t);
            setTasks(newTasks);

            await taskService.updateSubTaskStatus(item.Id, item.Admin_Job_ID, option.key as string);

            // Reload tasks to reflect the change
            await loadTasks();
        } catch (e: any) {
            console.error(e);

            // Check if error is due to incomplete children
            if (e.message === 'INCOMPLETE_CHILDREN' && e.incompleteTasks) {
                setIncompleteTasks(e.incompleteTasks);
                setPendingStatusChange({ item, option });
                setShowIncompleteDialog(true);
                loadTasks(); // Revert optimistic update
                return;
            }

            setMessage('Error updating status: ' + (e.message || e));
            loadTasks(); // Revert on error
        }
    };



    const openEditPanel = async (item: ISubTask) => {
        setSelectedSubtask(item);
        setEditStatus(item.TaskStatus || 'Not Started');
        setEditRemarks(item.User_Remarks || '');
        setNewAttachFiles([]);
        setMessage(undefined);

        // Fetch clarification history
        if (lastTaskIdRef.current !== item.Id) {
            setClarificationHistory([]);
            lastTaskIdRef.current = item.Id;
        }

        setLoadingClarification(true);
        setClarificationMessage('');
        try {
            const history = await taskService.getCorrespondenceByTaskId(item.Admin_Job_ID, item.Id);
            setClarificationHistory(prev => {
                // If SPO is lagging and returns older history, keep our local optimistic state
                if (history.length < prev.length && prev.length > 0) {
                    console.log("[Correspondence] SubView: SPO lag, keeping state");
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
        setSelectedSubtask(null);
        setEditStatus('');
        setEditRemarks('');
        setNewAttachFiles([]);
        setMessage(undefined);
    };

    const handleSaveSubtask = async () => {
        if (!selectedSubtask) return;
        setSaving(true);
        setMessage(undefined);
        try {
            // Update status and remarks
            await taskService.updateSubTaskStatus(selectedSubtask.Id, selectedSubtask.Admin_Job_ID, editStatus, editRemarks);

            // Upload new attachments if any
            if (newAttachFiles.length > 0) {
                await taskService.addAttachmentsToItem(LIST_SUB_TASKS, selectedSubtask.Id, newAttachFiles);
            }

            closeEditPanel();
            loadTasks(); // Reload to get updated data
        } catch (e: any) {
            console.error(e);

            // Check if error is due to incomplete children
            if (e.message === 'INCOMPLETE_CHILDREN' && e.incompleteTasks) {
                setIncompleteTasks(e.incompleteTasks);
                setPendingStatusChange({
                    item: selectedSubtask,
                    option: { key: editStatus, text: editStatus } as IDropdownOption,
                    remarks: editRemarks
                });
                setShowIncompleteDialog(true);
                return;
            }

            setMessage('Error saving: ' + (e.message || e));
        } finally {
            setSaving(false);
        }
    };

    const handleSendClarification = async () => {
        if (!selectedSubtask || !clarificationMessage.trim()) return;
        setSaving(true);
        try {
            // Smart Reply Logic:
            // 1. If there's conversation history, reply to the LAST SENDER
            // 2. Otherwise, send to Main Task Author (the requester)
            // 3. If user IS the author, send to assignee as fallback
            const mainTask = await taskService.getMainTaskById(selectedSubtask.Admin_Job_ID);
            const authorEmail = (mainTask as any)?.Author?.EMail || '';
            const assigneeEmail = getAssigneeEmail(selectedSubtask) || '';


            let toEmail = authorEmail; // Default to main task author

            // Check if there's existing conversation - reply to last sender
            if (clarificationHistory.length > 0) {
                const lastMessage = clarificationHistory[clarificationHistory.length - 1];
                const lastSender = lastMessage.FromAddress || '';

                // Only reply to last sender if it's NOT the current user (avoid self-messaging)
                if (lastSender && lastSender.toLowerCase() !== props.userEmail.toLowerCase()) {
                    toEmail = lastSender;
                    console.log(`[Clarification] Replying to last sender: ${lastSender}`);
                }
            }

            // Fallback: If user IS the author and no conversation history, send to assignee
            if (!toEmail || (toEmail.toLowerCase() === props.userEmail.toLowerCase() && clarificationHistory.length === 0)) {
                toEmail = assigneeEmail;
            }

            const tempMsg = {
                FromAddress: props.userEmail,
                MessageBody: clarificationMessage,
                Created: new Date().toISOString(),
                Author: { Title: 'You' }
            };
            setClarificationHistory(prev => [...prev, tempMsg]);
            const currentMsg = clarificationMessage;
            setClarificationMessage('');

            // Smart subject line based on conversation state
            const subject = clarificationHistory.length > 0
                ? `Reply: ${selectedSubtask.Task_Title}`
                : `Clarification Needed: ${selectedSubtask.Task_Title}`;

            await taskService.sendEmail(
                toEmail ? [toEmail] : [],
                subject,
                currentMsg,
                selectedSubtask.Admin_Job_ID,
                selectedSubtask.Id
            );

            // Fetch actual history after a short delay
            setTimeout(async () => {
                const history = await taskService.getCorrespondenceByTaskId(selectedSubtask.Admin_Job_ID, selectedSubtask.Id);
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
                next.set(selectedSubtask.Id, { hasCorrespondence: true, isReply: false }); // Reset reply flag since we just replied
                return next;
            });
        } catch (e) {
            console.error("Error sending clarification:", e);
        } finally {
            setSaving(false);
        }
    };
    const handleCreateSubtask = async () => {
        if (!newTitle.trim()) {
            setMessage('Error: Title is required.');
            return;
        }
        if (!newDesc.trim()) {
            setMessage('Error: Description is required.');
            return;
        }
        if (!selectedUserKey) {
            setMessage('Error: "Assign To" is required.');
            return;
        }
        if (!dueDate) {
            setMessage('Error: Due Date is required.');
            return;
        }
        if (!addingToParentId) {
            setMessage('Error: Parent task must be selected.');
            return;
        }

        setSaving(true);
        setMessage('');
        // Find parent to get Admin_Job_ID
        // Use filter instead of find for compatibility
        const parentTasks = tasks.filter(t => t.Id === addingToParentId);
        const parentTask = parentTasks.length > 0 ? parentTasks[0] : null;

        if (!parentTask) {
            setMessage('Error: Parent task not found.');
            setSaving(false);
            return;
        }

        // Ensure Admin_Job_ID exists (it should be loaded from the tasks array)
        if (!parentTask.Admin_Job_ID) {
            setMessage('Error: Parent task is missing Admin_Job_ID. Please refresh and try again.');
            setSaving(false);
            return;
        }

        // Date Validation
        const mainTask = mainTaskMap.get(parentTask.Admin_Job_ID);
        if (dueDate && mainTask && mainTask.TaskDueDate) {
            const subTaskDue = new Date(dueDate.getTime());
            const mainTaskDue = new Date(mainTask.TaskDueDate);
            subTaskDue.setHours(0, 0, 0, 0);
            mainTaskDue.setHours(0, 0, 0, 0);

            if (subTaskDue > mainTaskDue) {
                setMessage(`Due Date validation failed: Subtask due date cannot be later than Main Task due date.`);
                setSaving(false);
                return;
            }
        }

        setSaving(true);
        setMessage(undefined);

        try {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const newSubTask: any = {
                Title: newTitle, // Mandatory for SP item
                Task_Title: newTitle,
                Task_Description: newDesc,
                TaskStatus: 'Not Started',
                Admin_Job_ID: parentTask.Admin_Job_ID, // Link to same Main Task
                ParentSubtaskId: addingToParentId, // Link to Parent Subtask
                Category: newCategory,
                TaskDueDate: dueDate ? dueDate.toISOString() : undefined,
                Task_Created_Date: new Date().toISOString(),
                TaskAssignedToId: selectedUserKey ? (selectedUserKey as number) : undefined
            };

            const newId = await taskService.createSubTask(newSubTask, createAttachments);

            // Reload and close
            await loadTasks();
            setIsAdding(false);
            setAddingToParentId(undefined);
            setNewTitle('');
            setNewDesc('');
            setCreateAttachments([]);
            setMessage('Subtask created successfully!');
        } catch (e: any) {
            console.error(e);
            setMessage('Error creating subtask: ' + (e.message || e));
        } finally {
            setSaving(false);
        }
    };

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
        } else {
            setSelectedUserKey(undefined);
        }
    };

    // Handler to force complete all children
    const handleForceCompleteChildren = async () => {
        if (!pendingStatusChange) return;

        setSaving(true);
        setShowIncompleteDialog(false);

        try {
            // Call update with forceComplete flag
            await taskService.updateSubTaskStatus(
                pendingStatusChange.item.Id,
                pendingStatusChange.item.Admin_Job_ID,
                pendingStatusChange.option.key as string,
                pendingStatusChange.remarks,
                true // forceComplete = true
            );

            // Upload new attachments if any (from edit panel)
            if (selectedSubtask && newAttachFiles.length > 0) {
                await taskService.addAttachmentsToItem(LIST_SUB_TASKS, selectedSubtask.Id, newAttachFiles);
            }

            setMessage('Task and all children completed successfully!');
            closeEditPanel();
            await loadTasks();
        } catch (e: any) {
            console.error(e);
            setMessage('Error completing tasks: ' + (e.message || e));
        } finally {
            setSaving(false);
            setPendingStatusChange(null);
            setIncompleteTasks([]);
        }
    };

    // Handler to cancel the completion
    const handleCancelCompletion = () => {
        setShowIncompleteDialog(false);
        setPendingStatusChange(null);
        setIncompleteTasks([]);
        setMessage(undefined);
    };


    const resetFilters = () => {
        setStatusFilter(undefined);
        setCategoryFilter(undefined);
        setOverdueFilter(false);
        setSearchTerm('');
        setShowCompleted(false);
    };

    const showAllTasks = () => {
        setStatusFilter(undefined);
        setCategoryFilter(undefined);
        setOverdueFilter(false);
        setSearchTerm('');
        setShowCompleted(true);
    };

    const exportToExcel = () => {
        const itemsToExport = getRawFilteredTasks();
        if (itemsToExport.length === 0) {
            setMessage('No items to export.');
            return;
        }

        const csvRows = [];
        const headers = ['ID', 'Title', 'Description', 'Status', 'Category', 'Due Date', 'Assigned To'];
        csvRows.push(headers.join(','));

        for (const item of itemsToExport) {
            const assignee = getAssigneeName(item).replace(/,/g, ''); // Remove commas for CSV
            const values = [
                item.Id,
                `"${(item.Task_Title || '').replace(/"/g, '""')}"`,
                `"${(item.Task_Description || '').replace(/"/g, '""')}"`,
                item.TaskStatus,
                item.Category,
                item.TaskDueDate ? new Date(item.TaskDueDate).toLocaleDateString() : '',
                `"${assignee}"`
            ];
            csvRows.push(values.join(','));
        }

        const csvString = csvRows.join('\n');
        const blob = new Blob([csvString], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.setAttribute('hidden', '');
        a.setAttribute('href', url);
        a.setAttribute('download', 'Tasks_Export.csv');
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    };

    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        const newIsSortedDescending = column.key === sortedColumn ? !isSortedDescending : false;
        setSortedColumn(column.key);
        setIsSortedDescending(newIsSortedDescending);
    };



    // --- STYLES & ANIMATIONS ---
    const theme = React.useMemo(() => ({
        palette: {
            themePrimary: '#0078d4',
            themeLighterAlt: '#eff6fc',
            neutralLighter: '#f3f2f1',
            textSecondary: '#605e5c'
        }
    }), []);

    const dashboardStyles = {
        container: {
            animation: 'fadeIn 0.5s ease-in-out',
            padding: '20px',
            backgroundColor: '#faf9f8', // Very light grey bg for contrast
            minHeight: '100%'
        },
        headerTitle: {
            fontSize: '24px',
            fontWeight: 700,
            color: '#201F1E',
            marginBottom: '4px'
        },
        headerSubtitle: {
            fontSize: '14px',
            color: '#605e5c',
            marginBottom: '20px'
        },
        card: {
            backgroundColor: 'white',
            borderRadius: '8px',
            padding: '20px',
            boxShadow: '0 1.6px 3.6px 0 rgba(0,0,0,0.132), 0 0.3px 0.9px 0 rgba(0,0,0,0.108)',
            marginBottom: '20px',
            transition: 'transform 0.2s, box-shadow 0.2s',
            ':hover': {
                transform: 'translateY(-2px)',
                boxShadow: '0 6.4px 14.4px 0 rgba(0,0,0,0.132), 0 1.2px 3.6px 0 rgba(0,0,0,0.108)'
            }
        },
        summaryCard: {
            minWidth: '200px',
            flex: 1,
            cursor: 'pointer',
            borderLeft: '4px solid transparent' // For colorful accents
        }
    };

    // Inject fade-in keyframes
    React.useEffect(() => {
        const styleSheet = document.createElement("style");
        styleSheet.innerText = `
            @keyframes fadeIn {
                from { opacity: 0; transform: translateY(10px); }
                to { opacity: 1; transform: translateY(0); }
            }
            .task-row-hover:hover {
                background-color: #f3f2f1 !important;
                transition: background-color 0.2s;
            }
            .task-row-hover:hover .show-on-row-hover {
                opacity: 1 !important;
            }
        `;
        document.head.appendChild(styleSheet);
        return () => {
            document.head.removeChild(styleSheet);
        };
    }, []);

    const columns: IColumn[] = [
        {
            key: 'view', name: 'View', minWidth: 50, maxWidth: 60,
            onRender: (item: ISubTask) => (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }}>
                    <IconButton
                        iconProps={{ iconName: 'RedEye' }}
                        title="View Details"
                        ariaLabel="View Details"
                        styles={{
                            root: {
                                color: '#0078d4',
                                fontSize: 14,
                                backgroundColor: '#f3f2f1',
                                borderRadius: '4px',
                                height: 32,
                                width: 32
                            },
                            rootHovered: {
                                backgroundColor: '#e1dfdd',
                                color: '#005a9e'
                            },
                            icon: { fontSize: 14, fontWeight: 600 }
                        }}
                        onClick={(e) => { e.stopPropagation(); openEditPanel(item); }}
                    />
                </div>
            )
        },
        {
            key: 'Task_Title', name: 'Task Name', fieldName: 'Task_Title', minWidth: 180, maxWidth: 300, isResizable: true, isSorted: sortedColumn === 'Task_Title', isSortedDescending, onColumnClick,
            styles: { root: { fontWeight: 600 } }, // Bold Header
            onRender: (item: IUiSubTask) => {
                const indent = item.depth * 24;
                const isExpanded = expandedRows.has(item.Id);
                return (
                    <div style={{ display: 'flex', alignItems: 'center', paddingLeft: indent }}>
                        {item.hasChildren ? (
                            <IconButton
                                iconProps={{ iconName: isExpanded ? 'CaretDownSolid8' : 'CaretRightSolid8' }} // Distinct from View Chevron
                                styles={{ root: { height: 20, width: 20, marginRight: 8, color: '#605e5c' } }}
                                onClick={(e) => { e.stopPropagation(); toggleExpand(item.Id); }}
                            />
                        ) : (
                            <div style={{ width: 28 }} />
                        )}
                        <span
                            title={item.Task_Title}
                            style={{
                                fontWeight: item.hasChildren ? 700 : 500,
                                fontSize: '14px',
                                color: '#201F1E'
                            }}
                        >
                            {item.Task_Title}
                        </span>
                        {item.User_Remarks && item.User_Remarks.indexOf('[WF_NODE:') !== -1 && (
                            <Icon
                                iconName="BranchMerge"
                                title="Workflow Node"
                                styles={{ root: { marginLeft: 8, color: '#0078d4', fontSize: 14, fontWeight: 700 } }}
                            />
                        )}
                        {(() => {
                            const mainTask = mainTaskMap.get(item.Admin_Job_ID);
                            const isMainTaskCompleted = mainTask?.Status === 'Completed';

                            if (!isMainTaskCompleted) { // Removed item.TaskStatus !== 'Completed' check
                                return (
                                    <IconButton
                                        iconProps={{ iconName: 'Add' }}
                                        title="Add Sub-subtask"
                                        styles={{
                                            root: {
                                                height: 26,
                                                width: 26,
                                                marginLeft: 12,
                                                opacity: 0,
                                                borderRadius: '50%',
                                                background: 'linear-gradient(135deg, #0078d4 0%, #00bcf2 100%)',
                                                color: 'white',
                                                boxShadow: '0 2px 5px rgba(0,0,0,0.2)',
                                                transition: 'all 0.2s cubic-bezier(0.4, 0, 0.2, 1)',
                                                transform: 'scale(0.9)'
                                            },
                                            rootHovered: {
                                                transform: 'scale(1.1) rotate(90deg)',
                                                boxShadow: '0 4px 8px rgba(0, 120, 212, 0.4)',
                                                background: 'linear-gradient(135deg, #005a9e 0%, #0078d4 100%)',
                                                color: 'white' // Keep white on hover
                                            },
                                            rootPressed: {
                                                transform: 'scale(0.95)',
                                                background: '#004578'
                                            },
                                            icon: {
                                                fontSize: 14,
                                                fontWeight: 700,
                                                lineHeight: 26
                                            }
                                        }}
                                        className="show-on-row-hover"
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            setAddingToParentId(item.Id);
                                            setIsAdding(true);
                                            setNewTitle('');
                                            setNewDesc('');
                                            setNewCategory(item.Category || '');
                                            setSelectedUserKey(undefined);
                                            setMessage(undefined);
                                            if (!expandedRows.has(item.Id)) toggleExpand(item.Id);
                                        }}
                                    />
                                );
                            }
                            return null;
                        })()}
                    </div>
                );
            }
        },
        { key: 'Task_Description', name: 'Description', fieldName: 'Task_Description', minWidth: 150, maxWidth: 250, isResizable: true, isSorted: sortedColumn === 'Task_Description', isSortedDescending, onColumnClick },
        {
            key: 'assigned', name: 'Assigned To', minWidth: 120, isResizable: true,
            onRender: (i: ISubTask) => {
                const email = getAssigneeEmail(i);
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Persona
                            text={getAssigneeName(i)}
                            imageUrl={email ? `/_layouts/15/userphoto.aspx?size=S&username=${email}` : undefined}
                            size={PersonaSize.size24}
                            hidePersonaDetails={true}
                            styles={{ root: { cursor: 'default' } }}
                        />
                        <span>{getAssigneeName(i)}</span>
                    </div>
                );
            }
        },
        { key: 'Task_Created_Date', name: 'Created', minWidth: 90, onRender: (i: ISubTask) => formatDate(i.Task_Created_Date), isSorted: sortedColumn === 'Task_Created_Date', isSortedDescending, onColumnClick },
        { key: 'TaskDueDate', name: 'Due Date', minWidth: 90, isResizable: true, onRender: (item: ISubTask) => formatDate(item.TaskDueDate), isSorted: sortedColumn === 'TaskDueDate', isSortedDescending, onColumnClick },
        { key: 'Task_End_Date', name: 'End Date', minWidth: 90, onRender: (i: ISubTask) => formatDate(i.Task_End_Date), isSorted: sortedColumn === 'Task_End_Date', isSortedDescending, onColumnClick },
        {
            key: 'Category', name: 'Category', fieldName: 'Category', minWidth: 100, isResizable: true, isSorted: sortedColumn === 'Category', isSortedDescending, onColumnClick,
            onRender: (item: ISubTask) => (
                <span style={{
                    padding: '4px 10px',
                    borderRadius: '12px',
                    backgroundColor: '#f3f2f1',
                    fontSize: '12px',
                    fontWeight: 600,
                    color: '#605e5c'
                }}>
                    {item.Category || 'General'}
                </span>
            )
        },
        {
            key: 'status', name: 'Status', minWidth: 130, isSorted: sortedColumn === 'TaskStatus', isSortedDescending, onColumnClick,
            onRender: (item: ISubTask) => {
                const isItemCompleted = item.TaskStatus === 'Completed';
                return (
                    <Dropdown
                        selectedKey={item.TaskStatus}
                        options={[
                            { key: 'Not Started', text: 'Not Started' },
                            { key: 'In Progress', text: 'In Progress' },
                            { key: 'Completed', text: 'Completed' },
                            { key: 'On Hold', text: 'On Hold' }
                        ]}
                        onChange={(e, o) => o && onStatusChange(item, o)}
                        disabled={isItemCompleted}
                        styles={{
                            root: { width: '100%' },
                            title: {
                                border: 'none',
                                backgroundColor: 'transparent',
                                fontWeight: 600,
                                color: isItemCompleted ? '#107c10' : undefined
                            },
                            dropdown: {
                                ':hover .ms-Dropdown-title': { backgroundColor: '#f3f2f1', border: '1px solid #edebe9' }
                            }
                        }}
                    />
                );
            }
        },
        { key: 'User_Remarks', name: 'Remarks', minWidth: 100, isResizable: true, onRender: (i: ISubTask) => i.User_Remarks || '-', isSorted: sortedColumn === 'User_Remarks', isSortedDescending, onColumnClick },
        {
            key: 'attachments', name: 'Attachments', minWidth: 100, onRender: (item: ISubTask) => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const files = (item as any).AttachmentFiles;
                if (!files || files.length === 0) return <span style={{ color: '#d0d0d0' }}>-</span>;
                return (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                        {/* eslint-disable-next-line @typescript-eslint/no-explicit-any */}
                        {files.map((f: any, idx: number) => (
                            <a key={idx} href={f.ServerRelativeUrl} target="_blank" rel="noopener noreferrer" style={{ display: 'flex', alignItems: 'center', fontSize: 12, color: '#0078d4', textDecoration: 'none' }}>
                                <Icon iconName="Attach" styles={{ root: { fontSize: 10, marginRight: 4 } }} /> {f.FileName}
                            </a>
                        ))}
                    </div>
                );
            }
        }
    ];

    if (loading) return (
        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '50vh', flexDirection: 'column', gap: 20 }}>
            {/* Simple Loader Animation */}
            <div style={{ width: 40, height: 40, border: '4px solid #f3f3f3', borderTop: '4px solid #0078d4', borderRadius: '50%', animation: 'spin 1s linear infinite' }} />
            <div style={{ color: '#605e5c' }}>Loading your dashboard...</div>
            <style>{`@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }`}</style>
        </div>
    );

    const statusCounts: { [key: string]: { count: number; color: string; icon: string } } = {
        'Not Started': { count: tasks.filter(t => (t.TaskStatus || 'Not Started') === 'Not Started').length, color: '#6c757d', icon: 'CircleRing' },
        'In Progress': { count: tasks.filter(t => t.TaskStatus === 'In Progress').length, color: '#0078d4', icon: 'Sync' },
        'Completed': { count: tasks.filter(t => t.TaskStatus === 'Completed').length, color: '#107c10', icon: 'CheckMark' },
        'On Hold': { count: tasks.filter(t => t.TaskStatus === 'On Hold').length, color: '#ff8c00', icon: 'Pause' }
    };

    return (
        <div style={dashboardStyles.container}>
            {/* Header Section - Removed per user request */}
            <div style={{ marginBottom: 32 }}>

                {/* Summary Cards with Gradients */}
                <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap' }}>
                    {/* Total Card with Gradient */}
                    <div
                        style={{
                            ...dashboardStyles.card,
                            ...dashboardStyles.summaryCard,
                            background: 'linear-gradient(135deg, #0078d4 0%, #106ebe 100%)',
                            color: 'white',
                            border: 'none',
                            boxShadow: '0 4px 12px rgba(0, 120, 212, 0.3)',
                            transition: 'all 0.3s ease',
                            cursor: 'pointer',
                            animation: 'cardSlideIn 0.5s ease-out 0s both'
                        }}
                        onClick={showAllTasks}
                        onMouseEnter={(e) => {
                            e.currentTarget.style.transform = 'translateY(-4px)';
                            e.currentTarget.style.boxShadow = '0 8px 24px rgba(0, 120, 212, 0.4)';
                        }}
                        onMouseLeave={(e) => {
                            e.currentTarget.style.transform = 'translateY(0)';
                            e.currentTarget.style.boxShadow = '0 4px 12px rgba(0, 120, 212, 0.3)';
                        }}
                    >
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                            <div>
                                <div style={{ fontSize: '12px', fontWeight: 600, opacity: 0.9, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Total Tasks</div>
                                <div style={{ fontSize: '36px', fontWeight: 700, marginTop: 8, lineHeight: 1 }}>{tasks.length}</div>
                                <div style={{ fontSize: '12px', opacity: 0.8, marginTop: 8 }}>All assigned tasks</div>
                            </div>
                            <Icon iconName="BulletedList" styles={{ root: { fontSize: 32, opacity: 0.3 } }} />
                        </div>
                    </div>

                    {/* Status Cards with Gradients */}
                    {Object.keys(statusCounts).map((status, index) => {
                        const gradients: { [key: string]: string } = {
                            'Not Started': 'linear-gradient(135deg, #6c757d 0%, #5a6268 100%)',
                            'In Progress': 'linear-gradient(135deg, #0078d4 0%, #106ebe 100%)',
                            'Completed': 'linear-gradient(135deg, #107c10 0%, #0b6a0b 100%)',
                            'On Hold': 'linear-gradient(135deg, #ff8c00 0%, #d77d00 100%)'
                        };

                        return (
                            <div
                                key={status}
                                style={{
                                    ...dashboardStyles.card,
                                    ...dashboardStyles.summaryCard,
                                    background: gradients[status],
                                    color: 'white',
                                    border: 'none',
                                    boxShadow: `0 4px 12px ${statusCounts[status].color}40`,
                                    transition: 'all 0.3s ease',
                                    cursor: 'pointer',
                                    opacity: statusFilter && statusFilter !== status ? 0.6 : 1,
                                    animation: `cardSlideIn 0.5s ease-out ${(index + 1) * 0.1}s both`
                                }}
                                onClick={() => setStatusFilter(statusFilter === status ? undefined : status)}
                                onMouseEnter={(e) => {
                                    e.currentTarget.style.transform = 'translateY(-4px) scale(1.02)';
                                    e.currentTarget.style.boxShadow = `0 8px 24px ${statusCounts[status].color}60`;
                                }}
                                onMouseLeave={(e) => {
                                    e.currentTarget.style.transform = 'translateY(0) scale(1)';
                                    e.currentTarget.style.boxShadow = `0 4px 12px ${statusCounts[status].color}40`;
                                }}
                            >
                                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                                    <div>
                                        <div style={{ fontSize: '12px', fontWeight: 600, opacity: 0.9, textTransform: 'uppercase' }}>{status}</div>
                                        <div style={{ fontSize: '36px', fontWeight: 700, marginTop: 8, lineHeight: 1 }}>{statusCounts[status].count}</div>
                                        <div style={{ fontSize: '11px', opacity: 0.8, marginTop: 8 }}>
                                            {((statusCounts[status].count / (tasks.length || 1)) * 100).toFixed(0)}% of total
                                        </div>
                                    </div>
                                    <Icon iconName={statusCounts[status].icon} styles={{ root: { fontSize: 28, opacity: 0.3 } }} />
                                </div>
                            </div>
                        );
                    })}

                    {/* Overdue Card with Gradient */}
                    {overdueCount > 0 && (
                        <div
                            style={{
                                ...dashboardStyles.card,
                                ...dashboardStyles.summaryCard,
                                background: 'linear-gradient(135deg, #d13438 0%, #a4262c 100%)',
                                color: 'white',
                                border: 'none',
                                boxShadow: '0 4px 12px rgba(209, 52, 56, 0.3)',
                                transition: 'all 0.3s ease',
                                cursor: 'pointer',
                                animation: 'cardSlideIn 0.5s ease-out 0.5s both, pulse 2s ease-in-out infinite'
                            }}
                            onClick={() => setOverdueFilter(!overdueFilter)}
                            onMouseEnter={(e) => {
                                e.currentTarget.style.transform = 'translateY(-4px) scale(1.02)';
                                e.currentTarget.style.boxShadow = '0 8px 24px rgba(209, 52, 56, 0.5)';
                            }}
                            onMouseLeave={(e) => {
                                e.currentTarget.style.transform = 'translateY(0) scale(1)';
                                e.currentTarget.style.boxShadow = '0 4px 12px rgba(209, 52, 56, 0.3)';
                            }}
                        >
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                                <div>
                                    <div style={{ fontSize: '12px', fontWeight: 600, opacity: 0.9, textTransform: 'uppercase' }}>Attention Needed</div>
                                    <div style={{ fontSize: '36px', fontWeight: 700, marginTop: 8, lineHeight: 1 }}>{overdueCount}</div>
                                    <div style={{ fontSize: '11px', opacity: 0.8, marginTop: 8 }}>Overdue tasks</div>
                                </div>
                                <Icon iconName="Warning" styles={{ root: { fontSize: 28, opacity: 0.3 } }} />
                            </div>
                        </div>
                    )}
                </div>
            </div>

            {/* Task Progress Overview - Removed */}

            {/* CSS Animations */}
            <style>{`
                @keyframes cardSlideIn {
                    from {
                        opacity: 0;
                        transform: translateY(20px);
                    }
                    to {
                        opacity: 1;
                        transform: translateY(0);
                    }
                }
                @keyframes pulse {
                    0%, 100% {
                        box-shadow: 0 4px 12px rgba(209, 52, 56, 0.3);
                    }
                    50% {
                        box-shadow: 0 4px 20px rgba(209, 52, 56, 0.5);
                    }
                }
                @keyframes pulseBlue {
                    0% {
                        box-shadow: 0 0 0 0 rgba(0, 120, 212, 0.7);
                    }
                    70% {
                        box-shadow: 0 0 0 6px rgba(0, 120, 212, 0);
                    }
                    100% {
                        box-shadow: 0 0 0 0 rgba(0, 120, 212, 0);
                    }
                }
                @keyframes rotateIn {
                    from { transform: rotate(-90deg) scale(0.8); opacity: 0; }
                    to { transform: rotate(0) scale(1); opacity: 1; }
                }
                @keyframes barGrow {
                    from {
                        width: 0%;
                    }
                }
            `}</style>

            {/* Charts Section - Removed */}

            {/* CSS Animations */}
            <style>{`
                @keyframes pieSlideIn {
                    from {
                        opacity: 0;
                        transform: scale(0.8);
                    }
                    to {
                        opacity: 0.9;
                        transform: scale(1);
                    }
                }
            `}</style>

            {/* Filter Bar */}
            <div style={{ ...dashboardStyles.card, padding: '16px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 16 }}>
                <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center" wrap>
                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                        <SearchBox
                            placeholder="Search tasks..."
                            onChange={(_, newValue) => setSearchTerm(newValue || '')}
                            onSearch={(newValue) => setSearchTerm(newValue)}
                            value={searchTerm}
                            styles={{ root: { width: 300 } }}
                        />
                        {hasRecognitionSupport && (
                            <VoiceControl
                                isListening={isListening}
                                onToggleListening={isListening ? stopListening : startListening}
                            />
                        )}
                    </Stack>
                    <div style={{ width: 1, height: 24, backgroundColor: '#edebe9' }} /> {/* Separator */}

                    <Checkbox
                        label="Show Completed"
                        checked={showCompleted}
                        onChange={(_, checked) => setShowCompleted(!!checked)}
                        styles={{ root: { alignItems: 'center' } }}
                    />

                    <Dropdown
                        placeholder="Filter by Status"
                        selectedKey={statusFilter || ''}
                        onChange={(_, opt) => setStatusFilter(opt?.key as string || undefined)}
                        options={statusOptions}
                        styles={{ root: { width: 180 }, dropdown: { border: 'none', backgroundColor: '#f3f2f1' } }}
                    />
                    <Dropdown
                        placeholder="Filter by Category"
                        selectedKey={categoryFilter || ''}
                        onChange={(_, opt) => setCategoryFilter(opt?.key as string || undefined)}
                        options={categoryOptions}
                        styles={{ root: { width: 180 } }}
                    />
                    <Checkbox
                        label="Show Overdue Only"
                        checked={overdueFilter}
                        onChange={(_, v) => setOverdueFilter(!!v)}
                    />
                </Stack>
                <div style={{ display: 'flex', gap: 8 }}>
                    <PrimaryButton
                        iconProps={{ iconName: 'Download' }}
                        text="Export"
                        onClick={exportToExcel}
                    />
                    <DefaultButton
                        iconProps={{ iconName: 'ClearFilter' }}
                        text="Reset Filters"
                        onClick={resetFilters}
                        styles={{ root: { border: 'none', color: '#0078d4' }, rootHovered: { backgroundColor: '#eff6fc' } }}
                    />
                </div>
            </div>

            {/* Data Grid */}
            <div style={{
                backgroundColor: 'white',
                borderRadius: '8px',
                boxShadow: '0 1.6px 3.6px 0 rgba(0,0,0,0.132)',
                overflowX: 'auto', // CRITICAL: Allow horizontal scroll if content is too wide
                border: '1px solid #edebe9',
                minWidth: 0 // CRITICAL: Prevent grid expansion
            }}>
                <div style={{ position: 'relative', height: '65vh' }}>
                    <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                        <DetailsList
                            items={filteredTasks}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            constrainMode={ConstrainMode.unconstrained}
                            layoutMode={DetailsListLayoutMode.fixedColumns}
                            compact={false} // Comfortable spacing
                            onRenderDetailsHeader={
                                (props: IDetailsHeaderProps | undefined, defaultRender?: (props: IDetailsHeaderProps) => JSX.Element | null): JSX.Element | null => {
                                    if (!props || !defaultRender) return null;
                                    return (
                                        <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                                            {/* Custom Header Start */}
                                            <div style={{ backgroundColor: 'white', borderBottom: '1px solid #edebe9', zIndex: 10 }}>
                                                {defaultRender({
                                                    ...props,
                                                    styles: { root: { paddingTop: 12, paddingBottom: 12 } } // Taller header
                                                })}
                                            </div>
                                        </Sticky>
                                    );
                                }
                            }
                            onRenderRow={(props, defaultRender) => {
                                if (!props || !defaultRender) return null;
                                const item = props.item as IUiSubTask;
                                const isOverdueItem = isOverdue(item);

                                // Custom row style
                                return (
                                    <div className="task-row-hover">
                                        {defaultRender({
                                            ...props,
                                            styles: {
                                                root: {
                                                    backgroundColor: isOverdueItem && item.TaskStatus !== 'Completed' ? '#fff0f0' : 'white',
                                                    borderBottom: '1px solid #f3f2f1',
                                                    fontSize: '14px'
                                                },
                                                cell: { fontSize: '14px', alignItems: 'center', display: 'flex' }
                                            }
                                        })}
                                    </div>
                                );
                            }}
                        />
                    </ScrollablePane>
                </div>
            </div>
            {filteredTasks.length === 0 && (
                <div style={{ padding: 40, textAlign: 'center', color: '#605e5c', backgroundColor: 'white', marginTop: -20, borderRadius: '0 0 8px 8px' }}>
                    <Icon iconName="CheckList" styles={{ root: { fontSize: 48, color: '#e1dfdd', marginBottom: 16 } }} />
                    <div style={{ fontSize: 18, fontWeight: 600 }}>No tasks found</div>
                    <div>Try adjusting your filters or search criteria.</div>
                </div>
            )}

            {/* Edit Subtask Panel */}
            <Panel
                isOpen={!!selectedSubtask}
                onDismiss={closeEditPanel}
                type={PanelType.medium}
                headerText="Edit Subtask"
            >
                {selectedSubtask && (() => {
                    const isCompleted = selectedSubtask.TaskStatus === 'Completed';
                    return (
                        <Stack tokens={{ childrenGap: 15 }} styles={{ root: { padding: 10 } }}>
                            {isCompleted && (
                                <MessageBar messageBarType={MessageBarType.warning}>
                                    This subtask is completed and cannot be edited.
                                </MessageBar>
                            )}
                            <div>
                                <strong>Title:</strong> {selectedSubtask.Task_Title}
                            </div>
                            <div>
                                <strong>Description:</strong> {selectedSubtask.Task_Description || 'N/A'}
                            </div>
                            <div>
                                <strong>Assigned To:</strong> {getAssigneeName(selectedSubtask)}
                            </div>
                            <div>
                                <strong>Due Date:</strong> {formatDate(selectedSubtask.TaskDueDate)}
                            </div>
                            <div>
                                <strong>Category:</strong> {selectedSubtask.Category || 'N/A'}
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
                                onChange={(_, opt) => opt && setEditStatus(opt.key as string)}
                                disabled={isCompleted}
                            />

                            <TextField
                                label="Remarks"
                                multiline
                                rows={4}
                                value={editRemarks}
                                onChange={(_, v) => setEditRemarks(v || '')}
                                readOnly={isCompleted}
                                disabled={isCompleted}
                            />

                            <div>
                                <strong>Current Attachments:</strong>
                                {/* eslint-disable-next-line @typescript-eslint/no-explicit-any */}
                                {(() => {
                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                    const files = (selectedSubtask as any).AttachmentFiles;
                                    if (!files || files.length === 0) return <div style={{ color: '#999' }}>No attachments</div>;
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

                            {!isCompleted && (
                                <div>
                                    <label style={{ fontWeight: 600, display: 'block', marginBottom: 4 }}>Add New Attachments</label>
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
                                                setNewAttachFiles(files);
                                            } else {
                                                setNewAttachFiles([]);
                                            }
                                        }}
                                    />
                                    {newAttachFiles.length > 0 && (
                                        <div style={{ marginTop: 8, fontSize: 12, color: '#666' }}>
                                            {newAttachFiles.map((f, i) => <div key={i}>{f.name}</div>)}
                                        </div>
                                    )}
                                </div>
                            )}

                            {message && <MessageBar messageBarType={MessageBarType.error}>{message}</MessageBar>}

                            {/* Clarification Section */}
                            <div style={{ borderTop: '1px solid #edebe9', paddingTop: 20, marginTop: 10 }}>
                                <div style={{ fontSize: 16, fontWeight: 600, color: '#323130', marginBottom: 15, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                        <Icon iconName="Questionnaire" /> Task Correspondence
                                    </div>
                                    <IconButton
                                        iconProps={{ iconName: 'Refresh' }}
                                        title="Refresh history"
                                        onClick={async () => {
                                            if (!selectedSubtask) return;
                                            setLoadingClarification(true);
                                            const h = await taskService.getCorrespondenceByTaskId(selectedSubtask.Admin_Job_ID, selectedSubtask.Id);
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
                                        <div style={{ textAlign: 'center', padding: 20, fontSize: 12 }}>Loading history...</div>
                                    ) : clarificationHistory.length === 0 ? (
                                        <div style={{ textAlign: 'center', padding: 20, fontSize: 12, color: '#64748b' }}>No history found.</div>
                                    ) : (
                                        <Stack tokens={{ childrenGap: 10 }}>
                                            {clarificationHistory.map((msg, i) => (
                                                <div key={i} style={{
                                                    padding: '8px 12px',
                                                    background: msg.FromAddress === props.userEmail ? '#eff6ff' : 'white',
                                                    border: '1px solid #e2e8f0',
                                                    borderRadius: 8,
                                                    alignSelf: msg.FromAddress === props.userEmail ? 'flex-end' : 'flex-start',
                                                    maxWidth: '85%'
                                                }}>
                                                    <div style={{ fontSize: 10, color: '#64748b', marginBottom: 4, display: 'flex', justifyContent: 'space-between' }}>
                                                        <strong>{msg.Author?.Title || msg.FromAddress}</strong>
                                                        <span style={{ marginLeft: 8 }}>{new Date(msg.Created).toLocaleString([], { hour: '2-digit', minute: '2-digit' })}</span>
                                                    </div>
                                                    <div dangerouslySetInnerHTML={{ __html: msg.MessageBody }} style={{ fontSize: 13 }} />
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

                            <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                                {!isCompleted && <PrimaryButton text="Save" onClick={handleSaveSubtask} disabled={saving} />}
                                <DefaultButton text={isCompleted ? "Close" : "Cancel"} onClick={closeEditPanel} disabled={saving} />
                            </Stack>
                        </Stack>
                    );
                })()}
            </Panel>


            {/* Create Subtask Panel */}
            <Panel
                isOpen={isAdding}
                onDismiss={() => setIsAdding(false)}
                type={PanelType.medium}
                headerText="Create Sub-subtask"
            >
                <Stack tokens={{ childrenGap: 15 }} styles={{ root: { padding: 10 } }}>
                    <TextField label="Title" required value={newTitle} onChange={(_, v) => setNewTitle(v || '')} />
                    <TextField label="Description" required multiline rows={3} value={newDesc} onChange={(_, v) => setNewDesc(v || '')} />

                    <ComboBox
                        label="Assign To *"
                        required={true}
                        options={userOptions}
                        selectedKey={selectedUserKey}
                        onChange={onUserComboChange}
                        onRenderOption={onRenderUserOption}
                        allowFreeform={true}
                        autoComplete="on"
                    />

                    <DatePicker
                        label="Due Date"
                        isRequired={true}
                        value={dueDate}
                        onSelectDate={(d) => setDueDate(d || undefined)}
                    />




                    <ComboBox
                        label="Category"
                        options={categoryFormOptions}
                        selectedKey={newCategory}
                        onChange={(e, opt) => setNewCategory(opt?.key as string || opt?.text || '')}
                        allowFreeform={true} // Allow new categories if needed, or set to false to enforce
                        autoComplete="on"
                    />

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
                                    setCreateAttachments(files);
                                } else {
                                    setCreateAttachments([]);
                                }
                            }}
                        />
                        {createAttachments.length > 0 && (
                            <div style={{ marginTop: 8, fontSize: 12, color: '#666' }}>
                                {createAttachments.map((f, i) => <div key={i}>{f.name}</div>)}
                            </div>
                        )}
                    </div>

                    {message && <MessageBar messageBarType={message.indexOf('Error') >= 0 ? MessageBarType.error : MessageBarType.success}>{message}</MessageBar>}

                    <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
                        <PrimaryButton text="Create" onClick={handleCreateSubtask} disabled={saving} />
                        <DefaultButton text="Cancel" onClick={() => setIsAdding(false)} disabled={saving} />
                    </Stack>
                </Stack>
            </Panel>

            {/* Incomplete Children Dialog - Enhanced Design */}
            <Dialog
                hidden={!showIncompleteDialog}
                onDismiss={handleCancelCompletion}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: '⚠️ Incomplete Child Tasks',
                    // Use standard title but with emoji for impact
                    // Removed subText to have full control in body
                }}
                modalProps={{
                    isBlocking: true,
                    styles: {
                        main: {
                            width: '900px !important', // Explicit width
                            minWidth: '900px !important',
                            maxWidth: '100vw !important', // Allow expansion
                            borderRadius: 12 // softer corners
                        }
                    }
                }}
            >
                <div style={{ marginTop: 0 }}>
                    {/* Warning Box */}
                    <div style={{
                        backgroundColor: '#fff4ce',
                        padding: '20px',
                        borderRadius: '6px',
                        borderLeft: '6px solid #ff8c00',
                        marginBottom: '24px',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: 8
                    }}>
                        <div style={{ fontSize: '16px', fontWeight: 600, color: '#323130', display: 'flex', alignItems: 'center' }}>
                            <span style={{ marginRight: 10, fontSize: 22 }}>🤔</span> Do you want to complete all child tasks?
                        </div>
                        <div style={{ fontSize: '14px', color: '#605e5c', lineHeight: '20px', paddingLeft: 38 }}>
                            Completing this parent task will also mark all its child tasks as complete. Please review the detailed list below.
                        </div>
                    </div>

                    {/* Incomplete Tasks List */}
                    {incompleteTasks.length > 0 && (
                        <div>
                            {/* Scrollable Container */}
                            <div style={{
                                maxHeight: '400px',
                                overflowY: 'auto',
                                overflowX: 'hidden',
                                padding: '4px' // padding for shadows
                            }}>
                                {incompleteTasks.map((task, index) => {
                                    const getStatusColor = (status: string) => {
                                        // Specific colors from reference image style
                                        switch (status) {
                                            case 'Not Started': return { bg: '#e1dfdd', color: '#201F1E' }; // Grey
                                            case 'In Progress': return { bg: '#CCE1FF', color: '#005A9E' }; // Light Blue
                                            case 'Completed': return { bg: '#DFF6DD', color: '#107C10' }; // Green
                                            default: return { bg: '#F3F2F1', color: '#201F1E' };
                                        }
                                    };

                                    const statusColors = getStatusColor(task.TaskStatus || 'Not Started');

                                    return (
                                        <div
                                            key={task.Id}
                                            style={{
                                                backgroundColor: 'white',
                                                borderRadius: '12px',
                                                border: '1px solid #e1dfdd',
                                                marginBottom: '16px',
                                                padding: '20px',
                                                boxShadow: '0 2px 8px rgba(0,0,0,0.05)', // Softer, more modern shadow
                                                display: 'flex',
                                                gap: '20px',
                                                alignItems: 'center' // Center vertically like the reference? Or top? Ref image has badge somewhat centered.
                                            }}
                                        >
                                            {/* Large Number Badge */}
                                            <div style={{
                                                width: '40px',
                                                height: '40px',
                                                borderRadius: '50%',
                                                backgroundColor: '#ff8c00', // Orange
                                                color: 'white',
                                                display: 'flex',
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                fontSize: '18px',
                                                fontWeight: 'bold',
                                                flexShrink: 0
                                            }}>
                                                {index + 1}
                                            </div>

                                            {/* Task Content */}
                                            <div style={{ flex: 1, minWidth: 0 }}>
                                                {/* Row 1: Title + Status */}
                                                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 8 }}>
                                                    <div style={{
                                                        fontSize: '16px',
                                                        fontWeight: 700,
                                                        color: '#201F1E',
                                                        marginRight: 12,
                                                        flex: 1,
                                                        wordBreak: 'break-word',
                                                        lineHeight: '22px'
                                                    }}>
                                                        {task.Task_Title}
                                                    </div>
                                                    <span style={{
                                                        backgroundColor: statusColors.bg,
                                                        color: statusColors.color,
                                                        padding: '4px 12px',
                                                        borderRadius: '16px', // Pill shape
                                                        fontSize: '12px',
                                                        fontWeight: 600,
                                                        whiteSpace: 'nowrap',
                                                        flexShrink: 0
                                                    }}>
                                                        {task.TaskStatus || 'Not Started'}
                                                    </span>
                                                </div>

                                                {/* Row 2: User + Date */}
                                                <div style={{ display: 'flex', alignItems: 'center', gap: '24px', color: '#605e5c' }}>
                                                    {/* User */}
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0, overflow: 'hidden' }}>
                                                        <Icon iconName="Contact" styles={{ root: { fontSize: 16 } }} />
                                                        <span style={{ fontSize: '14px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                                            {getAssigneeName(task)}
                                                        </span>
                                                    </div>

                                                    {/* Date */}
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                                        <Icon iconName="Calendar" styles={{ root: { fontSize: 16 } }} />
                                                        <span style={{ fontSize: '14px', whiteSpace: 'nowrap' }}>
                                                            {formatDate(task.TaskDueDate)}
                                                        </span>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    );
                                })}
                            </div>
                        </div>
                    )}
                </div>
                <DialogFooter>
                    <DefaultButton
                        text="✕ No, Cancel"
                        onClick={handleCancelCompletion}
                        disabled={saving}
                        styles={{
                            root: {
                                backgroundColor: '#f3f2f1', // Light grey
                                border: '1px solid #d2d0ce',
                                borderRadius: '4px',
                                height: 36,
                                padding: '0 20px',
                                minWidth: 120
                            },
                            rootHovered: {
                                backgroundColor: '#edebe9',
                                borderColor: '#c8c6c4'
                            },
                            label: {
                                fontWeight: 600,
                                color: '#323130'
                            }
                        }}
                    />
                    <PrimaryButton
                        text="✓ Yes, Complete All"
                        onClick={handleForceCompleteChildren}
                        disabled={saving}
                        styles={{
                            root: {
                                backgroundColor: '#2ea44f', // Github-like green or similar vibrant green
                                borderColor: '#2ea44f',
                                borderRadius: '4px',
                                height: 36,
                                padding: '0 24px',
                                minWidth: 160,
                                boxShadow: '0 1px 0 rgba(27,31,35,0.1)'
                            },
                            rootHovered: {
                                backgroundColor: '#2c974b',
                                borderColor: '#2c974b'
                            },
                            label: {
                                fontWeight: 600,
                                fontSize: '14px'
                            }
                        }}
                    />
                </DialogFooter>
            </Dialog>
        </div>
    );
};
