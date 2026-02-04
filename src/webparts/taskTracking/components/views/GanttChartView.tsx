import * as React from 'react';
import { useState, useMemo, useEffect } from 'react';
import { Icon, Dialog, DialogType, DialogFooter, ActivityItem, PrimaryButton, DefaultButton, TextField, DatePicker, Dropdown, ComboBox, IDropdownOption, IconButton } from 'office-ui-fabric-react';
import styles from './GanttChartView.module.scss';
import { IMainTask, ISubTask } from '../../../../services/interfaces';
import { taskService } from '../../../../services/sp-service';
import * as XLSX from 'xlsx';
import { toPng } from 'html-to-image';
import { sanitizeHtml } from '../../../../utils/sanitize';

export interface IGanttChartViewProps {
    mainTasks: IMainTask[];
    subTasks: ISubTask[];
    onTaskClick?: (task: any, isSubTask: boolean) => void;
}

type ViewGranularity = 'Year' | 'Month' | 'Week' | 'Day';

interface HierarchyItem {
    id: string;
    dbId: number;
    title: string;
    description?: string;
    startDate: Date;
    endDate: Date; // Display/Bar end date
    dueDate: Date; // Planned due date
    actualEndDate?: Date; // Real completion date from SharePoint
    status: string;
    percentComplete: number;
    visualState: 'completed' | 'completedLate' | 'delayed' | 'onTrack';
    level: 1 | 2 | 3;
    parentId?: string;
    isExpanded: boolean;
    type: 'main' | 'sub' | 'sub-sub';
    assignee?: string;
    assigneeId?: number;
}

export const GanttChartView: React.FC<IGanttChartViewProps> = (props) => {
    const [view, setView] = useState<ViewGranularity>('Month');
    const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());
    const [selectedUser, setSelectedUser] = useState<string>('All');
    const [searchTerm, setSearchTerm] = useState<string>('');
    const [showOverdueOnly, setShowOverdueOnly] = useState<boolean>(false);
    const [sortField, setSortField] = useState<keyof HierarchyItem>('startDate');
    const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
    const timelineRef = React.useRef<HTMLDivElement>(null);

    const toggleSort = (field: keyof HierarchyItem) => {
        if (sortField === field) {
            setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
        } else {
            setSortField(field);
            setSortDirection('asc');
        }
    };

    const jumpToToday = () => {
        if (timelineRef.current) {
            const todayPos = getDatePosition(new Date());
            const containerWidth = timelineRef.current.clientWidth;
            timelineRef.current.scrollTo({
                left: todayPos - (containerWidth / 2),
                behavior: 'smooth'
            });
        }
    };
    const [isAdmin, setIsAdmin] = useState<boolean>(false);
    const [allUsers, setAllUsers] = useState<any[]>([]);

    // Correspondence Popup State
    const [isCorrespondenceOpen, setIsCorrespondenceOpen] = useState(false);
    const [correspondenceHistory, setCorrespondenceHistory] = useState<any[]>([]);
    const [loadingCorrespondence, setLoadingCorrespondence] = useState(false);
    const [selectedTask, setSelectedTask] = useState<HierarchyItem | null>(null);

    // Edit Mode State
    const [editMode, setEditMode] = useState<{ field: 'DueDate' | 'Assignee' | 'Description' | 'Status' | null, value: any, reason: string }>({ field: null, value: null, reason: '' });
    const [isSavingEdit, setIsSavingEdit] = useState(false);

    const [currentTime, setCurrentTime] = useState(new Date());

    const handleExpandAll = () => {
        const allIds = new Set<string>();
        // Add all main tasks
        props.mainTasks.forEach(mt => allIds.add(`main-${mt.Id}`));
        // Add all subtasks (only those that can have sub-subtasks)
        props.subTasks.forEach(st => {
            if (!st.ParentSubtaskId || st.ParentSubtaskId === 0) {
                allIds.add(`sub-${st.Id}`);
            }
        });
        setExpandedIds(allIds);
    };

    const handleCollapseAll = () => {
        setExpandedIds(new Set());
    };

    // Helper functions for date formatting
    useEffect(() => {
        // Check Admin
        taskService.verifyAdminUser('').then(setIsAdmin);
        // Fetch Users for Picker
        taskService.getSiteUsers().then(setAllUsers);
    }, []);

    // Extract unique users
    const users = useMemo(() => {
        const uniqueSet = new Set<string>();
        props.subTasks.forEach(st => {
            if ((st.TaskAssignedTo as any)?.Title) {
                uniqueSet.add((st.TaskAssignedTo as any).Title);
            }
        });
        return Array.from(uniqueSet).sort();
    }, [props.subTasks]);

    // 1. Build flattened hierarchy for rendering
    const flattenedData = useMemo(() => {
        const items: HierarchyItem[] = [];
        const now = new Date();
        now.setHours(0, 0, 0, 0);

        const lowerSearch = searchTerm.toLowerCase();

        props.mainTasks.forEach(mt => {
            // Find all subtasks for this main task (including sub-subs)
            const allMtSubTasks = props.subTasks.filter(st => st.Admin_Job_ID === mt.Id);

            // Search filter logic
            const matchesSearch = !searchTerm ||
                mt.Title.toLowerCase().indexOf(lowerSearch) !== -1 ||
                allMtSubTasks.some(st => st.Task_Title.toLowerCase().indexOf(lowerSearch) !== -1);

            if (!matchesSearch) return;

            // User filter logic
            let hasMatchingUser = selectedUser === 'All';
            if (!hasMatchingUser) {
                hasMatchingUser = allMtSubTasks.some(st => (st.TaskAssignedTo as any)?.Title === selectedUser);
            }

            if (!hasMatchingUser) return;

            const mtId = `main-${mt.Id}`;
            const mtDueDate = new Date(mt.TaskDueDate || new Date());
            const mtIsOverdue = mtDueDate < now && mt.Status !== 'Completed';

            let mtVisualState: 'completed' | 'completedLate' | 'delayed' | 'onTrack' = 'onTrack';
            if (mt.Status === 'Completed') {
                if (mt.Task_x0020_End_x0020_Date) {
                    const actualEnd = new Date(mt.Task_x0020_End_x0020_Date);
                    mtVisualState = actualEnd > mtDueDate ? 'completedLate' : 'completed';
                } else {
                    mtVisualState = 'completed';
                }
            } else if (mtIsOverdue) {
                mtVisualState = 'delayed';
            }

            const mtActualEnd = mt.Task_x0020_End_x0020_Date ? new Date(mt.Task_x0020_End_x0020_Date) : undefined;
            const mtItem: HierarchyItem = {
                id: mtId,
                dbId: mt.Id,
                title: mt.Title,
                description: mt.Task_x0020_Description,
                startDate: new Date(mt.TaskStartDate || mt.Created || new Date()),
                endDate: mtActualEnd || mtDueDate,
                dueDate: mtDueDate,
                actualEndDate: mtActualEnd,
                status: mtIsOverdue ? 'Overdue' : mt.Status,
                percentComplete: mt.PercentComplete || 0,
                visualState: mtVisualState,
                level: 1,
                isExpanded: expandedIds.has(mtId),
                type: 'main',
                assignee: (mt.TaskAssignedTo as any)?.Title || 'Unassigned',
                assigneeId: (mt.TaskAssignedTo as any)?.Id
            };
            items.push(mtItem);

            if (expandedIds.has(mtId) || searchTerm) { // Auto-expand if searching? Or just show if expanded
                // Find subtasks belonging to this main task
                const topLevelSubTasks = allMtSubTasks.filter(st => !st.ParentSubtaskId || st.ParentSubtaskId === 0);

                topLevelSubTasks.forEach(st => {
                    const subSubTasks = allMtSubTasks.filter(sst => sst.ParentSubtaskId === st.Id);

                    const subMatchesSearch = !searchTerm ||
                        st.Task_Title.toLowerCase().indexOf(lowerSearch) !== -1 ||
                        subSubTasks.some(sst => sst.Task_Title.toLowerCase().indexOf(lowerSearch) !== -1);

                    if (!subMatchesSearch) return;

                    if (selectedUser !== 'All' && (st.TaskAssignedTo as any)?.Title !== selectedUser && !subSubTasks.some(sst => (sst.TaskAssignedTo as any)?.Title === selectedUser)) {
                        return;
                    }

                    const stId = `sub-${st.Id}`;
                    const stDueDate = new Date(st.TaskDueDate || mtItem.endDate);
                    const stIsOverdue = stDueDate < now && st.TaskStatus !== 'Completed';

                    let stVisualState: 'completed' | 'completedLate' | 'delayed' | 'onTrack' = 'onTrack';
                    if (st.TaskStatus === 'Completed') {
                        if (st.Task_End_Date) {
                            const actualEnd = new Date(st.Task_End_Date);
                            stVisualState = actualEnd > stDueDate ? 'completedLate' : 'completed';
                        } else {
                            stVisualState = 'completed';
                        }
                    } else if (stIsOverdue) {
                        stVisualState = 'delayed';
                    }

                    const stActualEnd = st.Task_End_Date ? new Date(st.Task_End_Date) : undefined;
                    const stItem: HierarchyItem = {
                        id: stId,
                        dbId: st.Id,
                        title: st.Task_Title,
                        description: st.Task_Description,
                        startDate: new Date((st as any).TaskStartDate || (st as any).Created || mtItem.startDate),
                        endDate: stActualEnd || stDueDate,
                        dueDate: stDueDate,
                        actualEndDate: stActualEnd,
                        status: stIsOverdue ? 'Overdue' : st.TaskStatus,
                        percentComplete: st.TaskStatus === 'Completed' ? 100 : 50, // Approximation
                        visualState: stVisualState,
                        level: 2,
                        parentId: mtId,
                        isExpanded: expandedIds.has(stId),
                        type: 'sub',
                        assignee: (st.TaskAssignedTo as any)?.Title || 'Unassigned',
                        assigneeId: (st.TaskAssignedTo as any)?.Id
                    };
                    items.push(stItem);

                    if (expandedIds.has(stId) || searchTerm) {
                        subSubTasks.forEach(sst => {
                            const sstMatchesSearch = !searchTerm || sst.Task_Title.toLowerCase().indexOf(lowerSearch) !== -1;
                            if (!sstMatchesSearch) return;

                            if (selectedUser !== 'All' && (sst.TaskAssignedTo as any)?.Title !== selectedUser) return;

                            const sstId = `subsub-${sst.Id}`;
                            const sstDueDate = new Date(sst.TaskDueDate || stItem.endDate);
                            const sstIsOverdue = sstDueDate < now && sst.TaskStatus !== 'Completed';

                            let sstVisualState: 'completed' | 'completedLate' | 'delayed' | 'onTrack' = 'onTrack';
                            if (sst.TaskStatus === 'Completed') {
                                if (sst.Task_End_Date) {
                                    const actualEnd = new Date(sst.Task_End_Date);
                                    sstVisualState = actualEnd > sstDueDate ? 'completedLate' : 'completed';
                                } else {
                                    sstVisualState = 'completed';
                                }
                            } else if (sstIsOverdue) {
                                sstVisualState = 'delayed';
                            }

                            const sstActualEnd = sst.Task_End_Date ? new Date(sst.Task_End_Date) : undefined;
                            items.push({
                                id: sstId,
                                dbId: sst.Id,
                                title: sst.Task_Title,
                                startDate: new Date((sst as any).TaskStartDate || (sst as any).Created || stItem.startDate),
                                endDate: sstActualEnd || sstDueDate,
                                dueDate: sstDueDate,
                                actualEndDate: sstActualEnd,
                                status: sstIsOverdue ? 'Overdue' : sst.TaskStatus,
                                percentComplete: sst.TaskStatus === 'Completed' ? 100 : 50, // Approximation
                                visualState: sstVisualState,
                                level: 3,
                                parentId: stId,
                                isExpanded: false,
                                type: 'sub-sub',
                                assignee: (sst.TaskAssignedTo as any)?.Title || 'Unassigned',
                                assigneeId: (sst.TaskAssignedTo as any)?.Id
                            });
                        });
                    }
                });
            }
        });

        // Apply Sorting
        items.sort((a, b) => {
            const valA = a[sortField];
            const valB = b[sortField];

            if (valA === valB) return 0;
            if (valA === undefined || valA === null) return 1;
            if (valB === undefined || valB === null) return -1;

            const multiplier = sortDirection === 'asc' ? 1 : -1;
            return valA < valB ? -1 * multiplier : 1 * multiplier;
        });

        return items;
    }, [props.mainTasks, props.subTasks, expandedIds, selectedUser, showOverdueOnly, searchTerm, sortField, sortDirection]);

    const stats = useMemo(() => {
        if (!flattenedData.length) return { onTrack: 0, overdue: 0, completed: 0, completedLate: 0 };
        const total = flattenedData.length;
        const counts = {
            onTrack: flattenedData.filter(i => i.visualState === 'onTrack').length,
            overdue: flattenedData.filter(i => i.visualState === 'delayed').length,
            completed: flattenedData.filter(i => i.visualState === 'completed').length,
            completedLate: flattenedData.filter(i => i.visualState === 'completedLate').length,
        };
        return {
            onTrack: Math.round((counts.onTrack / total) * 100),
            overdue: Math.round((counts.overdue / total) * 100),
            completed: Math.round((counts.completed / total) * 100),
            completedLate: Math.round((counts.completedLate / total) * 100),
        };
    }, [flattenedData]);

    // 2. Timeline Logic
    const timelineDates = useMemo(() => {
        const start = new Date();
        const end = new Date();

        if (view === 'Day') {
            start.setDate(start.getDate() - 7);
            end.setDate(end.getDate() + 60);
        } else if (view === 'Week') {
            start.setMonth(start.getMonth() - 1);
            end.setMonth(end.getMonth() + 6);
        } else if (view === 'Month') {
            start.setMonth(start.getMonth() - 3);
            end.setMonth(end.getMonth() + 12);
        } else { // Year
            start.setFullYear(start.getFullYear() - 1);
            end.setFullYear(end.getFullYear() + 5);
        }

        // Normalize to start of unit for cleaner alignment
        if (view === 'Week') {
            const day = start.getDay();
            start.setDate(start.getDate() - day);
        } else if (view === 'Month') {
            start.setDate(1);
        } else if (view === 'Year') {
            start.setMonth(0, 1);
        }
        start.setHours(0, 0, 0, 0);

        const dates: Date[] = [];
        let curr = new Date(start);

        while (curr.getTime() <= end.getTime()) {
            dates.push(new Date(curr));
            const nextDate = new Date(curr);
            if (view === 'Day') nextDate.setDate(nextDate.getDate() + 1);
            else if (view === 'Week') nextDate.setDate(nextDate.getDate() + 7);
            else if (view === 'Month') nextDate.setMonth(nextDate.getMonth() + 1);
            else nextDate.setFullYear(nextDate.getFullYear() + 1);

            if (nextDate.getTime() === curr.getTime()) break;
            curr = nextDate; // Explicitly update loop variable
        }
        return dates;
    }, [view]);

    // Width calculation for 1 day in pixels depending on view
    const unitWidth = useMemo(() => {
        switch (view) {
            case 'Day': return 80;
            case 'Week': return 120;
            case 'Month': return 200;
            case 'Year': return 300;
            default: return 200;
        }
    }, [view]);

    const toggleExpand = (id: string, e: React.MouseEvent) => {
        e.stopPropagation();
        const newSet = new Set(expandedIds);
        if (newSet.has(id)) newSet.delete(id);
        else newSet.add(id);
        setExpandedIds(newSet);
    };

    const getDatePosition = (date: Date): number => {
        if (!timelineDates.length) return 0;
        const targetTime = date.getTime();
        const startTime = timelineDates[0].getTime();

        if (targetTime <= startTime) return 0;

        // Find which column this date falls into
        let columnIndex = -1;
        for (let i = 0; i < timelineDates.length; i++) {
            if (targetTime >= timelineDates[i].getTime()) {
                columnIndex = i;
            } else {
                break;
            }
        }

        if (columnIndex === -1) return 0;
        if (columnIndex >= timelineDates.length - 1) return columnIndex * unitWidth;

        // Calculate progress within that column
        const columnStart = timelineDates[columnIndex].getTime();
        const columnEnd = timelineDates[columnIndex + 1].getTime();
        const columnDuration = columnEnd - columnStart;
        const progressInColumn = (targetTime - columnStart) / columnDuration;

        return (columnIndex + progressInColumn) * unitWidth;
    };

    const getBarStyles = (item: HierarchyItem) => {
        const left = getDatePosition(item.startDate);
        const endPos = getDatePosition(item.endDate);
        const width = Math.max(4, endPos - left);

        return {
            left: `${left}px`,
            width: `${width}px`,
            minWidth: '4px'
        };
    };

    const handleOpenCorrespondence = async (item: HierarchyItem, e: React.MouseEvent) => {
        e.stopPropagation();
        setSelectedTask(item);
        setIsCorrespondenceOpen(true);
        setLoadingCorrespondence(true);
        setEditMode({ field: null, value: null, reason: '' });

        try {
            let parentId = 0;
            let childId = 0;

            if (item.type === 'main') {
                parentId = item.dbId;
            } else if (item.type === 'sub') {
                // Should extract parent from ID structure or find in props, but item.parentId is 'main-x'
                if (item.parentId && item.parentId.startsWith('main-')) {
                    parentId = parseInt(item.parentId.split('-')[1]);
                    childId = item.dbId;
                }
            } else if (item.type === 'sub-sub') {
                // For sub-sub, structure is trickier, generally we might not have 'Task Correspondence' linked deeply to level 3 in standardized way yet
                // But typically it links to parent task ID. Let's assume sub-sub links to its subtask as parent?
                // Actually sp-service logic is `ParentTaskID` and `ChildTaskID`.
                // Existing logic usually supports 2 levels.
                // Let's try to map to main and sub.
                // Use recursion or flattened lookup logic.
                // Simplified: Treat sub-sub as just a child logic? Or fetch correspondence where ReferenceID matches?
                // Given sp-service `getCorrespondenceByTaskId(parentId, childId)`, likely sub-sub isn't fully supported or treats sub as parent.
                // Let's just try to find its parent.
                // Actually in flattenedData, we tracked parentId.
                const parentSub = props.subTasks.find(s => `sub-${s.Id}` === item.parentId);
                if (parentSub) {
                    parentId = parentSub.Admin_Job_ID;
                    childId = item.dbId;
                }
            }

            if (parentId) {
                const history = await taskService.getCorrespondenceByTaskId(parentId, childId || undefined);
                setCorrespondenceHistory(history);
            } else {
                setCorrespondenceHistory([]);
            }
        } catch (error) {
            console.error(error);
            setCorrespondenceHistory([]);
        } finally {
            setLoadingCorrespondence(false);
        }
    };

    const formatDate = (date: Date | string): string => {
        if (!date) return '';
        const d = new Date(date);
        const day = ('0' + d.getDate()).slice(-2);
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const month = months[d.getMonth()];
        const year = d.getFullYear();
        return `${day}-${month}-${year}`;
    };

    const handleSaveEdit = async () => {
        if (!selectedTask || !editMode.field || editMode.value === null || !editMode.reason.trim()) {
            return;
        }

        setIsSavingEdit(true);
        try {
            // Construct "Old -> New" message
            let oldValue = '';
            let newValue = '';

            if (editMode.field === 'DueDate') {
                oldValue = formatDate(selectedTask.endDate);
                newValue = formatDate(editMode.value);
            } else if (editMode.field === 'Assignee') {
                oldValue = selectedTask.assignee || 'Unassigned';
                const newUser = allUsers.find(u => u.Id === editMode.value);
                newValue = newUser ? newUser.Title : 'Unknown';
            } else if (editMode.field === 'Description') {
                oldValue = selectedTask.description || 'No description';
                newValue = editMode.value.substring(0, 50) + (editMode.value.length > 50 ? '...' : ''); // Truncate for log
            } else if (editMode.field === 'Status') {
                oldValue = selectedTask.status;
                newValue = editMode.value;
            }

            // Enhanced Remark: "Field changed from Old to New. Reason: User Remark"
            // Note: service appends "Reason: " so we format it to flow naturally or override the service logic?
            // Service does: `logMessage = ... Reason: ${options.remark}`
            // So we pass: "Old to New. Reason" is handled by service?
            // Actually, service logic in `sp-service.ts` (lines 677 and 734) hardcodes: `... changed to ... Reason: ${remark}`.
            // It doesn't know the OLD value.
            // If we want "Old -> New", we should probably PREPEND it to the remark if we can't change the service, OR we expect the service output to be sufficient?
            // User asked: "old data to new data". Service only logs "changed to New".
            // Strategy: We will modify the remark passed to the service to include the old value, so the final log looks like:
            // "Due Date changed to [New]. Reason: (Previous: [Old]) [User Remark]" -> somewhat clumsy.
            // Better: Modify `sp-service.ts` to accept a custom formatted message? No, that's defined in interface.
            // Best approach without changing service interface deeply:
            // Pass a Combined Reason: "[Old -> New] My actual remark"

            const detailedRemark = `(Changed from ${oldValue}) ${editMode.reason}`;

            if (selectedTask.type === 'main') {
                // Main Task Update
                await taskService.updateMainTaskField({
                    mainTaskId: selectedTask.dbId,
                    title: selectedTask.title,
                    field: editMode.field as 'DueDate' | 'Description' | 'Status',
                    newValue: editMode.field === 'DueDate' ? editMode.value.toISOString() : editMode.value,
                    remark: detailedRemark
                });
            } else {
                // Sub Task Update
                let mainTaskId = 0;
                if (selectedTask.parentId && selectedTask.parentId.startsWith('main-')) {
                    mainTaskId = parseInt(selectedTask.parentId.split('-')[1]);
                } else {
                    const lookup = props.subTasks.find(s => s.Id === selectedTask.dbId);
                    if (lookup) mainTaskId = lookup.Admin_Job_ID;
                }

                await taskService.updateSubTaskField({
                    subTaskId: selectedTask.dbId,
                    mainTaskId: mainTaskId,
                    title: selectedTask.title,
                    field: editMode.field as any, // Cast to match extended type in service
                    newValue: editMode.field === 'DueDate' ? editMode.value.toISOString() : editMode.value,
                    remark: detailedRemark
                });
            }

            setEditMode({ field: null, value: null, reason: '' });

            // Update selectedTask with new values to reflect in UI
            if (selectedTask) {
                const updatedTask = { ...selectedTask };

                if (editMode.field === 'Assignee') {
                    const newUser = allUsers.find(u => u.Id === editMode.value);
                    updatedTask.assignee = newUser ? newUser.Title : updatedTask.assignee;
                    updatedTask.assigneeId = editMode.value as number;
                } else if (editMode.field === 'DueDate') {
                    updatedTask.endDate = new Date(editMode.value);
                } else if (editMode.field === 'Status') {
                    updatedTask.status = editMode.value as string;
                } else if (editMode.field === 'Description') {
                    updatedTask.description = editMode.value as string;
                }

                setSelectedTask(updatedTask);
            }

            const history = await taskService.getCorrespondenceByTaskId(
                selectedTask.type === 'main' ? selectedTask.dbId : (selectedTask.parentId ? parseInt(selectedTask.parentId.split('-')[1]) : 0),
                selectedTask.type !== 'main' ? selectedTask.dbId : undefined
            );
            setCorrespondenceHistory(history);

        } catch (error) {
            console.error(error);
            alert('Error updating task: ' + error.message);
        } finally {
            setIsSavingEdit(false);
        }
    };

    const renderTaskDetailRow = (label: string, value: string | JSX.Element, field: 'DueDate' | 'Assignee' | 'Description' | 'Status' | null, canEdit: boolean, rawValue?: any) => {
        const isEditing = field !== null && editMode.field === field;

        return (
            <div style={{ display: 'flex', alignItems: 'center', marginBottom: 8, minHeight: 32 }}>
                <div style={{ width: 120, fontWeight: 600, color: '#666' }}>{label}:</div>
                <div style={{ flex: 1, display: 'flex', alignItems: 'center' }}>
                    {isEditing ? (
                        <div style={{ display: 'flex', alignItems: 'center', flex: 1, gap: 10 }}>
                            {field === 'DueDate' && (
                                <DatePicker
                                    value={editMode.value ? new Date(editMode.value) : new Date()}
                                    onSelectDate={(d) => setEditMode({ ...editMode, value: d })}
                                    style={{ width: 140 }}
                                    formatDate={(d) => formatDate(d || new Date())}
                                />
                            )}
                            {field === 'Assignee' && (
                                <ComboBox
                                    options={allUsers.map(u => ({ key: u.Id, text: u.Title }))}
                                    selectedKey={editMode.value}
                                    onChange={(e, opt) => setEditMode({ ...editMode, value: opt?.key })}
                                    allowFreeform={true}
                                    autoComplete="on"
                                    useComboBoxAsMenuWidth={true}
                                    styles={{ root: { width: 200 } }}
                                />
                            )}
                            {field === 'Status' && (
                                <Dropdown
                                    options={[
                                        { key: 'Not Started', text: 'Not Started' },
                                        { key: 'In Progress', text: 'In Progress' },
                                        { key: 'Completed', text: 'Completed' },
                                        { key: 'Deferred', text: 'Deferred' },
                                        { key: 'Waiting on someone else', text: 'Waiting on someone else' }
                                    ]}
                                    selectedKey={editMode.value}
                                    onChange={(e, opt) => setEditMode({ ...editMode, value: opt?.key })}
                                    styles={{ root: { width: 150 } }}
                                />
                            )}
                            {field === 'Description' && (
                                <TextField
                                    multiline
                                    rows={3}
                                    value={editMode.value}
                                    onChange={(e, v) => setEditMode({ ...editMode, value: v || '' })}
                                    styles={{ root: { width: 300 } }}
                                />
                            )}
                            <TextField
                                placeholder="Reason for change (Required)"
                                value={editMode.reason}
                                onChange={(e, v) => setEditMode({ ...editMode, reason: v || '' })}
                                styles={{ root: { flex: 1 } }}
                            />
                            <IconButton iconProps={{ iconName: "CheckMark" }} onClick={handleSaveEdit} disabled={!editMode.reason || isSavingEdit} title="Save" />
                            <IconButton iconProps={{ iconName: "Cancel" }} onClick={() => setEditMode({ field: null, value: null, reason: '' })} title="Cancel" />
                        </div>
                    ) : (
                        <>
                            <div style={{ fontWeight: 500 }}>{field === 'DueDate' && typeof rawValue === 'object' ? formatDate(rawValue) : value}</div>
                            {canEdit && (
                                <Icon
                                    iconName="Edit"
                                    style={{ marginLeft: 10, cursor: 'pointer', color: '#0078d4' }}
                                    onClick={() => setEditMode({ field, value: rawValue, reason: '' })}
                                />
                            )}
                        </>
                    )}
                </div>
            </div>
        );
    };

    const handleExportExcel = () => {
        // 1. Prepare Full Data (regardless of UI expansion)
        const allItems: HierarchyItem[] = [];
        const now = new Date();
        now.setHours(0, 0, 0, 0);

        props.mainTasks.forEach(mt => {
            const mtId = `main-${mt.Id}`;
            const mtDueDate = mt.TaskDueDate ? new Date(mt.TaskDueDate) : new Date((mt as any).Created || new Date());
            const mtActualEnd = mt.Task_x0020_End_x0020_Date ? new Date(mt.Task_x0020_End_x0020_Date) : undefined;
            const mtIsOverdue = mtDueDate < now && mt.Status !== 'Completed';

            const mtItem: HierarchyItem = {
                id: mtId,
                dbId: mt.Id,
                title: mt.Title,
                startDate: new Date((mt as any).TaskStartDate || (mt as any).Created || new Date()),
                endDate: mtActualEnd || mtDueDate,
                dueDate: mtDueDate,
                actualEndDate: mtActualEnd,
                status: mtIsOverdue ? 'Overdue' : mt.Status,
                percentComplete: mt.Status === 'Completed' ? 100 : 0,
                visualState: mtIsOverdue ? 'delayed' : (mt.Status === 'Completed' ? 'completed' : 'onTrack'),
                level: 1,
                isExpanded: true,
                type: 'main',
                assignee: (mt.TaskAssignedTo as any)?.Title || 'Unassigned',
                assigneeId: (mt.TaskAssignedTo as any)?.Id
            };
            allItems.push(mtItem);

            const allMtSubTasks = props.subTasks.filter(st => st.Admin_Job_ID === mt.Id);
            const topLevelSubTasks = allMtSubTasks.filter(st => !st.ParentSubtaskId || st.ParentSubtaskId === 0);

            topLevelSubTasks.forEach(st => {
                const subSubTasks = allMtSubTasks.filter(sst => sst.ParentSubtaskId === st.Id);
                const stId = `sub-${st.Id}`;
                const stDueDate = new Date(st.TaskDueDate || mtItem.dueDate);
                const stActualEnd = st.Task_End_Date ? new Date(st.Task_End_Date) : undefined;
                const stIsOverdue = stDueDate < now && st.TaskStatus !== 'Completed';

                const stItem: HierarchyItem = {
                    id: stId,
                    dbId: st.Id,
                    title: st.Task_Title,
                    description: st.Task_Description,
                    startDate: new Date((st as any).TaskStartDate || (st as any).Created || mtItem.startDate),
                    endDate: stActualEnd || stDueDate,
                    dueDate: stDueDate,
                    actualEndDate: stActualEnd,
                    status: stIsOverdue ? 'Overdue' : st.TaskStatus,
                    percentComplete: st.TaskStatus === 'Completed' ? 100 : 50,
                    visualState: stIsOverdue ? 'delayed' : (st.TaskStatus === 'Completed' ? 'completed' : 'onTrack'),
                    level: 2,
                    parentId: mtId,
                    isExpanded: true,
                    type: 'sub',
                    assignee: (st.TaskAssignedTo as any)?.Title || 'Unassigned',
                    assigneeId: (st.TaskAssignedTo as any)?.Id
                };
                allItems.push(stItem);

                subSubTasks.forEach(sst => {
                    const sstId = `subsub-${sst.Id}`;
                    const sstDueDate = new Date(sst.TaskDueDate || stItem.dueDate);
                    const sstActualEnd = sst.Task_End_Date ? new Date(sst.Task_End_Date) : undefined;
                    const sstIsOverdue = sstDueDate < now && sst.TaskStatus !== 'Completed';

                    allItems.push({
                        id: sstId,
                        dbId: sst.Id,
                        title: sst.Task_Title,
                        startDate: new Date((sst as any).TaskStartDate || (sst as any).Created || stItem.startDate),
                        endDate: sstActualEnd || sstDueDate,
                        dueDate: sstDueDate,
                        actualEndDate: sstActualEnd,
                        status: sstIsOverdue ? 'Overdue' : sst.TaskStatus,
                        percentComplete: sst.TaskStatus === 'Completed' ? 100 : 50,
                        visualState: sstIsOverdue ? 'delayed' : (sst.TaskStatus === 'Completed' ? 'completed' : 'onTrack'),
                        level: 3,
                        parentId: stId,
                        isExpanded: false,
                        type: 'sub-sub',
                        assignee: (sst.TaskAssignedTo as any)?.Title || 'Unassigned',
                        assigneeId: (sst.TaskAssignedTo as any)?.Id
                    });
                });
            });
        });

        // 2. Prepare Data Headers
        const baseHeaders = ['Task Name', 'Assignee', 'Status', 'Start Date', 'Due Date', 'End Date', 'Days Overdue'];
        const timelineHeaders = timelineDates.map(d => {
            if (view === 'Day') return formatDate(d);
            if (view === 'Week') return `Wk ${d.getDate()} / ${d.getMonth() + 1}`;
            if (view === 'Month') return `${d.toLocaleString('default', { month: 'short' })} ${d.getFullYear()}`;
            return d.getFullYear().toString();
        });

        const allHeaders = [...baseHeaders, '', ...timelineHeaders];

        // 3. Prepare Rows
        const rows = allItems.map(item => {
            const now = new Date();
            now.setHours(0, 0, 0, 0);
            const endDate = new Date(item.endDate);
            endDate.setHours(0, 0, 0, 0);
            const daysOverdue = item.status !== 'Completed' && endDate < now
                ? Math.floor((now.getTime() - endDate.getTime()) / (1000 * 60 * 60 * 24))
                : 0;

            // Indent name by level
            const indentedName = '  '.repeat(item.level - 1) + item.title;

            const baseRow = [
                indentedName,
                item.assignee || 'Unassigned',
                item.status,
                formatDate(item.startDate),
                formatDate(item.dueDate),
                item.actualEndDate ? formatDate(item.actualEndDate) : '-',
                daysOverdue > 0 ? `${daysOverdue} days` : '-'
            ];

            // 4. Create Timeline Visualization in Cells
            const timelineRow = timelineDates.map((d, idx) => {
                const nextD = timelineDates[idx + 1] || new Date(d.getTime() + (d.getTime() - timelineDates[idx - 1]?.getTime() || 86400000));

                // Check if task overlaps this interval
                const taskStart = item.startDate.getTime();
                const taskEnd = item.endDate.getTime();
                const intervalStart = d.getTime();
                const intervalEnd = nextD.getTime();

                const isOngoing = taskStart < intervalEnd && taskEnd >= intervalStart;

                if (isOngoing) {
                    return '█'; // Visual block character
                }
                return '';
            });

            return [...baseRow, '', ...timelineRow];
        });

        // 5. Create Workbook
        const worksheet = XLSX.utils.aoa_to_sheet([allHeaders, ...rows]);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Gantt Export");

        // 6. Download
        XLSX.writeFile(workbook, `Gantt_Export_${formatDate(new Date())}.xlsx`);
    };

    const handleExportImage = async () => {
        const container = document.querySelector(`.${styles.ganttContainer}`) as HTMLElement;
        if (!container) return;

        try {
            const dataUrl = await toPng(container, {
                backgroundColor: '#111', // Match theme background
                quality: 1,
                pixelRatio: 2 // High res
            });

            const link = document.createElement('a');
            link.download = `Gantt_Timeline_${formatDate(new Date())}.png`;
            link.href = dataUrl;
            link.click();
        } catch (error) {
            console.error('Error exporting image:', error);
            alert('Failed to export image. Try in a different browser.');
        }
    };

    return (
        <div className={styles.ganttContainer}>
            <header className={styles.header}>
                <div className={styles.titleSection}>
                    <h2>Task Timeline</h2>
                    <p>Strategic Workflow Orchestration</p>
                </div>

                <div className={styles.controls}>
                    <div className={styles.legend}>
                        <span style={{ fontSize: 10, color: '#aaa', marginRight: 10, textTransform: 'uppercase', fontWeight: 600 }}>Stats ({flattenedData.length}):</span>
                        <div className={styles.legendItem}>
                            <span className={`${styles.dot} ${styles.onTrack}`} />
                            <span>On Track ({stats.onTrack}%)</span>
                        </div>
                        <div className={styles.legendItem}>
                            <span className={`${styles.dot} ${styles.overdue}`} />
                            <span>Overdue ({stats.overdue}%)</span>
                        </div>
                        <div className={styles.legendItem}>
                            <span className={`${styles.dot} ${styles.completed}`} />
                            <span>On Time ({stats.completed}%)</span>
                        </div>
                        <div className={styles.legendItem}>
                            <span className={`${styles.dot} ${styles.completedLate}`} />
                            <span>Completed Late ({stats.completedLate}%)</span>
                        </div>
                    </div>

                    <div className={styles.viewSelectorContainer}>
                        <div className={styles.searchContainer} style={{ marginRight: 15 }}>
                            <TextField
                                placeholder="Search by task name..."
                                value={searchTerm}
                                onChange={(_, v) => setSearchTerm(v || '')}
                                iconProps={{ iconName: 'Search' }}
                                styles={{
                                    root: { width: 220 },
                                    fieldGroup: { background: 'rgba(255, 255, 255, 0.05)', border: '1px solid rgba(255, 255, 255, 0.1)', borderRadius: 8 },
                                    field: { color: 'white' }
                                }}
                            />
                        </div>
                        <Icon iconName="Contact" style={{ marginRight: 8, color: '#00f2fe', fontSize: 16 }} />
                        <select
                            className={styles.viewSelector}
                            value={selectedUser}
                            onChange={(e) => setSelectedUser(e.target.value)}
                            style={{ marginRight: 10, width: 140 }}
                        >
                            <option value="All">All Users</option>
                            {users.map(u => <option key={u} value={u}>{u}</option>)}
                        </select>

                        <div
                            className={styles.viewSelector}
                            style={{ display: 'flex', alignItems: 'center', marginRight: 10, cursor: 'pointer', background: showOverdueOnly ? 'rgba(255, 0, 85, 0.2)' : undefined, borderColor: showOverdueOnly ? '#ff0055' : undefined }}
                            onClick={() => setShowOverdueOnly(!showOverdueOnly)}
                        >
                            <Icon iconName="Warning" style={{ marginRight: 5, color: showOverdueOnly ? '#ff0055' : 'white' }} />
                            <span style={{ fontSize: 12 }}>Overdue</span>
                        </div>

                        <Icon iconName="Calendar" style={{ marginRight: 8, color: '#00f2fe', fontSize: 16 }} />
                        <select
                            className={styles.viewSelector}
                            value={view}
                            onChange={(e) => setView(e.target.value as ViewGranularity)}
                        >
                            <option value="Day">Daily View</option>
                            <option value="Week">Weekly View</option>
                            <option value="Month">Monthly View</option>
                            <option value="Year">Yearly View</option>
                        </select>
                    </div>
                    {/* <button className={styles.actionButton}><Icon iconName="Filter" style={{ fontSize: 16 }} /></button> */}
                    <button
                        className={styles.actionButton}
                        onClick={handleExpandAll}
                        title="Expand All"
                    >
                        <Icon iconName="DoubleChevronDown" style={{ fontSize: 16 }} />
                    </button>
                    <button
                        className={styles.actionButton}
                        onClick={handleCollapseAll}
                        title="Collapse All"
                    >
                        <Icon iconName="DoubleChevronUp" style={{ fontSize: 16 }} />
                    </button>
                    <button
                        className={styles.actionButton}
                        onClick={() => jumpToToday()}
                        title="Jump to Today"
                    >
                        <Icon iconName="GotoToday" style={{ fontSize: 16 }} />
                    </button>
                    <button className={styles.actionButton} title="Export Timeline to Excel" onClick={handleExportExcel}>
                        <Icon iconName="ExcelDocument" style={{ fontSize: 16 }} />
                    </button>
                    <button className={styles.actionButton} title="Export Timeline to Image" onClick={handleExportImage}>
                        <Icon iconName="Photo2" style={{ fontSize: 16 }} />
                    </button>
                    <button className={styles.actionButton}><Icon iconName="FullScreen" style={{ fontSize: 16 }} /></button>
                </div>
            </header>

            <div className={styles.mainContent}>
                {/* Sidebar */}
                <div className={styles.sidebar}>
                    <div className={styles.sidebarHeader}>
                        <div onClick={() => toggleSort('title')} style={{ cursor: 'pointer' }}>
                            Task Name {sortField === 'title' && <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} />}
                        </div>
                        <div onClick={() => toggleSort('assignee')} style={{ cursor: 'pointer' }}>
                            Assignee {sortField === 'assignee' && <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} />}
                        </div>
                        <div onClick={() => toggleSort('status')} style={{ cursor: 'pointer' }}>
                            Status {sortField === 'status' && <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} />}
                        </div>
                        <div onClick={() => toggleSort('startDate')} style={{ cursor: 'pointer' }}>
                            Start Date {sortField === 'startDate' && <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} />}
                        </div>
                        <div onClick={() => toggleSort('dueDate')} style={{ cursor: 'pointer' }}>
                            Due Date {sortField === 'dueDate' && <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} />}
                        </div>
                        <div onClick={() => toggleSort('endDate')} style={{ cursor: 'pointer' }}>
                            End Date {sortField === 'endDate' && <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} />}
                        </div>
                        <div>Days Overdue</div>
                        <div>View</div>
                    </div>
                    <div className={styles.taskList}>
                        {flattenedData.map(item => {
                            const now = new Date();
                            now.setHours(0, 0, 0, 0);
                            const endDate = new Date(item.endDate);
                            endDate.setHours(0, 0, 0, 0);
                            const daysOverdue = item.status !== 'Completed' && endDate < now
                                ? Math.floor((now.getTime() - endDate.getTime()) / (1000 * 60 * 60 * 24))
                                : 0;

                            return (
                                <div
                                    key={item.id}
                                    className={`${styles.taskRow} ${item.level > 1 ? styles[`level${item.level}`] : ''} `}
                                >
                                    <div className={styles.mainColumn} style={{ paddingLeft: `${item.level * 16} px` }}>
                                        <div className={styles.toggleIcon} onClick={(e) => toggleExpand(item.id, e)}>
                                            {item.type !== 'sub-sub' && (
                                                item.isExpanded ? <Icon iconName="ChevronDown" style={{ fontSize: 12 }} /> : <Icon iconName="ChevronRight" style={{ fontSize: 12 }} />
                                            )}
                                        </div>
                                        <div className={styles.taskInfo}>
                                            <div className={styles.taskName} title={item.title}>{item.title}</div>
                                        </div>
                                    </div>

                                    <div className={styles.metaColumn} title={item.assignee || 'Unassigned'}>
                                        {item.assignee || '-'}
                                    </div>
                                    <div className={`${styles.metaColumn} ${styles.status} ${item.status === 'Completed' ? styles.completed :
                                        item.status === 'Overdue' ? styles.overdue :
                                            item.status === 'In Progress' ? styles.inprogress : ''
                                        } `}>
                                        {item.status || 'Not Started'}
                                    </div>
                                    <div className={styles.metaColumn}>
                                        {item.startDate ? formatDate(item.startDate) : '-'}
                                    </div>
                                    <div className={styles.metaColumn}>
                                        {item.dueDate ? formatDate(item.dueDate) : '-'}
                                    </div>
                                    <div className={styles.metaColumn} style={{ color: item.status === 'Overdue' ? '#ff0055' : undefined }}>
                                        {item.actualEndDate ? formatDate(item.actualEndDate) : '-'}
                                    </div>
                                    <div className={styles.metaColumn} style={{
                                        color: daysOverdue > 0 ? '#ff0055' : '#666',
                                        fontWeight: daysOverdue > 0 ? 600 : 400
                                    }}>
                                        {daysOverdue > 0 ? `${daysOverdue} days` : '-'}
                                    </div>
                                    <div className={styles.metaColumn}>
                                        <Icon
                                            iconName="RedEye"
                                            className={styles.actionIcon}
                                            title="View Correspondence"
                                            onClick={(e) => handleOpenCorrespondence(item, e)}
                                        />
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </div>

                {/* Timeline Grid */}
                <div className={styles.timelineContainer} ref={timelineRef}>
                    <div className={styles.timelineHeader}>
                        <div className={styles.dateGrid} style={{ display: 'flex', width: `${timelineDates.length * unitWidth}px` }}>
                            {timelineDates.map((date, idx) => (
                                <div
                                    key={idx}
                                    className={styles.dateCell}
                                    style={{
                                        width: unitWidth,
                                        background: (view === 'Day' && (date.getDay() === 0 || date.getDay() === 6)) ? 'rgba(255, 255, 255, 0.03)' : undefined
                                    }}
                                >
                                    {view === 'Day' ? date.toLocaleDateString('en-US', { day: 'numeric', month: 'short' }) :
                                        view === 'Week' ? `Week of ${date.toLocaleDateString('en-US', { day: 'numeric', month: 'short' })}` :
                                            view === 'Month' ? date.toLocaleString('default', { month: 'long', year: 'numeric' }) :
                                                date.getFullYear()}
                                </div>
                            ))}
                        </div>
                    </div>

                    <div className={styles.timelineBody} style={{ width: `${timelineDates.length * unitWidth}px` }}>
                        {/* Today Line */}
                        <div
                            className={styles.todayIndicator}
                            style={{ left: `${getDatePosition(new Date())}px` }}
                        />

                        {flattenedData.map(item => (
                            <div key={item.id} className={styles.timelineRow}>
                                <div
                                    className={`${styles.taskBar} ${styles[item.type] || ''} ${styles[item.visualState] || ''} `}
                                    title={`${item.title}\nStatus: ${item.status}\nProgress: ${item.percentComplete}%\nDue: ${formatDate(item.endDate)}\nDescription: ${item.description || 'No description'}`}
                                    style={{
                                        ...getBarStyles(item),
                                        transition: 'width 0.5s ease-out, opacity 0.5s ease-out',
                                        opacity: 1
                                    }}
                                >
                                    <div
                                        className={styles.progressBar}
                                        style={{ width: `${item.percentComplete}%` }}
                                    />
                                    <span className={styles.barLabel}>{item.title} ({item.percentComplete}%)</span>
                                </div>
                            </div>
                        ))}

                        {/* Background Grid Lines */}
                        {timelineDates.map((date, idx) => (
                            <div
                                key={`grid - ${idx} `}
                                className={styles.gridLine}
                                style={{ left: `${idx * unitWidth}px` }}
                            />
                        ))}
                    </div>
                </div>
            </div>

            <Dialog
                hidden={!isCorrespondenceOpen}
                onDismiss={() => setIsCorrespondenceOpen(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: `Task Details: ${selectedTask?.title || 'Unknown'} `,
                    subText: 'View details and correspondence history.'
                }}
                minWidth={600}
                maxWidth={700}
            >
                <div style={{ maxHeight: '500px', overflowY: 'auto', paddingRight: 10 }}>
                    {selectedTask && (
                        <div style={{ marginBottom: 20, padding: 15, background: '#f8f8f8', borderRadius: 4, borderLeft: '4px solid #0078d4' }}>
                            <h3 style={{ marginTop: 0, marginBottom: 15, borderBottom: '1px solid #ddd', paddingBottom: 5 }}>Task Information ({selectedTask.type === 'main' ? 'Main' : 'Sub'})</h3>

                            {renderTaskDetailRow('Description', selectedTask.description || 'No description', 'Description', selectedTask.type === 'main' ? isAdmin : true, selectedTask.description)}

                            {renderTaskDetailRow('Status', (
                                <span style={{
                                    padding: '2px 8px',
                                    borderRadius: 12,
                                    background: selectedTask.status === 'Completed' ? '#dff6dd' : selectedTask.status === 'Overdue' ? '#fde7e9' : '#e1f0fa',
                                    color: selectedTask.status === 'Completed' ? '#107c10' : selectedTask.status === 'Overdue' ? '#a80000' : '#0078d4',
                                    fontWeight: 600,
                                    fontSize: 12
                                }}>
                                    {selectedTask.status}
                                </span>
                            ), 'Status', selectedTask.type === 'main' ? isAdmin : true, selectedTask.status)}

                            {renderTaskDetailRow(
                                'Due Date',
                                formatDate(selectedTask.endDate),
                                'DueDate',
                                selectedTask.type === 'main' ? isAdmin : true, // Main: Admin only, Sub: All
                                selectedTask.endDate
                            )}

                            {renderTaskDetailRow(
                                'Assignee',
                                selectedTask.assignee || 'Unassigned',
                                'Assignee',
                                selectedTask.type !== 'main', // Allow editing for both 'sub' and 'sub-sub' (anything not main)
                                selectedTask.assigneeId
                            )}
                        </div>
                    )}

                    <h3 style={{ marginBottom: 10 }}>Correspondence History</h3>
                    <div style={{ borderTop: '1px solid #eee', paddingTop: 10 }}>
                        {loadingCorrespondence ? (
                            <div>Loading history...</div>
                        ) : correspondenceHistory.length === 0 ? (
                            <div>No correspondence logs found.</div>
                        ) : (
                            correspondenceHistory.map((msg, idx) => (
                                <ActivityItem
                                    key={idx}
                                    activityDescription={[
                                        <span key={1} style={{ fontWeight: 'bold' }}>{msg.Author?.Title || "Unknown"}</span>,
                                        <span key={2}> &bull; {new Date(msg.Created).toLocaleString()}</span>
                                    ]}
                                    comments={
                                        <div style={{ marginTop: 8 }}>
                                            <div style={{ fontWeight: 600 }}>{msg.Title}</div>
                                            <div dangerouslySetInnerHTML={{ __html: sanitizeHtml(msg.MessageBody) }} />
                                        </div>
                                    }
                                    activityIcon={<Icon iconName="Message" />}
                                    style={{ marginBottom: 20 }}
                                />
                            ))
                        )}
                    </div>
                </div>
                <DialogFooter>
                    <DefaultButton onClick={() => setIsCorrespondenceOpen(false)} text="Close" />
                </DialogFooter>
            </Dialog>
        </div>
    );
};
