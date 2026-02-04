import * as React from 'react';
import styles from './ModernDashboard.module.scss';
import { IMainTask } from '../../../../services/interfaces';
import { DetailsList, SelectionMode, IColumn, Icon } from 'office-ui-fabric-react';

const s = styles as any; // Cast for module access

export interface ISmartTaskBoardProps {
    tasks: IMainTask[];
    onTaskClick: (task: IMainTask) => void;
}

export const SmartTaskBoard: React.FC<ISmartTaskBoardProps> = ({ tasks, onTaskClick }) => {
    const [viewMode, setViewMode] = React.useState<'list' | 'kanban'>('kanban');

    // Helper to format date
    const formatDate = (dateStr?: string) => dateStr ? new Date(dateStr).toLocaleDateString() : '-';

    // Helper to get Assigned User Name
    const getAssignedUser = (users: any) => {
        if (!users) return 'Unassigned';
        if (Array.isArray(users)) {
            return users.length > 0 ? users.map((u: any) => u.Title).join(', ') : 'Unassigned';
        }
        // Handle single object case
        return users.Title || 'Unassigned';
    }

    // Columns for List View
    const columns: IColumn[] = [
        { key: 'Title', name: 'Task', fieldName: 'Title', minWidth: 180, isResizable: true },
        {
            key: 'AssignedTo', name: 'Assigned To', minWidth: 120, isResizable: true,
            onRender: (i: IMainTask) => getAssignedUser(i.TaskAssignedTo)
        },
        { key: 'Project', name: 'Project', fieldName: 'Project', minWidth: 100, isResizable: true },
        {
            key: 'StartDate', name: 'Start Date', minWidth: 90,
            onRender: (i: IMainTask) => formatDate(i.TaskStartDate)
        },
        {
            key: 'DueDate', name: 'Due Date', minWidth: 90,
            onRender: (i: IMainTask) => formatDate(i.TaskDueDate)
        },
        {
            key: 'EndDate', name: 'End Date', minWidth: 90,
            onRender: (i: IMainTask) => formatDate(i.Task_x0020_End_x0020_Date)
        },
        { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100 }
    ];

    // Kanban Logic
    const statuses = ['Not Started', 'In Progress', 'Completed'];
    const getTasksByStatus = (status: string) => tasks.filter(t => (t.Status || 'Not Started') === status);

    return (
        <div className={s.boardContainer}>
            {/* View Toggles */}
            <div className={s.viewToggle}>
                <button
                    className={viewMode === 'kanban' ? s.active : ''}
                    onClick={() => setViewMode('kanban')}>
                    <Icon iconName="Board" style={{ marginRight: 6 }} /> Board
                </button>
                <button
                    className={viewMode === 'list' ? s.active : ''}
                    onClick={() => setViewMode('list')}>
                    <Icon iconName="List" style={{ marginRight: 6 }} /> List
                </button>
            </div>

            {/* Content */}
            {viewMode === 'list' ? (
                <div style={{ background: 'white', borderRadius: 12, padding: 0, overflowX: 'auto' }}>
                    <DetailsList
                        items={tasks}
                        columns={columns}
                        selectionMode={SelectionMode.none} // Click handled by row click?
                        onActiveItemChanged={onTaskClick}
                        styles={{
                            root: { selectors: { '.ms-DetailsRow': { background: 'transparent' } } }
                        }}
                    />
                </div>
            ) : (
                <div className={s.kanbanBoard}>
                    {statuses.map(status => (
                        <div key={status} className={s.kanbanColumn}>
                            <h4>
                                {status}
                                <span className={s.count}>{getTasksByStatus(status).length}</span>
                            </h4>
                            {getTasksByStatus(status).map(task => (
                                <div key={task.Id} className={s.kanbanCard} onClick={() => onTaskClick(task)}>
                                    <div style={{ fontWeight: 600, marginBottom: 4 }}>{task.Title}</div>
                                    <div className={s.cardMeta}>
                                        <span>#{task.Id}</span>
                                        {task.TaskDueDate && (
                                            <span style={{ color: new Date(task.TaskDueDate) < new Date() ? '#a80000' : 'inherit' }}>
                                                {new Date(task.TaskDueDate).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}
                                            </span>
                                        )}
                                    </div>
                                </div>
                            ))}
                        </div>
                    ))}
                </div>
            )}
        </div>
    );
};
