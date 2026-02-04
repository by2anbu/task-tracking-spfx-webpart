import * as React from 'react';
import styles from '../TaskTracking.module.scss';
import { DetailsList, SelectionMode, IColumn, Stack, Dropdown, DefaultButton, ScrollablePane, ScrollbarVisibility, Pivot, PivotItem, Panel, PanelType } from 'office-ui-fabric-react';
import { taskService } from '../../../../services/sp-service';
import { IMainTask, ISubTask } from '../../../../services/interfaces';
import { TaskDetail } from './TaskDetail';
import modernStyles from './ModernDashboard.module.scss';
import { StatCard } from './StatCard';
import { SmartTaskBoard } from './SmartTaskBoard';
import { PerformanceWidget } from './PerformanceWidget';

export interface IUserDashboardProps {
    userEmail: string;
}

export const UserDashboard: React.FunctionComponent<IUserDashboardProps> = (props) => {
    const [mainTasks, setMainTasks] = React.useState<IMainTask[]>([]);
    const [subTasks, setSubTasks] = React.useState<ISubTask[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);
    const [selectedTask, setSelectedTask] = React.useState<IMainTask | null>(null);
    const [errorMsg, setErrorMsg] = React.useState<string | undefined>(undefined);

    // Filter state
    const [categoryFilter, setCategoryFilter] = React.useState<string | undefined>(undefined);
    const [statusFilter, setStatusFilter] = React.useState<string | undefined>(undefined);

    React.useEffect(() => {
        loadData();
    }, [props.userEmail]);

    const loadData = async () => {
        setLoading(true);
        try {
            const [tasks, subs] = await Promise.all([
                taskService.getMainTasksForUser(props.userEmail),
                taskService.getSubTasksForUser(props.userEmail)
            ]);
            setMainTasks(tasks);
            setSubTasks(subs);
        } catch (e: any) {
            console.error(e);
            setErrorMsg(e.message || "Failed to load data");
        } finally {
            setLoading(false);
        }
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

    // Get unique values
    const getUniqueValues = (arr: string[]): string[] => {
        const seen: { [key: string]: boolean } = {};
        return arr.filter((item) => {
            if (seen[item]) return false;
            seen[item] = true;
            return true;
        });
    };

    // Status counts
    const getStatusCounts = () => {
        const counts: { [key: string]: number } = {
            'Not Started': 0,
            'In Progress': 0,
            'Completed': 0,
            'On Hold': 0
        };
        mainTasks.forEach(t => {
            const status = t.Status || 'Not Started';
            if (counts[status] !== undefined) counts[status]++;
        });
        return counts;
    };

    // Category counts from subtasks
    const getCategoryCounts = () => {
        const counts: { [key: string]: number } = {};
        subTasks.forEach(t => {
            const cat = t.Category || 'Unknown';
            if (!counts[cat]) counts[cat] = 0;
            counts[cat]++;
        });
        return counts;
    };

    // Get overdue tasks
    const getOverdueTasks = () => {
        return mainTasks.filter(t => {
            if (t.Status === 'Completed' || !t.TaskDueDate) return false;
            return new Date(t.TaskDueDate) < new Date();
        });
    };

    const statusCounts = getStatusCounts();
    const categoryCounts = getCategoryCounts();
    const overdueTasks = getOverdueTasks();
    const categoryList = Object.keys(categoryCounts);

    // Filter options
    const statusOptions = [{ key: '', text: 'All Status' }].concat(Object.keys(statusCounts).map(s => ({ key: s, text: s })));
    const categoryOptions = [{ key: '', text: 'All Categories' }].concat(categoryList.map(c => ({ key: c, text: c })));

    // Filtered tasks
    const getFilteredTasks = (): IMainTask[] => {
        let filtered = mainTasks;
        if (statusFilter) {
            filtered = filtered.filter(t => (t.Status || 'Not Started') === statusFilter);
        }
        return filtered;
    };

    const getFilteredSubTasks = (): ISubTask[] => {
        let filtered = subTasks;
        if (categoryFilter) {
            filtered = filtered.filter(t => (t.Category || 'Unknown') === categoryFilter);
        }
        return filtered;
    };

    const filteredTasks = getFilteredTasks();
    const filteredSubTasks = getFilteredSubTasks();

    const resetFilters = () => {
        setStatusFilter(undefined);
        setCategoryFilter(undefined);
    };

    const getStatusClass = (status: string) => {
        const s = (status || '').replace(/\s+/g, '');
        if (!styles) return '';
        return (styles as any)[`status_${s}`] || styles.status_NotStarted;
    };

    const columns: IColumn[] = [
        { key: 'Id', name: '#', fieldName: 'Id', minWidth: 30, maxWidth: 50 },
        { key: 'Title', name: 'Task', fieldName: 'Title', minWidth: 150, isResizable: true },
        {
            key: 'Status', name: 'Status', minWidth: 100,
            onRender: (item: IMainTask) => (
                <span className={`${styles.statusBadge} ${getStatusClass(item.Status || '')}`}>
                    {item.Status || 'Not Started'}
                </span>
            )
        },
        { key: 'TaskDueDate', name: 'Due Date', minWidth: 100, onRender: (i: IMainTask) => formatDate(i.TaskDueDate) },
    ];

    // --- Modern Dashboard Logic ---
    const total = mainTasks.length || 0;
    const completed = statusCounts.Completed || 0;
    const inProgress = statusCounts['In Progress'] || 0;
    const overdue = overdueTasks.length || 0;

    // Calculate Efficiency (Mock logic for now, could be based on on-time completion)
    const efficiency = total > 0 ? Math.round((completed / total) * 100) : 0;

    return (
        <div className={modernStyles.modernDashboard}>
            {/* 1. Header Section */}
            <div className={modernStyles.header}>
                <div className={modernStyles.greeting}>
                    <h1>Good Morning!</h1>
                    <p>Here&apos;s what&apos;s happening with your projects today.</p>
                </div>
                <div className={modernStyles.actions}>
                    {/* Actions to be implemented later */}
                </div>
            </div>

            {/* 2. Stats Grid */}
            <div className={modernStyles.statsGrid}>
                <StatCard
                    title="Total Tasks"
                    value={total}
                    iconName="TaskList"
                    trend="+2 new today"
                    trendDirection="positive"
                    colorClass="iconBlue"
                />
                <StatCard
                    title="Efficiency"
                    value={`${efficiency}%`}
                    iconName="SpeedHigh"
                    trend="On track"
                    trendDirection="neutral"
                    colorClass="iconGreen"
                />
                <StatCard
                    title="In Progress"
                    value={inProgress}
                    iconName="SyncStatusSolid"
                    colorClass="iconOrange"
                />
                <StatCard
                    title="Overdue"
                    value={overdue}
                    iconName="WarningSolid"
                    trend={overdue > 0 ? "Needs attention" : "All good"}
                    trendDirection={overdue > 0 ? "negative" : "positive"}
                    colorClass="iconRed"
                    onClick={() => setStatusFilter("Overdue")} // Simple filter trigger
                />
            </div>

            {/* 3. Main Content Area */}
            <div className={modernStyles.mainContent}>
                {/* Left Column: Smart Task Board */}
                <div>
                    <SmartTaskBoard
                        tasks={filteredTasks}
                        onTaskClick={(t) => setSelectedTask(t)}
                    />
                </div>

                {/* Right Column: Analytics & Quick Views */}
                <Stack tokens={{ childrenGap: 24 }}>
                    {/* Activity / Categories */}
                    <PerformanceWidget
                        categoryCounts={categoryCounts}
                        totalTasks={total}
                    />

                    {/* Timeline / Upcoming */}
                    <div className={modernStyles.sectionCard}>
                        <div className={modernStyles.cardHeader}>
                            <h3>Upcoming Deadlines</h3>
                        </div>
                        <div className={modernStyles.cardBody}>
                            <Stack tokens={{ childrenGap: 12 }}>
                                {mainTasks
                                    .filter(t => t.TaskDueDate && new Date(t.TaskDueDate) > new Date())
                                    .slice(0, 5)
                                    .map(t => (
                                        <div key={t.Id} style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                                            <div style={{ width: 4, height: 40, background: '#0078d4', borderRadius: 4 }} />
                                            <div>
                                                <div style={{ fontWeight: 600, fontSize: 14 }}>{t.Title}</div>
                                                <div style={{ fontSize: 12, color: '#6b7280' }}>{formatDate(t.TaskDueDate)}</div>
                                            </div>
                                        </div>
                                    ))}
                            </Stack>
                        </div>
                    </div>
                </Stack>
            </div>

            {/* Task Detail Panel (Preserved) */}
            <Panel
                isOpen={!!selectedTask}
                onDismiss={() => setSelectedTask(null)}
                type={PanelType.medium}
                headerText=""
            >
                {selectedTask && <TaskDetail mainTask={selectedTask} />}
            </Panel>
        </div>
    );
};
