import * as React from 'react';
import { IMainTask } from '../../../../services/interfaces';
import { Stack, IStackStyles, IStackTokens, Text, DetailsList, SelectionMode, PrimaryButton, IColumn } from 'office-ui-fabric-react';


export interface IAdminReportsProps {
    tasks: IMainTask[];
}

const cardStyles: IStackStyles = {
    root: {
        background: 'white',
        padding: 20,
        borderRadius: 4,
        boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
        minWidth: 200,
        flex: 1
    }
};

const stackTokens: IStackTokens = { childrenGap: 20 };

export const AdminReports: React.FunctionComponent<IAdminReportsProps> = ({ tasks }) => {
    const [filterType, setFilterType] = React.useState<string | undefined>(undefined);
    const [filterValue, setFilterValue] = React.useState<string | undefined>(undefined);

    const calculateStats = () => {
        const total = tasks.length;
        const completed = tasks.filter(t => t.Status === 'Completed').length;
        const inProgress = tasks.filter(t => t.Status === 'In Progress').length;
        const notStarted = tasks.filter(t => t.Status === 'Not Started').length;

        // Overdue: Due Date < Now AND Status != Completed
        const today = new Date();
        const overdue = tasks.filter(t => {
            if (t.Status === 'Completed' || !t.TaskDueDate) return false;
            return new Date(t.TaskDueDate) < today;
        }).length;

        return { total, completed, inProgress, notStarted, overdue };
    };

    const stats = calculateStats();

    // Group by status for "Graph"
    const getStatusData = () => {
        // Just reuse stats
        return [
            { label: 'Not Started', value: stats.notStarted, color: '#d0d0d0' },
            { label: 'In Progress', value: stats.inProgress, color: '#0078d4' },
            { label: 'Completed', value: stats.completed, color: '#107c10' },
            { label: 'Overdue', value: stats.overdue, color: '#a80000' }
        ];
    };

    // Helper to extract user name
    const getUserName = (t: IMainTask) => {
        let uName = 'Unassigned';
        const assignee = t.TaskAssignedTo;
        if (Array.isArray(assignee) && assignee.length > 0) uName = assignee[0].Title;
        else if (typeof assignee === 'object' && (assignee as any).Title) uName = (assignee as any).Title;
        return uName;
    };

    // Group by User
    const getTasksByUser = () => {
        const groups: { [key: string]: number } = {};
        tasks.forEach(t => {
            const uName = getUserName(t);
            groups[uName] = (groups[uName] || 0) + 1;
        });
        return Object.keys(groups).map(key => ({ label: key, value: groups[key], color: '#605e5c' }));
    };

    // Group by Business Unit
    const getBusinessUnitData = () => {
        const groups: { [key: string]: number } = {};
        tasks.forEach(t => {
            const bu = t.BusinessUnit || 'Unknown';
            groups[bu] = (groups[bu] || 0) + 1;
        });
        return Object.keys(groups).map(key => ({ label: key, value: groups[key] }));
    };

    const BarChart = ({ data, title, onBarClick }: { data: { label: string, value: number, color?: string }[], title: string, onBarClick: (label: string) => void }) => {
        const maxVal = Math.max(...data.map(d => d.value));
        return (
            <div style={{ background: 'white', padding: 20, borderRadius: 4, boxShadow: '0 2px 4px rgba(0,0,0,0.1)', flex: 1, minWidth: 300 }}>
                <h3 style={{ marginTop: 0 }}>{title}</h3>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                    {data.map((d, i) => (
                        <div
                            key={i}
                            style={{ display: 'flex', alignItems: 'center', fontSize: 12, cursor: 'pointer' }}
                            onClick={() => onBarClick(d.label)}
                            title={`Filter by ${d.label}`}
                        >
                            <div style={{ width: 100, fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{d.label}</div>
                            <div style={{ flexGrow: 1, backgroundColor: '#f3f2f1', height: 16, borderRadius: 4, overflow: 'hidden', margin: '0 10px' }}>
                                <div style={{
                                    width: `${maxVal > 0 ? (d.value / maxVal) * 100 : 0}%`,
                                    backgroundColor: d.color || '#0078d4',
                                    height: '100%',
                                    transition: 'width 0.5s ease'
                                }} />
                            </div>
                            <div style={{ width: 30, textAlign: 'right' }}>{d.value}</div>
                        </div>
                    ))}
                </div>
            </div>
        );
    };

    const getFilteredTasks = () => {
        if (!filterType || !filterValue) return tasks;
        return tasks.filter(t => {
            if (filterType === 'status') {
                if (filterValue === 'Overdue') {
                    if (t.Status === 'Completed' || !t.TaskDueDate) return false;
                    return new Date(t.TaskDueDate) < new Date();
                }
                return t.Status === filterValue;
            }
            if (filterType === 'bu') return (t.BusinessUnit || 'Unknown') === filterValue;
            if (filterType === 'user') return getUserName(t) === filterValue;
            return true;
        });
    };

    const filteredTasksForGrid = getFilteredTasks();

    const exportToExcel = () => {
        const headers = ['ID', 'Task', 'Assigned To', 'Status', 'Due Date'];
        const rows = filteredTasksForGrid.map(t => {
            const uName = getUserName(t);
            return [
                t.Id,
                `"${(t.Title || '').replace(/"/g, '""')}"`,
                `"${uName}"`,
                t.Status,
                t.TaskDueDate ? new Date(t.TaskDueDate).toLocaleDateString() : ''
            ].join(',');
        });
        const csvContent = [headers.join(',')].concat(rows).join('\r\n');
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', 'Admin_Report_Export.csv');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    // Columns for the grid
    const columns: IColumn[] = [
        { key: 'id', name: 'ID', fieldName: 'Id', minWidth: 30, maxWidth: 50 },
        { key: 'title', name: 'Task', fieldName: 'Title', minWidth: 150, isResizable: true },
        {
            key: 'assigned', name: 'Assigned To', minWidth: 100,
            onRender: (t: IMainTask) => getUserName(t)
        },
        { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100 },
        { key: 'due', name: 'Due Date', minWidth: 100, onRender: (t: IMainTask) => t.TaskDueDate ? new Date(t.TaskDueDate).toLocaleDateString() : '' }
    ];

    const handleFilter = (type: string, value: string) => {
        // Toggle if same
        if (filterType === type && filterValue === value) {
            setFilterType(undefined);
            setFilterValue(undefined);
        } else {
            setFilterType(type);
            setFilterValue(value);
        }
    };

    return (
        <Stack tokens={stackTokens}>
            {/* KPI Cards */}
            <Stack horizontal wrap tokens={stackTokens}>
                <Stack styles={cardStyles}>
                    <Text variant="large" styles={{ root: { fontWeight: 600, color: '#666' } }}>Total Tasks</Text>
                    <Text variant="xxLarge" styles={{ root: { color: '#333' } }}>{stats.total}</Text>
                </Stack>
                <Stack styles={cardStyles}>
                    <Text variant="large" styles={{ root: { fontWeight: 600, color: '#107c10' } }}>Completed</Text>
                    <Text variant="xxLarge" styles={{ root: { color: '#107c10' } }}>{stats.completed}</Text>
                </Stack>
                <Stack styles={cardStyles}>
                    <Text variant="large" styles={{ root: { fontWeight: 600, color: '#0078d4' } }}>In Progress</Text>
                    <Text variant="xxLarge" styles={{ root: { color: '#0078d4' } }}>{stats.inProgress}</Text>
                </Stack>
                <Stack styles={cardStyles}>
                    <Text variant="large" styles={{ root: { fontWeight: 600, color: '#a80000' } }}>Overdue</Text>
                    <Text variant="xxLarge" styles={{ root: { color: '#a80000' } }}>{stats.overdue}</Text>
                </Stack>
            </Stack>

            {/* Charts Row */}
            <Stack horizontal wrap tokens={stackTokens}>
                <BarChart data={getStatusData()} title="Tasks by Status" onBarClick={(v) => handleFilter('status', v)} />
                <BarChart data={getTasksByUser()} title="Tasks by User" onBarClick={(v) => handleFilter('user', v)} />
                <BarChart data={getBusinessUnitData()} title="Tasks by Business Unit" onBarClick={(v) => handleFilter('bu', v)} />
            </Stack>

            {/* Grid Section */}
            <div style={{ background: 'white', padding: 20, borderRadius: 4, boxShadow: '0 2px 4px rgba(0,0,0,0.1)', marginTop: 20 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 15 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                        <Text variant="xLarge">Task Details</Text>
                        {filterType && (
                            <div style={{ backgroundColor: '#e1dfdd', padding: '4px 8px', borderRadius: 4, display: 'flex', alignItems: 'center', fontSize: 12 }}>
                                <span>Filtered by {filterType === 'bu' ? 'Business Unit' : filterType === 'user' ? 'User' : 'Status'}: <strong>{filterValue}</strong></span>
                                <span
                                    style={{ marginLeft: 8, cursor: 'pointer', fontWeight: 'bold' }}
                                    onClick={() => handleFilter('', '')}
                                    title="Clear Filter"
                                >
                                    X
                                </span>
                            </div>
                        )}
                    </Stack>
                    <PrimaryButton text="Export to Excel" iconProps={{ iconName: 'ExcelDocument' }} onClick={exportToExcel} />
                </div>
                <DetailsList
                    items={filteredTasksForGrid}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    compact={true}
                />
            </div>
        </Stack>
    );
};
