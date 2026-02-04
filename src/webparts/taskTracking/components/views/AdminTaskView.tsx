
import * as React from 'react';
import { DetailsList, SelectionMode, IColumn, IconButton, Stack, MessageBar, MessageBarType, Panel, PanelType, TextField, PrimaryButton, DefaultButton, ComboBox, IComboBoxOption, DatePicker, SelectableOptionMenuItemType, Persona, PersonaSize, IDropdownOption, Sticky, StickyPositionType, ScrollablePane, ScrollbarVisibility, DetailsListLayoutMode, ConstrainMode, IDetailsHeaderProps } from 'office-ui-fabric-react';
import { taskService } from '../../../../services/sp-service';
import { IMainTask, ISubTask, LIST_SUB_TASKS } from '../../../../services/interfaces';

export interface IAdminTaskViewProps {
    userEmail: string;
}

// Unified Node Interface
export interface IUiHierarchyNode {
    nodeId: string; // "M-1" or "S-10"
    originalId: number;
    type: 'Main' | 'Sub';
    title: string;
    description: string;
    status: string;
    assignedTo: any[];
    dueDate?: string;
    category?: string;

    depth: number;
    hasChildren: boolean;

    // Links
    parentId?: string; // nodeId of parent

    // Original Data
    data: IMainTask | ISubTask;
}

export const AdminTaskView: React.FunctionComponent<IAdminTaskViewProps> = (props) => {
    const [nodes, setNodes] = React.useState<IUiHierarchyNode[]>([]);
    const [loading, setLoading] = React.useState<boolean>(true);
    const [expandedRows, setExpandedRows] = React.useState<Set<string>>(new Set());

    // Edit/View State (Reused from SubtaskView logic, simplified for now)
    const [selectedNode, setSelectedNode] = React.useState<IUiHierarchyNode | null>(null);
    const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);

    React.useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {
        setLoading(true);
        try {
            const mainTasks = await taskService.getAllMainTasks();
            const subTasks = await taskService.getAllSubTasks();

            const processed = buildHierarchy(mainTasks, subTasks);
            setNodes(processed);
        } catch (error) {
            console.error(error);
        } finally {
            setLoading(false);
        }
    };

    const buildHierarchy = (mainTasks: IMainTask[], subTasks: ISubTask[]): IUiHierarchyNode[] => {
        const nodeList: IUiHierarchyNode[] = [];

        // Maps
        const mainMap = new Map<number, IMainTask>();
        mainTasks.forEach(m => mainMap.set(m.Id, m));

        const subMap = new Map<number, ISubTask>();
        const subChildrenMap = new Map<number, ISubTask[]>(); // ParentSubtaskId -> Children
        const mainChildrenMap = new Map<number, ISubTask[]>(); // Admin_Job_ID -> Level 1 Subtasks

        subTasks.forEach(s => {
            subMap.set(s.Id, s);

            const pId = s.ParentSubtaskId;
            if (pId && pId > 0) {
                // It is a sub-subtask
                if (!subChildrenMap.has(pId)) subChildrenMap.set(pId, []);
                subChildrenMap.get(pId)!.push(s);
            } else {
                // It is a Level 1 subtask, belongs to Main Task
                const mId = s.Admin_Job_ID;
                if (!mainChildrenMap.has(mId)) mainChildrenMap.set(mId, []);
                mainChildrenMap.get(mId)!.push(s);
            }
        });

        // Function to process Subtask Recursively
        const processSubtree = (sub: ISubTask, depth: number, result: IUiHierarchyNode[], parentNodeId: string) => {
            const nodeId = `S-${sub.Id}`;
            const children = subChildrenMap.get(sub.Id) || [];

            result.push({
                nodeId: nodeId,
                originalId: sub.Id,
                type: 'Sub',
                title: sub.Task_Title || sub.Title,
                description: sub.Task_Description,
                status: sub.TaskStatus,
                assignedTo: sub.TaskAssignedTo,
                dueDate: sub.TaskDueDate,
                category: sub.Category,
                depth: depth,
                hasChildren: children.length > 0,
                parentId: parentNodeId, // Link to parent
                data: sub
            });

            children.sort((a, b) => a.Id - b.Id);
            children.forEach(child => processSubtree(child, depth + 1, result, nodeId));
        };

        // Build Tree starting from Main Tasks
        mainTasks.forEach(m => {
            const nodeId = `M-${m.Id}`;
            const children = mainChildrenMap.get(m.Id) || [];

            nodeList.push({
                nodeId: nodeId,
                originalId: m.Id,
                type: 'Main',
                title: m.Title, // Main Task Title
                description: m.Task_x0020_Description,
                status: m.Status,
                assignedTo: m.TaskAssignedTo,
                dueDate: m.TaskDueDate,
                depth: 0,
                hasChildren: children.length > 0,
                data: m
            });

            children.sort((a, b) => a.Id - b.Id);
            children.forEach(child => processSubtree(child, 1, nodeList, nodeId));
        });

        return nodeList;
    };

    const toggleExpand = (nodeId: string) => {
        const newExpanded = new Set<string>();
        expandedRows.forEach(r => newExpanded.add(r));

        if (newExpanded.has(nodeId)) {
            newExpanded.delete(nodeId);
        } else {
            newExpanded.add(nodeId);
        }
        setExpandedRows(newExpanded);
    };

    // Rendering Helper which respects expansion
    const getVisibleNodes = (): IUiHierarchyNode[] => {
        // Since list is DFS ordered
        const visible: IUiHierarchyNode[] = [];
        const visibleParents = new Set<string>();

        // Logic: Root (depth 0) is always visible.
        // Child is visible if Parent is visible AND Parent is Expanded.

        // We need to know who the parent IS to check this efficiently.
        // But our flat list doesn't explicitly store parent Node ID in the loop easily without lookups.
        // Actually, relies on Depth order.

        // Let's track the "Current Visible Path".
        // But sibling subtrees make this tricky.

        // Simpler: iterate. If depth > 0, we need to find if its parent is expanded.
        // We didn't store parentNodeId in the object efficiently above.
        // Let's fix buildHierarchy to include parentId?
        // Or simpler: We know Main Tasks (depth 0) are roots.

        // Wait, standard approach:
        // Set of Expanded IDs.
        // Check if all ancestors are expanded.

        // Let's assume we re-run filter.
        // To do this fast:
        // The list is structurally sorted: Parent -> Children

        // We can maintain a "Stack" of visible parents?
        // Or just a set of currently visible parent IDs.

        // NOTE: Main Tasks start with "M-". Subtasks with "S-".

        // We need the PARENT ID for each node to check against `visibleParents`.
        // I will add `parentId` to IUiHierarchyNode (logic needs update).

        // RE-IMPLEMENT buildHierarchy to include parentId, then logic works.
        // For now, let's just create the component. Ideally I'll update it before closing this file.
        // Actually, let's assume I fix it below in a second tool call or just logic fix now.
        // I'll update `buildHierarchy` in this same `write_to_file`.

        return nodes; // Placeholder, applied in render filtering
    };

    // Format helpers
    const getAssigneeName = (assigned: any[] | any): string => {
        if (Array.isArray(assigned)) return assigned.map(u => u.Title).join(', ');
        if (assigned && assigned.Title) return assigned.Title;
        return '';
    };

    const formatDate = (date: string | undefined): string => {
        if (!date) return '';
        const d = new Date(date);
        return d.toLocaleDateString();
    };

    const columns: IColumn[] = [
        {
            key: 'title', name: 'Task Hierarchy', minWidth: 250, maxWidth: 400,
            onRender: (item: IUiHierarchyNode) => {
                const indent = item.depth * 20;
                const isExpanded = expandedRows.has(item.nodeId);
                const isMain = item.type === 'Main';

                return (
                    <div style={{ display: 'flex', alignItems: 'center', paddingLeft: indent }}>
                        {item.hasChildren ? (
                            <IconButton
                                iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                                styles={{ root: { height: 24, width: 24, marginRight: 4 } }}
                                onClick={(e) => { e.stopPropagation(); toggleExpand(item.nodeId); }}
                            />
                        ) : <div style={{ width: 28 }} />}

                        <span style={{
                            fontWeight: isMain ? 700 : (item.hasChildren ? 600 : 400),
                            fontSize: isMain ? '14px' : '13px',
                            color: isMain ? '#0078d4' : 'inherit'
                        }}>
                            {item.title}
                            {isMain && <span style={{ fontSize: '10px', color: '#666', marginLeft: 5 }}>(Ref: {item.originalId})</span>}
                        </span>
                    </div>
                );
            }
        },
        { key: 'status', name: 'Status', minWidth: 100, onRender: (i: IUiHierarchyNode) => i.status },
        { key: 'assigned', name: 'Assigned To', minWidth: 150, onRender: (i: IUiHierarchyNode) => getAssigneeName(i.assignedTo) },
        { key: 'due', name: 'Due Date', minWidth: 100, onRender: (i: IUiHierarchyNode) => formatDate(i.dueDate) },
        { key: 'desc', name: 'Description', minWidth: 200, isResizable: true, onRender: (i: IUiHierarchyNode) => i.description }
    ];

    // Filter Logic for Rendering
    const nodesToRender = nodes.filter(node => {
        // Root is always visible
        if (node.depth === 0) return true;

        // For others, we need to check parents.
        // This is inefficient O(N*Depth) if we walk up.
        // Better: Pre-tag visibility?
        // Or simple: Just traverse visible list logic?

        // Since we didn't store Parent Node ID cleanly string-wise in buildHierarchy (yet),
        // let's do a trick: we rely on `processed` order.
        return true;
    });

    // BETTER Filtering:
    // Let's redo the filter properly.
    const visibleNodes: IUiHierarchyNode[] = [];
    const visibleParentIds = new Set<string>(); // nodeIds that are visible AND expanded

    // To make this work, we need `parentId` in the node. 
    // I defined it in interface. I need to populate it.

    // Re-run population logic in render would be slow? No, logic is fast.
    // Actually, `nodes` is flat list.
    // If I ensured `parentId` is set, I can do:
    for (const node of nodes) {
        if (node.depth === 0) {
            visibleNodes.push(node);
            if (expandedRows.has(node.nodeId)) visibleParentIds.add(node.nodeId);
        } else {
            if (node.parentId && visibleParentIds.has(node.parentId)) {
                visibleNodes.push(node);
                if (expandedRows.has(node.nodeId)) visibleParentIds.add(node.nodeId);
            }
        }
    }


    const onRenderDetailsHeader = (props: IDetailsHeaderProps, defaultRender?: (props: IDetailsHeaderProps) => JSX.Element | null): JSX.Element | null => {
        if (!props || !defaultRender) {
            return null;
        }
        return (
            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                {defaultRender({
                    ...props,
                    onRenderColumnHeaderTooltip: (tooltipHostProps) => {
                        if (!tooltipHostProps) return null;
                        return <span className={tooltipHostProps.className}>{tooltipHostProps.content}</span>;
                    }
                })}
            </Sticky>
        );
    };

    return (
        <div style={{ padding: 20 }}>
            <h2 style={{ marginBottom: 20 }}>All Subtasks (Admin View)</h2>
            <div style={{ position: 'relative', height: '70vh', border: '1px solid #edebe9', borderRadius: 4 }}>
                <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                    <DetailsList
                        items={visibleNodes}
                        columns={columns}
                        selectionMode={SelectionMode.none}
                        compact={true}
                        onRenderDetailsHeader={onRenderDetailsHeader}
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        constrainMode={ConstrainMode.unconstrained}
                    />
                </ScrollablePane>
            </div>
        </div>
    );
};
