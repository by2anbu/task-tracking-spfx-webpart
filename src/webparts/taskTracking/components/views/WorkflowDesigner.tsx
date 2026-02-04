/* eslint-disable max-lines */
import * as React from 'react';
import { useState, useCallback, useEffect } from 'react';
import ReactFlow, {
    Node,
    Edge,
    Controls,
    Background,
    BackgroundVariant,
    addEdge,
    Connection,
    applyNodeChanges,
    applyEdgeChanges,
    NodeChange,
    EdgeChange,
    NodeTypes,
    MiniMap,
    MarkerType,
    Handle,
    Position,
    ReactFlowProvider,
    useReactFlow
} from 'react-flow-renderer';
import { toPng } from 'html-to-image';
import {
    TextField,
    ComboBox,
    IComboBoxOption,
    Dropdown,
    IDropdownOption,
    Stack,
    PrimaryButton,
    DefaultButton,
    Label,
    Dialog,
    DialogType,
    DialogFooter,
    Separator,
    Icon,
    MessageBar,
    MessageBarType,
    Panel,
    PanelType,
    IconButton
} from 'office-ui-fabric-react';
import {
    Play,
    CheckCircle,
    GitBranch,
    Mail,
    Database,
    FileText,
    Clock,
    AlertCircle,
    Bell,
    Layers,
    Plus,
    X,
    Layout,
    RefreshCw,
    Maximize,
    Link as LinkIcon,
    Info,
    AlertTriangle,
    Search,
    ChevronRight,
    Copy,
    Trash2,
    DraftingCompass,
    PlusCircle,
    Zap,
    Lock,
    Download
} from 'lucide-react';
import styles from './WorkflowDesigner.module.scss';
import { taskService } from '../../../../services/sp-service';
import { LIST_MAIN_TASKS, LIST_SUB_TASKS } from '../../../../services/interfaces';
import { NodePropertiesDialog } from './NodePropertiesDialog';
import { sanitizeHtml } from '../../../../utils/sanitize';


const iconMapping: { [key: string]: any } = {
    'Main Task': Layers,
    'Start': Play,
    'Task': CheckCircle,
    'Condition': GitBranch,
    'Email': Mail,
    'Alert': Bell,
    'Link': LinkIcon
};

// Custom Node Component
interface CustomNodeData {
    label: string;
    type: string;
    description?: string;
    assignee?: string;
    emailTo?: string;
    emailSubject?: string;
    condition?: string;
    dbQuery?: string;
    category?: string;
    // Task Integration
    linkedTaskId?: number;
    status?: string;
    linkUrl?: string;
    onPlusClick?: (e: React.MouseEvent) => void;
    onDrop?: (files: File[]) => void;
    // New Fields for Main Task
    year?: string;
    month?: string;
    department?: string;
    project?: string;
    dueDate?: string;
    // Feature Upgrades
    progress?: number;
    isBlocked?: boolean;
    isHighlighted?: boolean;
}

const iconColorMapping: { [key: string]: string } = {
    'Main Task': '#ff6d5b',
    'Start': '#10b981',
    'Task': '#3b82f6',
    'Condition': '#f59e0b',
    'Email': '#8b5cf6',
    'Alert': '#ef4444',
    'Link': '#0ea5e9',
};

const CustomNode: React.FC<{ data: CustomNodeData; selected?: boolean }> = ({ data, selected }) => {
    const IconComponent = iconMapping[data.type] || Play;
    const nodeColor = iconColorMapping[data.type] || '#64748b';
    const isMainTask = data.type === 'Main Task';

    return (
        <div
            className={`${styles.customNode} ${selected ? styles.selected : ''} ${data.isHighlighted ? styles.highlighted : ''}`}
            onDragOver={(e) => {
                if (data.linkedTaskId) {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'copy';
                }
            }}
            onDrop={async (e) => {
                if (data.linkedTaskId) {
                    e.preventDefault();
                    const files = Array.from(e.dataTransfer.files);
                    if (files.length > 0 && data.onDrop) {
                        data.onDrop(files);
                    }
                }
            }}
        >
            <Handle type="target" position={Position.Left} className={styles.targetHandle} />

            <div
                className={`${styles.iconBox} ${selected ? styles.selected : ''}`}
                style={{
                    borderColor: selected ? nodeColor : (data.isHighlighted ? '#3b82f6' : '#e2e8f0'),
                    backgroundColor: selected ? `${nodeColor}10` : 'white'
                }}
            >
                {/* Progress Ring for Main Task */}
                {isMainTask && data.progress !== undefined && (
                    <svg className={styles.progressRing} width="44" height="44">
                        <circle
                            className={styles.progressRingBg}
                            stroke="#e2e8f0"
                            strokeWidth="3"
                            fill="transparent"
                            r="18"
                            cx="22"
                            cy="22"
                        />
                        <circle
                            className={styles.progressRingIndicator}
                            stroke={nodeColor}
                            strokeWidth="3"
                            strokeDasharray={`${18 * 2 * Math.PI}`}
                            strokeDashoffset={`${18 * 2 * Math.PI * (1 - (data.progress || 0) / 100)}`}
                            strokeLinecap="round"
                            fill="transparent"
                            r="18"
                            cx="22"
                            cy="22"
                        />
                    </svg>
                )}

                {/* Status Ring for Subtasks */}
                {!isMainTask && data.status && (
                    <div className={`${styles.statusRing} ${data.status === 'Completed' ? styles.completed :
                        data.status === 'In Progress' ? styles.inProgress :
                            styles.notStarted
                        }`} />
                )}

                <IconComponent size={24} color={nodeColor} />

                {/* Blocked Indicator */}
                {data.isBlocked && (
                    <div className={styles.blockedBadge} title="Blocked by incomplete parents">
                        <Lock size={10} color="white" />
                    </div>
                )}

                {data.type === 'Link' && data.linkUrl && (
                    <div
                        className={styles.externalLinkIcon}
                        onClick={(e) => {
                            e.stopPropagation();
                            window.open(data.linkUrl, '_blank');
                        }}
                    >
                        <LinkIcon size={12} />
                    </div>
                )}

                <div
                    className={`${styles.plusBtn} nodrag`}
                    onPointerDown={(e) => e.stopPropagation()}
                    onMouseDown={(e) => e.stopPropagation()}
                    onClick={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        if (data.onPlusClick) {
                            data.onPlusClick(e);
                        }
                    }}
                    title="Add connected node"
                >
                    <Plus size={14} strokeWidth={3} />
                </div>
            </div>

            {/* Node Info Container for Animation */}
            <div className={styles.nodeInfo}>
                <div className={styles.nodeLabel}>
                    {data.label}
                    {isMainTask && data.progress !== undefined && (
                        <span className={styles.progressPercent}>{Math.round(data.progress)}%</span>
                    )}
                </div>

                {/* Metadata Wrapper */}
                <div className={styles.nodeMetadata}>
                    {/* Status Badge */}
                    {data.linkedTaskId && data.status && (
                        <div className={`${styles.nodeStatusLabel} ${data.status === 'Completed' ? styles.statusCompleted :
                            data.status === 'In Progress' ? styles.statusInProgress :
                                styles.statusNotStarted
                            }`}>
                            <div className={styles.statusDot} />
                            {data.status}
                        </div>
                    )}

                    {/* Assignee Info */}
                    {data.linkedTaskId && data.assignee && (
                        <div className={styles.nodeAssignee}>
                            <div className={styles.assigneeAvatar}>
                                {data.assignee.charAt(0).toUpperCase()}
                            </div>
                            <span className={styles.assigneeName} title={data.assignee}>
                                {data.assignee.split('@')[0]}
                            </span>
                        </div>
                    )}

                    {/* SharePoint ID Badge */}
                    {data.linkedTaskId && (
                        <div
                            className={styles.nodeStatusBadge}
                            onClick={(e) => {
                                e.stopPropagation();
                                const url = new URL(window.location.href);
                                url.searchParams.set('ViewTaskID', data.linkedTaskId!.toString());
                                window.open(url.toString(), '_blank');
                            }}
                            title="View SharePoint Task"
                        >
                            <Database size={10} />
                            <span>#{data.linkedTaskId}</span>
                        </div>
                    )}
                </div>
            </div>

            <Handle type="source" position={Position.Right} className={styles.sourceHandle} />
        </div>
    );
};

const nodeTypes: NodeTypes = {
    customNode: CustomNode as any
};

// Node templates for the sidebar
const nodeTemplates = [
    { type: 'Main Task', icon: Layers, color: iconColorMapping['Main Task'] },
    { type: 'Task', icon: CheckCircle, color: iconColorMapping.Task },
    { type: 'Condition', icon: GitBranch, color: iconColorMapping.Condition },
];

export interface IWorkflowDesignerProps {
    mainTaskId?: number;
    readonly?: boolean;
    userDisplayName?: string;
    userEmail?: string;
}

export const WorkflowDesigner: React.FC<IWorkflowDesignerProps> = (props) => {
    return (
        <ReactFlowProvider>
            <WorkflowDesignerInternal {...props} />
        </ReactFlowProvider>
    );
};

const WorkflowDesignerInternal: React.FC<IWorkflowDesignerProps> = (props) => {
    const { mainTaskId: propMainTaskId, readonly, userDisplayName, userEmail } = props;
    const [mainTaskId, setMainTaskId] = useState<number | undefined>(propMainTaskId);

    useEffect(() => {
        setMainTaskId(propMainTaskId);
    }, [propMainTaskId]);

    const [nodes, setNodes] = useState<Node[]>([]);
    const [edges, setEdges] = useState<Edge[]>([]);
    const [nodeIdCounter, setNodeIdCounter] = useState(1);
    const [selectedNodeId, setSelectedNodeId] = useState<string | null>(null);
    const [isSaving, setIsSaving] = useState(false);
    const [isSyncing, setIsSyncing] = useState(false);
    const [mainTaskTitle, setMainTaskTitle] = useState<string>('');

    // Dropdown options
    const [userOptions, setUserOptions] = useState<IComboBoxOption[]>([]);
    const [categoryOptions, setCategoryOptions] = useState<IDropdownOption[]>([]);
    const [yearOptions, setYearOptions] = useState<IDropdownOption[]>([]);
    const [monthOptions, setMonthOptions] = useState<IDropdownOption[]>([]);
    const [departmentOptions, setDepartmentOptions] = useState<IDropdownOption[]>([]);

    const { fitView, project } = useReactFlow();

    // UI States for Context Menu & Quick Add
    const [menu, setMenu] = useState<{ x: number, y: number, nodeId?: string } | null>(null);
    const [isWhatsNextOpen, setIsWhatsNextOpen] = useState(false);
    const [whatsNextSourceNode, setWhatsNextSourceNode] = useState<string | null>(null);
    const [librarySearchQuery, setLibrarySearchQuery] = useState('');
    const [isSidebarMinimized, setIsSidebarMinimized] = useState(false);
    const [addingNodeType, setAddingNodeType] = useState<string | null>(null);

    // Create Task Modal State
    const [isCreateTaskOpen, setIsCreateTaskOpen] = useState(false);
    const [newTaskTitle, setNewTaskTitle] = useState('');
    const [newTaskDesc, setNewTaskDesc] = useState('');
    const [newTaskAssignee, setNewTaskAssignee] = useState<string | number | undefined>(undefined);
    const [newTaskDueDate, setNewTaskDueDate] = useState<Date | undefined>(new Date());
    const [parentSubtaskId, setParentSubtaskId] = useState<number | undefined>(undefined);

    // Create Main Task Modal State
    const [isCreateMainTaskOpen, setIsCreateMainTaskOpen] = useState(false);
    const [newMainTaskTitle, setNewMainTaskTitle] = useState('');
    const [newMainTaskDesc, setNewMainTaskDesc] = useState('');
    const [newMainTaskYear, setNewMainTaskYear] = useState<string | undefined>(new Date().getFullYear().toString());
    const [newMainTaskMonth, setNewMainTaskMonth] = useState<string | undefined>(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][new Date().getMonth()]);
    const [newMainTaskAssignee, setNewMainTaskAssignee] = useState<string | undefined>(undefined);
    const [newMainTaskDept, setNewMainTaskDept] = useState<string | undefined>(undefined);
    const [newMainTaskProject, setNewMainTaskProject] = useState('');
    const [newMainTaskDueDate, setNewMainTaskDueDate] = useState<Date | undefined>(new Date());
    const [newMainTaskFiles, setNewMainTaskFiles] = useState<File[]>([]);

    // Clarification States
    const [isClarificationOpen, setIsClarificationOpen] = useState(false);
    const [clarificationHistory, setClarificationHistory] = useState<any[]>([]);
    const [clarificationMessage, setClarificationMessage] = useState('');
    const [isClarificationLoading, setIsClarificationLoading] = useState(false);
    const [currentUserEmail, setCurrentUserEmail] = useState<string>('');
    const hasInitialLoaded = React.useRef(false);
    const lastMainTaskIdRef = React.useRef<number | undefined>(undefined);
    const reactFlowWrapper = React.useRef<HTMLDivElement>(null);

    // -- Toast System --
    const [toasts, setToasts] = useState<{ id: string, title: string, message: string, type: 'success' | 'info' | 'warning' | 'error', removing?: boolean }[]>([]);

    const showToast = useCallback((title: string, message: string, type: 'success' | 'info' | 'warning' | 'error' = 'success') => {
        const id = Math.random().toString(36).substr(2, 9);
        setToasts(prev => [...prev, { id, title, message, type }]);
        setTimeout(() => {
            setToasts(prev => prev.map(t => t.id === id ? { ...t, removing: true } : t));
            setTimeout(() => {
                setToasts(prev => prev.filter(t => t.id !== id));
            }, 500);
        }, 4000);
    }, []);




    // Magic Tidy Up - Improved Horizontal Layout
    const tidyUp = useCallback((currentNodes?: any[], currentEdges?: any[]) => {
        const nodesToProcess = currentNodes || nodes;
        const edgesToProcess = currentEdges || edges;

        const levelGapX = 260;
        const siblingGapY = 140;
        const nodeLevels = new Map<string, number>();
        const visitCount = new Map<string, number>();

        // Find root nodes (nodes with no incoming edges)
        const targetIds = new Set(edgesToProcess.map(e => e.target));
        const rootNodes = nodesToProcess.filter(n => !targetIds.has(n.id) || n.data.type === 'Start');

        const queue: { id: string, level: number }[] = rootNodes.map(n => ({ id: n.id, level: 0 }));
        queue.forEach(q => nodeLevels.set(q.id, q.level));

        let head = 0;
        while (head < queue.length) {
            const { id, level } = queue[head++];
            const children = edgesToProcess.filter(e => e.source === id).map(e => e.target);

            children.forEach(childId => {
                const currentLevel = nodeLevels.get(childId) || 0;
                const nextLevel = Math.max(currentLevel, level + 1);
                nodeLevels.set(childId, nextLevel);

                const count = (visitCount.get(childId) || 0) + 1;
                visitCount.set(childId, count);

                // Avoid infinite loops in cycles
                if (count < 10) {
                    queue.push({ id: childId, level: nextLevel });
                }
            });
        }

        const levelsGroup: { [key: number]: string[] } = {};
        nodesToProcess.forEach(n => {
            const l = nodeLevels.get(n.id) ?? 0;
            if (!levelsGroup[l]) levelsGroup[l] = [];
            levelsGroup[l].push(n.id);
        });

        const newNodes = nodesToProcess.map(n => {
            const l = nodeLevels.get(n.id) ?? 0;
            const levelNodes = levelsGroup[l] || [];
            const idxInLevel = levelNodes.indexOf(n.id);
            const totalInLevel = levelNodes.length;
            const yOffset = (idxInLevel - (totalInLevel - 1) / 2) * siblingGapY;

            return {
                ...n,
                position: {
                    x: 100 + l * levelGapX,
                    y: 300 + yOffset
                }
            };
        });

        setNodes(newNodes);
        showToast("Auto-Aligned", "Workflow nodes have been perfectly organized.", "info");
        fitView({ padding: 0.2, duration: 800 });
    }, [nodes, edges, setNodes, fitView, showToast]);

    // Keyboard Shortcuts
    useEffect(() => {
        const handleKeyPress = (e: KeyboardEvent) => {
            // Ctrl + S (Save)
            if ((e.ctrlKey || e.metaKey) && e.key === 's') {
                e.preventDefault();
                const saveBtn = document.getElementById('wf-save-btn');
                if (saveBtn) saveBtn.click();
            }

            // Delete / Backspace (Delete Node)
            if ((e.key === 'Delete' || e.key === 'Backspace') && selectedNodeId) {
                // Don't delete if typing in inputs
                const activeTag = document.activeElement?.tagName.toLowerCase();
                if (activeTag === 'input' || activeTag === 'textarea' || activeTag === 'select') return;

                setNodes((nds) => nds.filter((node) => node.id !== selectedNodeId));
                setEdges((eds) => eds.filter((edge) => edge.source !== selectedNodeId && edge.target !== selectedNodeId));
                setSelectedNodeId(null);
            }

            // Ctrl + D (Duplicate Node)
            if ((e.ctrlKey || e.metaKey) && e.key === 'd' && selectedNodeId) {
                e.preventDefault();
                const nodeToCopy = nodes.filter(n => n.id === selectedNodeId)[0];
                if (nodeToCopy && nodeToCopy.data.type !== 'Start' && nodeToCopy.data.type !== 'Main Task') {
                    const newId = `node_${nodeIdCounter}_${Date.now()}`;
                    const newNode = {
                        ...nodeToCopy,
                        id: newId,
                        position: { x: nodeToCopy.position.x + 40, y: nodeToCopy.position.y + 40 },
                        data: {
                            ...nodeToCopy.data,
                            onPlusClick: (e: React.MouseEvent) => handlePlusClick(e, newId)
                        },
                        selected: true
                    };
                    setNodes((nds) => nds.map(n => ({ ...n, selected: false })).concat(newNode));
                    setNodeIdCounter(prev => prev + 1);
                    setSelectedNodeId(newId);
                }
            }
        };

        window.addEventListener('keydown', handleKeyPress);
        return () => window.removeEventListener('keydown', handleKeyPress);
    }, [selectedNodeId, nodes, nodeIdCounter, setNodes, setEdges]);

    const handlePlusClick = useCallback((e: React.MouseEvent, nodeId: string) => {
        e.stopPropagation();
        e.preventDefault();
        setWhatsNextSourceNode(nodeId);
        setIsWhatsNextOpen(true);
        setSelectedNodeId(null); // Close properties panel to avoid confusion
    }, []);

    // -- Data Loading --
    const loadGlobalOptions = useCallback(async () => {
        try {
            const email = await taskService.getCurrentUserEmail();
            setCurrentUserEmail(email);

            // Fetch all options in parallel for efficiency
            const [users, categories, depts, years, months] = await Promise.all([
                taskService.getSiteUsers().catch(e => { console.error("User fetch failed", e); return []; }),
                taskService.getChoiceFieldOptions('Task Tracking System User', 'Category').catch(e => { console.error("Category fetch failed", e); return []; }),
                taskService.getChoiceFieldOptions('Task Tracking System', 'Departments').catch(e => { console.error("Dept fetch failed", e); return []; }),
                taskService.getChoiceFieldOptions('Task Tracking System', 'SMTYear').catch(e => { console.error("Year fetch failed", e); return []; }),
                taskService.getChoiceFieldOptions('Task Tracking System', 'SMTMonth').catch(e => { console.error("Month fetch failed", e); return []; })
            ]);

            const mappedUsers: IComboBoxOption[] = (users || []).map(u => ({
                key: (u.Email || u.EMail || u.LoginName || u.Id?.toString() || '').toLowerCase(),
                text: u.Title || u.Name || 'Unknown User'
            })).filter(u => u.key);

            console.log(`[WorkflowDesigner] Fetched ${mappedUsers.length} users from SharePoint`);

            // 1. HARDCODED SAFETY USER FOR DEBUGGING
            mappedUsers.push({ key: 'safety_user@example.com', text: '!!! Safety Debug User !!!' });

            // 2. Ensure current user from TaskService is present
            if (email && !mappedUsers.some(u => u.key === email.toLowerCase())) {
                console.log(`[WorkflowDesigner] Adding current user from TaskService: ${email}`);
                mappedUsers.push({ key: email.toLowerCase(), text: userDisplayName || email });
            }

            // 3. Ensure current user from Props is present (DOUBLE SAFETY)
            if (userEmail && !mappedUsers.some(u => u.key === userEmail.toLowerCase())) {
                console.log(`[WorkflowDesigner] Adding current user from PROPS: ${userEmail}`);
                mappedUsers.push({ key: userEmail.toLowerCase(), text: userDisplayName || userEmail });
            }

            console.log(`[WorkflowDesigner] FINAL userOptions count: ${mappedUsers.length}`);
            console.log("[WorkflowDesigner] User list sample:", mappedUsers.slice(0, 3));
            setUserOptions(mappedUsers);

            if (categories && categories.length > 0) setCategoryOptions(categories.map(c => ({ key: c, text: c })));
            if (depts && depts.length > 0) setDepartmentOptions(depts.map(d => ({ key: d, text: d })));

            // Year Options with Fallback
            if (years && years.length > 0) {
                setYearOptions(years.map(y => ({ key: y, text: y })));
            } else {
                const currentYear = new Date().getFullYear();
                const fallbackYears = [currentYear - 1, currentYear, currentYear + 1, currentYear + 2];
                setYearOptions(fallbackYears.map(y => ({ key: y.toString(), text: y.toString() })));
            }

            // Month Options with Fallback
            if (months && months.length > 0) {
                setMonthOptions(months.map(m => ({ key: m, text: m })));
            } else {
                const fallbackMonths = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                setMonthOptions(fallbackMonths.map(m => ({ key: m, text: m })));
            }

        } catch (e) {
            console.error("[WorkflowDesigner] Error loading global options", e);
        }
    }, [userDisplayName, userEmail]);

    const refreshWorkflowData = useCallback(async (isRefresh: boolean = false) => {
        if (!isRefresh && hasInitialLoaded.current && lastMainTaskIdRef.current === mainTaskId) return;
        hasInitialLoaded.current = true;
        lastMainTaskIdRef.current = mainTaskId;

        try {
            setIsSyncing(true);
            if (mainTaskId) {
                const mt = await taskService.getMainTaskById(mainTaskId);
                if (mt) setMainTaskTitle(mt.Title);
            }

            let workflow;
            if (mainTaskId) {
                workflow = await taskService.getWorkflowByTitle(`TASK_WF_${mainTaskId}`);
            } else {
                workflow = await taskService.getActiveWorkflow();
            }

            let currentNodes: Node[] = [];
            let currentEdges: Edge[] = [];

            if (workflow?.WorkflowJson) {
                try {
                    const parsed = JSON.parse(workflow.WorkflowJson);
                    if (parsed.nodes) {
                        currentNodes = parsed.nodes;
                        // If we are in generic mode (no mainTaskId) but load a workflow WITH a main task node, lock it in
                        if (!mainTaskId) {
                            const mtNode = currentNodes.find((n: any) => n.data.type === 'Main Task');
                            if (mtNode && mtNode.data.linkedTaskId) {
                                console.log(`[WorkflowDesigner] Auto-linking to Main Task: ${mtNode.data.linkedTaskId}`);
                                setMainTaskId(mtNode.data.linkedTaskId);
                            }
                        }
                    }
                    if (parsed.edges) {
                        currentEdges = parsed.edges.map((edge: Edge) => ({
                            ...edge,
                            type: 'smoothstep',
                            style: { stroke: '#94a3b8', strokeWidth: 3 },
                            markerEnd: { type: MarkerType.ArrowClosed, color: '#94a3b8' },
                        }));
                    }
                } catch (e) {
                    console.error("Error parsing workflow JSON", e);
                }
            }

            // Sync with SharePoint tasks
            let latestTasks: any[] = [];
            let mt_obj: any = null;

            if (mainTaskId) {
                [latestTasks, mt_obj] = await Promise.all([
                    taskService.getSubTasksForMainTask(mainTaskId).catch(e => { console.error("Subtasks fetch failed", e); return []; }),
                    taskService.getMainTaskById(mainTaskId).catch(e => { console.error("Main task fetch failed", e); return null; })
                ]);
            }

            const taskToNodeIdMap = new Map<number, string>();
            currentNodes.forEach(n => {
                if (n.data.linkedTaskId) taskToNodeIdMap.set(n.data.linkedTaskId, n.id);
            });

            if (mt_obj && !currentNodes.some(n => n.data.type === 'Main Task')) {
                currentNodes.push({
                    id: 'node-main',
                    type: 'customNode',
                    position: { x: 100, y: 300 },
                    data: {
                        label: mt_obj.Title,
                        type: 'Main Task',
                        linkedTaskId: mt_obj.Id,
                        status: mt_obj.Status,
                        description: mt_obj.Task_x0020_Description,
                        assignee: mt_obj.TaskAssignedTo?.EMail?.toLowerCase() || '',
                        year: mt_obj.SMTYear,
                        month: mt_obj.SMTMonth,
                        department: mt_obj.Departments,
                        project: mt_obj.Project,
                        dueDate: mt_obj.TaskDueDate
                    }
                });
                taskToNodeIdMap.set(mt_obj.Id, 'node-main');
            }

            const maxId = currentNodes.reduce((acc, n) => {
                const num = parseInt(n.id.replace('node-', '').replace('node_', ''));
                return !isNaN(num) && num > acc ? num : acc;
            }, 0);

            // 1. Sync Nodes: Add missing subtasks as nodes
            let countImported = 0;
            latestTasks.forEach(task => {
                if (!taskToNodeIdMap.has(task.Id)) {
                    const node_id = `node-${maxId + 1 + countImported}`;
                    currentNodes.push({
                        id: node_id,
                        type: 'customNode',
                        position: { x: 300 + (countImported * 50), y: 300 + (countImported * 50) },
                        data: {
                            label: task.Task_Title,
                            type: 'Task',
                            linkedTaskId: task.Id,
                            status: task.TaskStatus,
                            description: task.Task_Description,
                            assignee: task.TaskAssignedTo?.EMail?.toLowerCase() || ''
                        }
                    });
                    taskToNodeIdMap.set(task.Id, node_id);
                    countImported++;
                } else {
                    // Update status for existing nodes
                    const node_id = taskToNodeIdMap.get(task.Id);
                    const nodeIndex = currentNodes.findIndex(n => n.id === node_id);
                    if (nodeIndex !== -1) {
                        currentNodes[nodeIndex] = {
                            ...currentNodes[nodeIndex],
                            data: {
                                ...currentNodes[nodeIndex].data,
                                status: task.TaskStatus,
                                assignee: task.TaskAssignedTo?.EMail?.toLowerCase() || currentNodes[nodeIndex].data.assignee
                            }
                        };
                    }
                }
            });

            // 2. Sync Edges: Link children to parents based on SP data
            latestTasks.forEach(task => {
                const node_id = taskToNodeIdMap.get(task.Id);
                let parent_node_id = '';

                if (task.ParentSubtaskId && taskToNodeIdMap.has(task.ParentSubtaskId)) {
                    parent_node_id = taskToNodeIdMap.get(task.ParentSubtaskId)!;
                } else if (mt_obj) {
                    parent_node_id = 'node-main';
                }

                if (node_id && parent_node_id) {
                    // Avoid duplicate edges
                    if (!currentEdges.some(e => e.source === parent_node_id && e.target === node_id)) {
                        currentEdges.push({
                            id: `e-${parent_node_id}-${node_id}`,
                            source: parent_node_id,
                            target: node_id,
                            type: 'smoothstep',
                            animated: true,
                            style: { stroke: '#94a3b8', strokeWidth: 3 },
                            markerEnd: { type: MarkerType.ArrowClosed, color: '#94a3b8' },
                        });
                    }
                }
            });

            // --- 3. Feature Upgrades: Progress & Dependencies ---

            // Calculate Main Task Progress
            const subTaskNodes = currentNodes.filter(n => n.data.type === 'Task' && n.data.linkedTaskId);
            const completedCount = subTaskNodes.filter(n => n.data.status === 'Completed').length;
            const progress = subTaskNodes.length > 0 ? (completedCount / subTaskNodes.length) * 100 : 0;

            // Deep-link highlighting logic
            const urlParams = new URLSearchParams(window.location.search);
            const highlightChildId = urlParams.get('ChildTaskID');
            const highlightParentId = urlParams.get('ParentTaskID');
            const targetTaskId = highlightChildId ? parseInt(highlightChildId) : (highlightParentId ? parseInt(highlightParentId) : null);

            // Updated nodes with progress, blocked status, and highlighting
            const finalNodes = currentNodes.map(node => {
                let isBlocked = false;

                // For subtasks, check if any parent (incoming edge) is NOT completed
                if (node.data.type === 'Task' && node.data.linkedTaskId) {
                    const incomingEdges = currentEdges.filter(e => e.target === node.id);
                    const parentsIncomplete = incomingEdges.some(edge => {
                        const parentNode = currentNodes.find(n => n.id === edge.source);
                        return parentNode && parentNode.data.type === 'Task' && parentNode.data.status !== 'Completed';
                    });
                    isBlocked = parentsIncomplete;
                }

                // Highlighting logic
                const isHighlighted = targetTaskId && node.data.linkedTaskId === targetTaskId;

                return {
                    ...node,
                    data: {
                        ...node.data,
                        progress: node.data.type === 'Main Task' ? progress : undefined,
                        isBlocked: isBlocked,
                        isHighlighted: isHighlighted,
                        onPlusClick: (e: React.MouseEvent) => handlePlusClick(e, node.id),
                        onDrop: (files: File[]) => handleNodeDrop(node, files)
                    }
                };
            });


            setNodes(finalNodes);
            setEdges(currentEdges);
            setNodeIdCounter(maxId + countImported + 1);

            if (countImported > 0) {
                setTimeout(() => tidyUp(finalNodes, currentEdges), 100);
            }

        } catch (e) {
            console.error("[WorkflowDesigner] Error loading workflow data", e);
        } finally {
            setIsSyncing(false);
        }
    }, [mainTaskId]);

    useEffect(() => {
        loadGlobalOptions();
    }, [loadGlobalOptions]);

    useEffect(() => {
        refreshWorkflowData();
    }, [mainTaskId, refreshWorkflowData]);

    const onNodesChange = useCallback(
        (changes: NodeChange[]) => setNodes((nds) => applyNodeChanges(changes, nds)),
        []
    );

    const onEdgesChange = useCallback(
        (changes: EdgeChange[]) => setEdges((eds) => applyEdgeChanges(changes, eds)),
        []
    );

    // Context Menu Handlers
    const onNodeContextMenu = useCallback(
        (event: React.MouseEvent, node: Node) => {
            event.preventDefault();
            setMenu({
                x: event.clientX,
                y: event.clientY,
                nodeId: node.id,
            });
        },
        [setMenu]
    );

    const onPaneClick = useCallback(() => {
        setMenu(null);
        setIsWhatsNextOpen(false);
        setSelectedNodeId(null);
    }, [setMenu, setIsWhatsNextOpen, setSelectedNodeId]);

    const handleDuplicateNode = useCallback(() => {
        if (!menu?.nodeId) return;
        const nodeToCopy = nodes.filter(n => n.id === menu.nodeId)[0];
        if (!nodeToCopy) return;

        const newNodeId = `node_${Date.now()}`;
        const newNode: Node = {
            ...nodeToCopy,
            id: newNodeId,
            position: {
                x: nodeToCopy.position.x + 100,
                y: nodeToCopy.position.y + 100,
            },
            data: {
                ...nodeToCopy.data,
                label: `${nodeToCopy.data.label} (Copy)`,
                linkedTaskId: undefined,
                status: 'Planned',
                onPlusClick: (e: React.MouseEvent) => handlePlusClick(e, newNodeId)
            },
            selected: true
        };

        setNodes((nds) => {
            const deselected = nds.map(n => ({ ...n, selected: false })) as Node[];
            return deselected.concat([newNode]);
        });
        setMenu(null);
    }, [menu, nodes, setNodes, handlePlusClick]);

    const handleDeleteSelected = useCallback(() => {
        const selectedIds = nodes.filter(n => !!n.selected).map(n => n.id);
        if (selectedIds.length === 0 && menu?.nodeId) {
            selectedIds.push(menu.nodeId);
        }

        setNodes((nds) => nds.filter((node) => selectedIds.indexOf(node.id) === -1));
        setEdges((eds) => eds.filter((edge) => selectedIds.indexOf(edge.source) === -1 && selectedIds.indexOf(edge.target) === -1));
        setMenu(null);
    }, [menu, nodes, setNodes, setEdges]);

    const handleCreateMainTaskFromDialog = async () => {
        if (!newMainTaskTitle || !newMainTaskDesc || !newMainTaskYear || !newMainTaskMonth || !newMainTaskAssignee || !newMainTaskDueDate) {
            alert("Please fill in all required fields.");
            return;
        }

        let finalAssigneeId: number | undefined;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let matched: any = null;

        if (typeof newMainTaskAssignee === 'string') {
            const matchedUsers = userOptions.filter(u => u.key === newMainTaskAssignee || u.text === newMainTaskAssignee || (u.data && u.data.email === newMainTaskAssignee));
            matched = matchedUsers.length > 0 ? matchedUsers[0] : null;

            if (matched) {
                try {
                    finalAssigneeId = await taskService.ensureUser(matched.key as string);
                } catch (e) {
                    alert("Could not resolve assigned user.");
                    return;
                }
            } else {
                try {
                    finalAssigneeId = await taskService.ensureUser(newMainTaskAssignee);
                } catch (e) {
                    alert("Invalid user selected.");
                    return;
                }
            }
        }

        setIsCreateMainTaskOpen(false);
        setIsSyncing(true);

        try {
            const newMainTaskId = await taskService.createMainTask({
                Title: newMainTaskTitle,
                Task_x0020_Description: newMainTaskDesc,
                SMTYear: newMainTaskYear,
                SMTMonth: newMainTaskMonth,
                TaskAssignedToId: finalAssigneeId,
                Departments: newMainTaskDept,
                Project: newMainTaskProject,
                TaskDueDate: newMainTaskDueDate.toISOString(),
                Status: 'Not Started'
            } as any, newMainTaskFiles);

            // Optimistic Update
            const newNodeId = `node-main`;
            const newNode: Node = {
                id: newNodeId,
                type: 'customNode',
                position: { x: 300, y: 300 },
                data: {
                    label: newMainTaskTitle,
                    type: 'Main Task',
                    description: newMainTaskDesc,
                    assignee: matched?.key || newMainTaskAssignee, // Use KEY (Email) not Text (Title)
                    linkedTaskId: newMainTaskId,
                    status: 'Not Started',
                    year: newMainTaskYear,
                    month: newMainTaskMonth,
                    department: newMainTaskDept,
                    project: newMainTaskProject,
                    dueDate: newMainTaskDueDate.toISOString(),
                    onPlusClick: (e: React.MouseEvent) => handlePlusClick(e, newNodeId)
                }
            };

            setNodes(nds => [...nds, newNode]);
            setNodeIdCounter(prev => prev + 1);
            setMainTaskId(newMainTaskId); // IMPORTANT: Lock in this Main Task ID for subsequent tasks

            refreshWorkflowData();
            setTimeout(() => fitView({ padding: 50, duration: 800 }), 100);
            showToast("Main Task Created", "Workflow initialized.", "success");

        } catch (e) {
            console.error("Error creating main task", e);
            alert("Error creating main task.");
        } finally {
            setIsSyncing(false);
        }
    };

    const handleCreateTaskFromDialog = async () => {
        if (!newTaskAssignee) {
            alert("Please select an assignee.");
            return;
        }

        let finalAssigneeId: number | undefined;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let matched: any = null;

        if (typeof newTaskAssignee === 'string') {
            const matchedUsers = userOptions.filter(u => u.key === newTaskAssignee || u.text === newTaskAssignee || (u.data && u.data.email === newTaskAssignee));
            matched = matchedUsers.length > 0 ? matchedUsers[0] : null;

            if (matched) {
                try {
                    console.log(`[CreateTask] Resolving user: ${matched.key}`);
                    finalAssigneeId = await taskService.ensureUser(matched.key as string);
                } catch (e) {
                    console.error("Could not resolve user", e);
                    alert("Could not resolve assigned user. Please try again.");
                    return;
                }
            } else {
                try {
                    finalAssigneeId = await taskService.ensureUser(newTaskAssignee as string);
                } catch (e) {
                    alert("Invalid user selected.");
                    return;
                }
            }
        } else {
            finalAssigneeId = newTaskAssignee as number;
        }

        setIsCreateTaskOpen(false);
        setIsSyncing(true);
        try {
            console.log(`[WorkflowDesigner] Creating subtask linked to Main Task ID: ${mainTaskId}`);
            const newSubTaskId = await taskService.createSubTask({
                Admin_Job_ID: mainTaskId,
                Title: newTaskTitle,
                Task_Title: newTaskTitle,
                Task_Description: newTaskDesc,
                Category: 'Development', // Or select from dropdown
                TaskDueDate: newTaskDueDate ? newTaskDueDate.toISOString() : undefined,
                TaskStatus: 'Not Started',
                TaskAssignedToId: finalAssigneeId as number,
                ParentSubtaskId: parentSubtaskId // Link to source task!
            } as any);

            // --- Optimistic Update ---
            const newNodeId = `node-${nodeIdCounter}`;
            // Determine parent node ID for the edge
            let sourceNodeId = '';
            if (parentSubtaskId) {
                const parentNode = nodes.find(n => n.data.linkedTaskId === parentSubtaskId);
                if (parentNode) sourceNodeId = parentNode.id;
            } else {
                // Main task or unknown. Try finding Main Task node
                const mainNode = nodes.find(n => n.data.type === 'Main Task');
                if (mainNode) sourceNodeId = mainNode.id;
                else if (whatsNextSourceNode) sourceNodeId = whatsNextSourceNode;
            }

            if (sourceNodeId) {
                const sourceNode = nodes.find(n => n.id === sourceNodeId);
                const newNode: Node = {
                    id: newNodeId,
                    type: 'customNode',
                    position: {
                        x: sourceNode ? sourceNode.position.x + 350 : 300,
                        y: sourceNode ? sourceNode.position.y : 300
                    },
                    data: {
                        label: newTaskTitle,
                        type: 'Task',
                        description: newTaskDesc,
                        assignee: matched?.key || (typeof newTaskAssignee === 'string' ? newTaskAssignee : matched?.text || ''),
                        linkedTaskId: newSubTaskId,
                        status: 'Not Started',
                        onPlusClick: (e: React.MouseEvent) => handlePlusClick(e, newNodeId)
                    }
                };

                const newEdge: Edge = {
                    id: `e-${sourceNodeId}-${newNodeId}`,
                    source: sourceNodeId,
                    target: newNodeId,
                    type: 'smoothstep',
                    animated: true,
                    style: { stroke: '#94a3b8', strokeWidth: 3 },
                    markerEnd: { type: MarkerType.ArrowClosed, color: '#94a3b8' },
                };

                setNodes(nds => [...nds, newNode]);
                setEdges(eds => [...eds, newEdge]);
                setNodeIdCounter(prev => prev + 1);
            }

            // Still do a background refresh to ensure consistency
            refreshWorkflowData();
            setTimeout(() => fitView({ padding: 50, duration: 800 }), 100);
            showToast("Subtask Created", `Task "${newTaskTitle}" added to SharePoint.`, "success");

        } catch (e) {
            console.error("Error creating task", e);
            alert("Error creating task: " + (e as any).message);
        } finally {
            setIsSyncing(false);
        }
    };

    const handleQuickAdd = useCallback((template: typeof nodeTemplates[0]) => {
        if (!whatsNextSourceNode || addingNodeType) return;

        setAddingNodeType(template.type);

        // n8n Style artificial delay for "loading" feel
        setTimeout(() => {
            const sourceNodes = nodes.filter((n: Node) => n.id === whatsNextSourceNode);
            const sourceNode = sourceNodes.length > 0 ? sourceNodes[0] : null;
            if (!sourceNode) {
                setAddingNodeType(null);
                return;
            }

            // *** INTERCEPT TASK CREATION ***
            if (template.type === 'Task') {
                setAddingNodeType(null); // Stop spinner
                setIsWhatsNextOpen(false); // Close panel

                // Pre-fill dialog
                setNewTaskTitle('');
                setNewTaskDesc('');
                setNewTaskAssignee(undefined);
                setNewTaskDueDate(new Date());

                // Find Linked Task ID of Source Node to set as Parent
                const sourceTaskId = sourceNode.data.linkedTaskId;
                // Fix: If source is Main Task, ParentSubtaskId should be undefined (direct child of main task)
                if (sourceNode.data.type === 'Main Task') {
                    setParentSubtaskId(undefined);
                } else {
                    setParentSubtaskId(sourceTaskId);
                }

                setIsCreateTaskOpen(true);
                return;
            }

            const newNodeId = `node-${nodeIdCounter}`;
            const newNode: Node = {
                id: newNodeId,
                type: 'customNode',
                position: {
                    x: sourceNode.position.x + 350,
                    y: sourceNode.position.y
                },
                data: {
                    label: `${template.type} ${nodeIdCounter}`,
                    type: template.type,
                    description: '',
                    assignee: '',
                    onPlusClick: (e: React.MouseEvent) => handlePlusClick(e, newNodeId)
                }
            };

            const newEdge: Edge = {
                id: `edge-${sourceNode.id}-${newNodeId}`,
                source: sourceNode.id,
                target: newNodeId,
                type: 'smoothstep',
                animated: true,
                style: { stroke: '#3b82f6', strokeWidth: 4 },
                markerEnd: { type: MarkerType.ArrowClosed, color: '#3b82f6' },
                data: { isManual: true }
            };

            setNodes(nds => [...nds, newNode]);
            setEdges(eds => [...eds, newEdge]);
            setNodeIdCounter(prev => prev + 1);
            setIsWhatsNextOpen(false);
            setLibrarySearchQuery('');
            setAddingNodeType(null);
        }, 600);
    }, [whatsNextSourceNode, nodes, nodeIdCounter, addingNodeType, setNodes, setEdges, setNodeIdCounter, handlePlusClick]);

    const handleTidyUp = useCallback(async () => {
        setIsSyncing(true);
        try {
            const latestTasks = await taskService.getSubTasksByMainTaskIds([mainTaskId || 0]);
            const mt = mainTaskId ? await taskService.getMainTaskById(mainTaskId) : null;
            tidyUp();
            setTimeout(() => fitView({ padding: 50, duration: 800 }), 100);
        } finally {
            setIsSyncing(false);
        }
        setMenu(null);
    }, [mainTaskId, nodes, edges, tidyUp, fitView]);

    const onConnect = useCallback(
        (params: Connection) => setEdges((eds) => addEdge({
            ...params,
            type: 'smoothstep',
            animated: true,
            style: {
                stroke: '#3b82f6',
                strokeWidth: 4,
                strokeDasharray: '0',
                filter: 'drop-shadow(0px 0px 5px rgba(59, 130, 246, 0.5))'
            }, // Pro blue with glow
            markerEnd: {
                type: MarkerType.ArrowClosed,
                color: '#3b82f6',
            },
            data: { isManual: true }
        }, eds)),
        [setEdges]
    );

    const onNodeClick = (event: React.MouseEvent, node: Node) => {
        event.preventDefault();
        setSelectedNodeId(node.id);
    };

    const onPaneContextMenu = useCallback((event: React.MouseEvent) => {
        event.preventDefault();
        setMenu({ x: event.clientX, y: event.clientY });
    }, [setMenu]);

    const updateSelectedNode = (field: string, value: any) => {
        setNodes((nds) => nds.map((node) => {
            if (node.id === selectedNodeId) {
                return {
                    ...node,
                    data: {
                        ...node.data,
                        [field]: value
                    }
                };
            }
            return node;
        }));
    };

    const handleSave = async () => {
        if (nodes.length === 0) {
            showToast("Empty Canvas", "Add nodes before saving.", "warning");
            return;
        }

        // 1. Check for Admin Job ID (Main Task ID)
        if (!mainTaskId) {
            showToast("Sync Error", "No Main Task ID (Admin Job ID) detected. Please open the designer from a valid task.", "error");
            return;
        }

        // 2. Check if all Task nodes are linked to SharePoint
        const unlinkedNodes = nodes.filter(n => (n.data.type === 'Task' || n.data.type === 'Main Task') && !n.data.linkedTaskId);
        if (unlinkedNodes.length > 0) {
            showToast("Action Required", "Please create a SharePoint task for all nodes before saving the workflow.", "error");
            // Focus on the first unlinked node
            setSelectedNodeId(unlinkedNodes[0].id);
            return;
        }

        setIsSaving(true);
        try {
            const nodesToSave = nodes.map(n => ({
                id: n.id,
                type: n.type,
                position: n.position,
                data: n.data
            }));

            const edgesToSave = edges.map(e => ({
                id: e.id,
                source: e.source,
                target: e.target,
                label: e.label,
                type: e.type,
                style: e.style,
                markerEnd: e.markerEnd
            }));

            await taskService.updateOrCreateWorkflow(`TASK_WF_${mainTaskId}`, nodesToSave, edgesToSave);

            showToast("Saved Successfully", "Your workflow design has been stored.", "success");
        } catch (e) {
            console.error("Save Error:", e);
            showToast("Save Failed", "Could not save the workflow.", "error");
        } finally {
            setIsSaving(false);
        }
    };

    const handleClear = () => {
        if (confirm("Are you sure you want to clear the canvas?")) {
            setNodes([]);
            setEdges([]);
            setNodeIdCounter(1);
            setSelectedNodeId(null);
        }
    };

    const onDragStart = (event: React.DragEvent, nodeType: typeof nodeTemplates[0]): void => {
        event.dataTransfer.setData('application/reactflow', JSON.stringify(nodeType));
        event.dataTransfer.effectAllowed = 'move';
    };


    const onDrop = (event: React.DragEvent): void => {
        event.preventDefault();
        const reactFlowBounds = event.currentTarget.getBoundingClientRect();
        const typeData = JSON.parse(event.dataTransfer.getData('application/reactflow'));

        // Project coordinate system for accurate drops
        const position = project({
            x: event.clientX - reactFlowBounds.left,
            y: event.clientY - reactFlowBounds.top
        });

        const newNodeId = `node-${nodeIdCounter}`;
        const newNode: Node = {
            id: newNodeId,
            type: 'customNode',
            position,
            data: {
                label: `${typeData.type} ${nodeIdCounter}`,
                type: typeData.type,
                description: '',
                assignee: '',
                onPlusClick: (e: React.MouseEvent) => handlePlusClick(e, newNodeId)
            }
        };

        setNodes((nds) => [...nds, newNode]);
        setNodeIdCounter((prev) => prev + 1);
    };

    const onDragOver = (event: React.DragEvent): void => {
        event.preventDefault();
        event.dataTransfer.dropEffect = 'move';
    };

    const getBestParentNode = useCallback((targetNodeId: string | null) => {
        if (!targetNodeId) return null;
        const incomingEdges = edges.filter(e => e.target === targetNodeId);
        if (incomingEdges.length === 0) return null;

        console.log(`[Hierarchy] Detecting parent for ${targetNodeId}. Incoming edges:`, incomingEdges.length);

        // Priority 1: Manual Blue Links (User Intent)
        // Check both data flag and style as fallback
        for (const edge of incomingEdges) {
            const isManual = (edge as any).data?.isManual || (edge.style as any)?.stroke === '#3b82f6';
            if (isManual) {
                const srcNode = nodes.filter(n => n.id === edge.source)[0];
                if (srcNode) {
                    console.log(`[Hierarchy] Found manual parent: ${srcNode.data.label} (ID: ${srcNode.id})`);
                    return srcNode;
                }
            }
        }

        // Priority 2: Linked Subtasks (Solid Gray)
        for (const edge of incomingEdges) {
            const srcNode = nodes.filter(n => n.id === edge.source)[0];
            if (srcNode && srcNode.data.linkedTaskId && srcNode.data.type !== 'Main Task') {
                console.log(`[Hierarchy] Found linked subtask parent: ${srcNode.data.label}`);
                return srcNode;
            }
        }

        // Priority 3: Main Task
        for (const edge of incomingEdges) {
            const srcNode = nodes.filter(n => n.id === edge.source)[0];
            if (srcNode && srcNode.data.type === 'Main Task') {
                console.log(`[Hierarchy] Found Main Task parent`);
                return srcNode;
            }
        }

        // Fallback: First incoming node
        const fallback = nodes.filter(n => n.id === incomingEdges[0].source)[0];
        if (fallback) console.log(`[Hierarchy] Fallback to: ${fallback.data.label}`);
        return fallback || null;
    }, [edges, nodes]);

    const handleCreateActualTask = async (node: Node, overrides?: any) => {
        console.log('[Debug] handleCreateActualTask. MainTaskId:', mainTaskId);
        console.log('[Debug] node:', node);
        console.log('[Debug] overrides:', overrides);

        let assignee = overrides?.assignee || node.data.assignee;
        if (!assignee) {
            showToast("Required Field", "Assignee is mandatory. Please select a user.", "warning");
            return;
        }

        if (node.data.type === 'Main Task') {
            if (!node.data.year || !node.data.month) {
                showToast("Required Fields", "Year and Month are mandatory for Main Tasks.", "warning");
                return;
            }
        }

        // Resolve name to email if needed
        const matchedUsers = userOptions.filter((u: IComboBoxOption) => u.text === assignee || u.key === assignee);
        if (matchedUsers.length > 0) {
            assignee = matchedUsers[0].key as string;
        }

        setIsSyncing(true);
        try {
            const userId = await taskService.ensureUser(assignee);
            let taskId: number;

            // --- Smart Hierarchy Detection ---
            let parentTaskId: number | undefined = undefined;
            const bestSourceNode = getBestParentNode(node.id);

            if (bestSourceNode) {
                if (!bestSourceNode.data.linkedTaskId && bestSourceNode.data.type !== 'Main Task') {
                    showToast("Parent Required", `Please link the parent node "${bestSourceNode.data.label}" to SharePoint first.`, "warning");
                    setSelectedNodeId(bestSourceNode.id);
                    return;
                }

                if (bestSourceNode.data.linkedTaskId && bestSourceNode.data.type !== 'Main Task') {
                    parentTaskId = bestSourceNode.data.linkedTaskId;
                }
            }

            if (node.data.type === 'Main Task') {
                taskId = await taskService.createMainTask({
                    Title: overrides?.title || node.data.label,
                    Task_x0020_Description: overrides?.description || node.data.description,
                    TaskAssignedToId: userId as any,
                    Status: 'Not Started',
                    Project: node.data.project || 'Workflow Task',
                    SMTYear: node.data.year || new Date().getFullYear().toString(),
                    SMTMonth: node.data.month || new Intl.DateTimeFormat('en-US', { month: 'long' }).format(new Date()),
                    Departments: node.data.department,
                    TaskDueDate: overrides?.dueDate ? new Date(overrides.dueDate).toISOString() : (node.data.dueDate ? new Date(node.data.dueDate).toISOString() : new Date().toISOString()),
                    UserRemarks: `[WF_NODE:${node.id}]`
                } as any);
            } else if (mainTaskId) {
                taskId = await taskService.createSubTask({
                    Task_Title: overrides?.title || node.data.label,
                    Task_Description: overrides?.description || node.data.description,
                    TaskAssignedToId: userId as any,
                    TaskStatus: 'Not Started',
                    Admin_Job_ID: mainTaskId,
                    TaskDueDate: overrides?.dueDate ? new Date(overrides.dueDate).toISOString() : new Date().toISOString(),
                    Category: node.data.category || 'Workflow',
                    User_Remarks: `[WF_NODE:${node.id}]`,
                    ParentSubtaskId: parentTaskId
                });
            } else {
                taskId = await taskService.createMainTask({
                    Title: overrides?.title || node.data.label,
                    Task_x0020_Description: overrides?.description || node.data.description,
                    TaskAssignedToId: userId as any,
                    Status: 'Not Started',
                    Project: node.data.project || 'Workflow Task',
                    SMTYear: node.data.year || new Date().getFullYear().toString(),
                    SMTMonth: node.data.month || new Intl.DateTimeFormat('en-US', { month: 'long' }).format(new Date()),
                    Departments: node.data.department,
                    TaskDueDate: overrides?.dueDate ? new Date(overrides.dueDate).toISOString() : (node.data.dueDate ? new Date(node.data.dueDate).toISOString() : new Date().toISOString()),
                    UserRemarks: `[WF_NODE:${node.id}]`
                } as any);
            }

            if (!taskId) throw new Error("Task creation returned no ID");

            updateSelectedNode('linkedTaskId', taskId);
            updateSelectedNode('status', 'Not Started');

            if (overrides) {
                if (overrides.title) updateSelectedNode('label', overrides.title);
                if (overrides.description) updateSelectedNode('description', overrides.description);
                if (overrides.assignee) updateSelectedNode('assignee', overrides.assignee);
                if (overrides.dueDate) updateSelectedNode('dueDate', overrides.dueDate);
            }

            const successMsg = parentTaskId
                ? `Sub-task created and linked to parent Task ID: #${parentTaskId}`
                : `Main task created successfully with ID: #${taskId}`;

            showToast("Linked Successfully", successMsg, "success");
        } catch (e) {
            console.error("Task Creation Error:", e);
            showToast("Creation Failed", e.message || "Failed to create SharePoint task.", "error");
        } finally {
            setIsSyncing(false);
        }
    };

    const handleSyncStatus = async (node: Node) => {
        if (!node.data.linkedTaskId) return;
        setIsSyncing(true);
        try {
            let taskStatus = '';
            if (mainTaskId) {
                const subTask = await taskService.getSubTaskById(node.data.linkedTaskId);
                if (subTask) taskStatus = subTask.TaskStatus;
            } else {
                const task = await taskService.getMainTaskById(node.data.linkedTaskId);
                if (task) taskStatus = task.Status;
            }

            if (taskStatus) {
                updateSelectedNode('status', taskStatus);
            }
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
        }
    };

    const handleNodeDrop = useCallback(async (node: Node, files: File[]) => {
        if (!node.data.linkedTaskId) return;

        try {
            setIsSyncing(true);
            const listName = node.data.type === 'Main Task' ? LIST_MAIN_TASKS : LIST_SUB_TASKS;
            await taskService.addAttachmentsToItem(listName, node.data.linkedTaskId, files);
            showToast("Files Uploaded", `Successfully attached ${files.length} file(s) to Task #${node.data.linkedTaskId}.`, "success");
        } catch (e) {
            console.error("[WorkflowDesigner] Drop upload error:", e);
            showToast("Upload Failed", "Could not upload files to SharePoint.", "error");
        } finally {
            setIsSyncing(false);
        }
    }, [showToast]);


    const handleExport = async () => {
        if (reactFlowWrapper.current === null) return;

        try {
            setIsSyncing(true);
            const dataUrl = await toPng(reactFlowWrapper.current, {
                backgroundColor: '#ffffff',
                quality: 0.95,
                cacheBust: true,
            });

            const link = document.createElement('a');
            link.download = `Workflow-${mainTaskTitle || 'Design'}-${new Date().getTime()}.png`;
            link.href = dataUrl;
            link.click();
            showToast("Exported", "Workflow image downloaded successfully.", "success");
        } catch (e) {
            console.error("Export Error:", e);
            showToast("Export Failed", "Could not generate workflow image.", "error");
        } finally {
            setIsSyncing(false);
        }
    };

    const handleSendEmail = async () => {
        const node = selectedNode;
        if (!node || node.data.type !== 'Email') return;

        let to = node.data.emailTo;
        const subject = node.data.emailSubject;

        if (!to || !subject) {
            showToast("Missing Info", "Please provide both a Recipient and a Subject.", "warning");
            return;
        }

        // Validate Email
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        // If it's a valid email key from dropdown or freeform
        if (!emailRegex.test(to)) {
            // Try to find it in userOptions by text match if key wasn't an email
            const matched = userOptions.filter(u => u.text === to);
            if (matched.length > 0 && matched[0].key) {
                to = matched[0].key as string;
            } else {
                // Final check - if still not email, warn
                if (!emailRegex.test(to)) {
                    showToast("Invalid Email", `"${to}" is not a valid email address.`, "error");
                    return;
                }
            }
        }

        setIsSyncing(true);
        try {
            await taskService.sendEmail(
                [to],
                subject,
                `<h2>${subject}</h2>
                 <p>This is an automated notification from the Task Tracking Workflow.</p>
                 <p><b>Node ID:</b> ${node.id}</p>
                 <p><b>Label:</b> ${node.data.label}</p>
                 <hr/>
                 <p><i>Sent via SharePoint Task Tracking System</i></p>`,
                mainTaskId,
                undefined // or node.data.linkedTaskId if applicable
            );
            showToast("Email Sent", `Successfully queued for ${to}.`, "success");
        } catch (e) {
            console.error("Email Error:", e);
            showToast("Email Failed", "Could not send notification.", "error");
        } finally {
            setIsSyncing(false);
        }
    };


    const handleCreateTaskWithOverrides = async () => {
        console.log('[Debug] handleCreateTaskWithOverrides called. SelectedNodeId:', selectedNodeId);
        console.log('[Debug] Dialog Inputs:', { newTaskTitle, newTaskAssignee }); // Log inputs

        const node = nodes.filter(n => n.id === selectedNodeId)[0];
        if (!node) {
            console.error('[Debug] No node found for ID:', selectedNodeId);
            showToast("System Error", "No node selected. Try re-selecting the node.", "error");
            return;
        }

        await handleCreateActualTask(node, {
            title: newTaskTitle,
            description: newTaskDesc,
            assignee: newTaskAssignee,
            dueDate: newTaskDueDate
        });

        setIsCreateTaskOpen(false);
    };

    const handleOpenClarification = async () => {
        const selectedNode = nodes.filter(n => n.id === selectedNodeId)[0];
        if (!selectedNode?.data.linkedTaskId) return;

        setIsClarificationOpen(true);
        setIsClarificationLoading(true);
        setClarificationMessage('');
        try {
            const history = await taskService.getCorrespondenceByTaskId(mainTaskId || 0, selectedNode.data.linkedTaskId);
            setClarificationHistory(history);
        } catch (e) {
            console.error("Error fetching clarification history:", e);
        } finally {
            setIsClarificationLoading(false);
        }
    };

    const handleSendClarification = async () => {
        const selectedNode = nodes.filter(n => n.id === selectedNodeId)[0];
        if (!selectedNode?.data.linkedTaskId || !clarificationMessage.trim()) return;

        setIsSyncing(true);
        try {
            // Smart Reply Logic:
            // 1. If there's conversation history, reply to the LAST SENDER
            // 2. Otherwise, send to Main Task Author (the requester)
            // 3. If user IS the author, send to assignee as fallback
            let to = '';

            // Check if there's existing conversation - reply to last sender
            if (clarificationHistory.length > 0) {
                const lastMessage = clarificationHistory[clarificationHistory.length - 1];
                const lastSender = lastMessage.FromAddress || '';

                // Only reply to last sender if it's NOT the current user (avoid self-messaging)
                if (lastSender && lastSender.toLowerCase() !== currentUserEmail.toLowerCase()) {
                    to = lastSender;
                    console.log(`[Clarification] Replying to last sender: ${lastSender}`);
                }
            }

            // Fallback: If no conversation history, try Main Task author
            if (!to) {
                try {
                    const mt = await taskService.getMainTaskById(mainTaskId || 0);
                    if (mt && (mt as any).Author && (mt as any).Author.EMail) {
                        to = (mt as any).Author.EMail;
                    }
                } catch (err) {
                    console.warn("Could not fetch main task author", err);
                }
            }

            // Final fallback: If user IS the author or no author found, send to assignee
            if (!to || to.toLowerCase() === currentUserEmail.toLowerCase()) {
                to = selectedNode.data.assignee || '';
            }

            // Smart subject line based on conversation state
            const subject = clarificationHistory.length > 0
                ? `Reply: ${selectedNode.data.label}`
                : `Clarification Needed: ${selectedNode.data.label}`;

            await taskService.sendEmail(
                [to],
                subject,
                clarificationMessage,
                mainTaskId,
                selectedNode.data.linkedTaskId
            );

            // Clear and Refresh history
            setClarificationMessage('');
            const history = await taskService.getCorrespondenceByTaskId(mainTaskId || 0, selectedNode.data.linkedTaskId);
            setClarificationHistory(history);
        } catch (e) {
            console.error("Error sending clarification:", e);
            alert("Failed to send clarification.");
        } finally {
            setIsSyncing(false);
        }
    };

    const handleZoomToFit = () => {
        fitView({ padding: 50, duration: 800 });
    };

    const selectedNode = nodes.filter((n: Node) => n.id === selectedNodeId)[0];

    // UI Components
    const ContextMenuUI = () => {
        if (!menu) return null;
        return (
            <div className={styles.contextMenu} style={{ top: menu.y, left: menu.x }}>
                <div className={styles.menuItem} onClick={handleDuplicateNode}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Copy size={14} />
                        <span>Duplicate</span>
                    </div>
                    <span className={styles.shortcut}>Ctrl+D</span>
                </div>
                <div className={styles.menuDivider} />
                <div className={`${styles.menuItem} ${styles.danger}`} onClick={handleDeleteSelected}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Trash2 size={14} />
                        <span>Delete</span>
                    </div>
                    <span className={styles.shortcut}>Del</span>
                </div>
                <div className={styles.menuDivider} />
                <div className={styles.menuItem} onClick={handleTidyUp}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Layout size={14} />
                        <span>Tidy Up</span>
                    </div>
                </div>
            </div>
        );
    };

    const WhatsNextPanelUI = () => {
        if (!isWhatsNextOpen || nodes.length === 0) return null;

        const filteredTemplates = nodeTemplates.filter(t =>
            t.type !== 'Main Task' &&
            t.type.toLowerCase().indexOf(librarySearchQuery.toLowerCase()) !== -1
        );

        return (
            <div className={styles.whatsNextPanel}>
                <div className={styles.panelHeader}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                        <h3>What&apos;s Next?</h3>
                        <X size={20} className={styles.closeBtn} style={{ cursor: 'pointer' }} onClick={() => setIsWhatsNextOpen(false)} />
                    </div>
                    <div className={styles.searchWrapper}>
                        <Search size={16} className={styles.searchIcon} />
                        <input
                            type="text"
                            placeholder="Search nodes..."
                            value={librarySearchQuery}
                            onChange={(e) => setLibrarySearchQuery(e.target.value)}
                            autoFocus
                        />
                    </div>
                </div>
                <div className={styles.panelContent}>
                    <div className={styles.nodeGrid}>
                        {filteredTemplates.map((template, idx) => (
                            <div
                                key={idx}
                                className={`${styles.nodeItem} ${addingNodeType === template.type ? styles.isLoading : ''}`}
                                onClick={() => handleQuickAdd(template)}
                            >
                                <div className={styles.nodeIconWrapper}>
                                    <template.icon size={20} color={template.color} />
                                </div>
                                <div className={styles.nodeInfo}>
                                    <span className={styles.nodeTitle}>{template.type}</span>
                                    <span className={styles.nodeDesc}>Add a new {template.type} node</span>
                                </div>
                            </div>
                        ))}
                    </div>
                </div>
            </div>
        );
    };


    return (
        <div className={styles.workflowDesigner}>
            <header className={styles.header}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
                    <h2><Layout size={24} /> Workflow Designer</h2>
                    {mainTaskTitle && (
                        <div style={{
                            padding: '4px 12px',
                            background: 'rgba(255,255,255,0.1)',
                            borderRadius: '20px',
                            fontSize: '14px',
                            border: '1px solid rgba(255,255,255,0.2)',
                            color: '#fff',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px'
                        }}>
                            <span style={{ opacity: 0.7 }}>Main Task:</span>
                            <strong>{mainTaskTitle}</strong> (ID: #{mainTaskId})
                        </div>
                    )}
                    <div style={{
                        padding: '4px 12px',
                        background: 'rgba(59, 130, 246, 0.1)',
                        borderRadius: '20px',
                        fontSize: '14px',
                        border: '1px solid rgba(59, 130, 246, 0.2)',
                        color: '#3b82f6',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '6px'
                    }}>
                        <RefreshCw size={14} className={isSyncing ? 'animate-spin' : ''} />
                        <strong>{nodes.filter(n => n.data.linkedTaskId).length}</strong> Linked |
                        <strong>{nodes.filter(n => !n.data.linkedTaskId && n.data.type === 'Task').length}</strong> Planned
                    </div>
                </div>
                <div className={styles.actions}>
                    <button
                        className={styles.btnSecondary}
                        onClick={handleExport}
                        title="Export as Image"
                        disabled={isSyncing}
                    >
                        <Download size={16} />
                    </button>
                    <PrimaryButton
                        className={styles.btnSecondary}
                        onClick={async () => {
                            setIsSyncing(true);
                            try {
                                const latestTasks = await taskService.getSubTasksByMainTaskIds([mainTaskId || 0]);
                                const mt = mainTaskId ? await taskService.getMainTaskById(mainTaskId) : null;
                                tidyUp();
                                setTimeout(() => fitView({ padding: 50, duration: 800 }), 100);
                            } finally {
                                setIsSyncing(false);
                            }
                        }}
                        disabled={isSyncing}
                    >
                        {isSyncing ? <RefreshCw className="animate-spin" size={16} /> : <Layout size={16} />} Refine Layout
                    </PrimaryButton>
                    <button className={styles.btnSecondary} onClick={() => { refreshWorkflowData(); showToast("Refreshed", "Latest data fetched.", "info"); }} title="Refresh Data">
                        <RefreshCw size={16} className={isSyncing ? 'animate-spin' : ''} />
                    </button>
                    <button className={styles.btnSecondary} onClick={handleZoomToFit} title="Zoom to Fit">
                        <Maximize size={16} />
                    </button>
                    <button className={styles.btnSecondary} onClick={handleClear}>Clear Canvas</button>
                    <button className={styles.btnPrimary} onClick={handleSave} disabled={isSaving}>
                        {isSaving ? <RefreshCw size={16} className="animate-spin" /> : <Plus size={16} />}
                        {isSaving ? 'Saving...' : 'Save Workflow'}
                    </button>
                </div>
            </header>

            <div className={styles.container}>
                {nodes.length === 0 ? (
                    <div className={styles.emptyStateContainer}>
                        <div className={styles.emptyStateIconGroup}>
                            <DraftingCompass size={80} className={styles.mainIcon} strokeWidth={1.5} />
                            <div className={styles.subIcon}>
                                <PlusCircle size={24} strokeWidth={2.5} />
                            </div>
                        </div>
                        <h1 className={styles.emptyStateTitle}>Start Your Workflow</h1>
                        <p className={styles.emptyStateDesc}>
                            Begin by adding your first node. Drag elements from the library or use the quick actions below to map out your process.
                        </p>
                        <div className={styles.emptyStateActions}>
                            <button
                                className={`${styles.startBtn} ${styles.primary}`}
                                onClick={() => setIsCreateMainTaskOpen(true)}
                            >
                                <Layers size={20} /> Add Main Task
                            </button>
                        </div>
                    </div>
                ) : (
                    <aside className={`${styles.sidebar} ${isSidebarMinimized ? styles.minimized : ''}`}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                            {!isSidebarMinimized && <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: 700, color: '#1e293b' }}>Library</h3>}
                            <IconButton
                                iconProps={{ iconName: isSidebarMinimized ? 'ChevronRight' : 'ChevronLeft' }}
                                onClick={() => setIsSidebarMinimized(!isSidebarMinimized)}
                                styles={{ root: { color: '#94a3b8' } }}
                            />
                        </div>

                        {!isSidebarMinimized && (
                            <>
                                <div className={styles.searchWrapper} style={{ marginBottom: '20px' }}>
                                    <div style={{ position: 'relative' }}>
                                        <Search size={16} className={styles.searchIcon} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: '#64748b' }} />
                                        <input
                                            type="text"
                                            placeholder="Search nodes..."
                                            value={librarySearchQuery}
                                            onChange={(e) => setLibrarySearchQuery(e.target.value)}
                                            style={{
                                                width: '100%',
                                                padding: '10px 12px 10px 36px',
                                                border: '1px solid #e2e8f0',
                                                borderRadius: '12px',
                                                fontSize: '0.875rem',
                                                outline: 'none',
                                                background: 'rgba(255,255,255,0.5)',
                                                transition: 'all 0.2s ease'
                                            }}
                                        />
                                    </div>
                                </div>

                                <div className={styles.nodeTemplates}>
                                    {nodeTemplates
                                        .filter(t => t.type.toLowerCase().indexOf(librarySearchQuery.toLowerCase()) !== -1)
                                        .map((t, idx) => {
                                            const isMainTaskAdded = nodes.some(n => n.data.type === 'Main Task');
                                            const isDisabled = t.type === 'Main Task' && isMainTaskAdded;

                                            return (
                                                <div
                                                    key={idx}
                                                    className={`${styles.nodeTemplate} ${isDisabled ? styles.disabled : ''}`}
                                                    onDragStart={(event) => !isDisabled && onDragStart(event, t as any)}
                                                    draggable={!isDisabled}
                                                    title={isDisabled ? "Only one Main Task allowed per workflow" : `Drag to add ${t.type}`}
                                                >
                                                    <div className={styles.iconWrapper} style={{ color: (t as any).color }}>
                                                        {React.createElement(t.icon, { size: 22 })}
                                                    </div>
                                                    <span>{t.type}</span>
                                                </div>
                                            );
                                        })}
                                </div>
                            </>
                        )}
                    </aside>
                )}

                <main className={styles.canvas} onDrop={onDrop} onDragOver={onDragOver} ref={reactFlowWrapper}>
                    <ReactFlow
                        nodes={nodes}
                        edges={edges}
                        onNodesChange={onNodesChange}
                        onEdgesChange={onEdgesChange}
                        onConnect={onConnect}
                        onNodeClick={onNodeClick}
                        onInit={() => setTimeout(handleZoomToFit, 500)}
                        nodeTypes={nodeTypes}
                        connectionLineStyle={{ stroke: '#3b82f6', strokeWidth: 3 }}
                        connectionLineType={'smoothstep' as any}
                        onDrop={onDrop}
                        onDragOver={onDragOver}
                        fitView
                        fitViewOptions={{ padding: 0.2 }}
                        deleteKeyCode={46}
                    >
                        <Background
                            color="#94a3b8"
                            gap={25}
                            size={1.5}
                            variant={BackgroundVariant.Dots}
                            style={{ opacity: 0.4 }}
                        />
                        <Controls />
                        <MiniMap />
                    </ReactFlow>

                    <ContextMenuUI />

                    <div style={{
                        position: 'absolute',
                        bottom: '20px',
                        right: '20px',
                        display: 'flex',
                        gap: '10px',
                        zIndex: 10
                    }}>
                        <button className={styles.btnSecondary} onClick={() => tidyUp()} style={{ padding: '8px 16px', borderRadius: '30px', background: 'white', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.05)', fontWeight: 600 }}>
                            <RefreshCw size={16} /> Tidy Up
                        </button>
                    </div>
                </main>

                <NodePropertiesDialog
                    selectedNodeId={selectedNodeId}
                    nodes={nodes}
                    edges={edges}
                    iconMapping={iconMapping}
                    iconColorMapping={iconColorMapping}
                    readonly={readonly || false}
                    isSyncing={isSyncing}
                    yearOptions={yearOptions}
                    monthOptions={monthOptions}
                    userOptions={userOptions}
                    categoryOptions={categoryOptions}
                    departmentOptions={departmentOptions}
                    updateSelectedNode={updateSelectedNode}
                    handleSyncStatus={handleSyncStatus}
                    handleCreateActualTask={handleCreateActualTask}
                    handleOpenClarification={handleOpenClarification}
                    setNodes={setNodes}
                    setSelectedNodeId={setSelectedNodeId}
                    showToast={showToast}
                />
            </div>

            {/* Toast System */}
            <div className={styles.toastContainer}>
                {toasts.map(toast => (
                    <div key={toast.id} className={`${styles.toast} ${styles[toast.type]} ${toast.removing ? styles.removing : ''}`}>
                        <div className={styles.toastIcon}>
                            {toast.type === 'success' && <CheckCircle size={20} />}
                            {toast.type === 'error' && <AlertCircle size={20} />}
                            {toast.type === 'info' && <Info size={20} />}
                            {toast.type === 'warning' && <AlertTriangle size={20} />}
                        </div>
                        <div style={{ flex: 1 }}>
                            <div style={{ fontWeight: 700, fontSize: '13px' }}>{toast.title}</div>
                            <div style={{ fontSize: '12px', opacity: 0.9 }}>{toast.message}</div>
                        </div>
                        <X size={14} className={styles.toastClose} onClick={() => {
                            setToasts(prev => prev.filter(t => t.id !== toast.id));
                        }} />
                    </div>
                ))}
            </div>

            <WhatsNextPanelUI />

            <Dialog
                hidden={!isCreateTaskOpen}
                onDismiss={() => setIsCreateTaskOpen(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Create SharePoint Task',
                    subText: 'Link this node to an actual task in SharePoint.'
                }}
            >
                <Stack tokens={{ childrenGap: 15 }}>
                    <TextField label="Title" value={newTaskTitle} onChange={(_, v) => setNewTaskTitle(v || '')} required />
                    <TextField label="Description" multiline rows={3} value={newTaskDesc} onChange={(_, v) => setNewTaskDesc(v || '')} />
                    <ComboBox label="Assignee" required options={userOptions} selectedKey={newTaskAssignee} onChange={(_, opt) => setNewTaskAssignee(opt?.key as string)} placeholder="Select or type a user..." />
                    <TextField label="Due Date" type="date" value={newTaskDueDate ? newTaskDueDate.toISOString().split('T')[0] : ''} onChange={(_, v) => setNewTaskDueDate(v ? new Date(v) : undefined)} />
                </Stack>
                <DialogFooter>
                    <PrimaryButton
                        onClick={handleCreateTaskFromDialog}
                        text={isSyncing ? "Creating..." : "Create Task"}
                        disabled={isSyncing}
                        iconProps={isSyncing ? { iconName: 'Sync' } : { iconName: 'Add' }}
                    />
                    <DefaultButton onClick={() => setIsCreateTaskOpen(false)} text="Cancel" disabled={isSyncing} />
                </DialogFooter>
            </Dialog>

            <Dialog
                hidden={!isClarificationOpen}
                onDismiss={() => setIsClarificationOpen(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Clarification History',
                }}
                minWidth={550}
            >
                <div style={{ maxHeight: '400px', overflowY: 'auto', padding: '10px', background: '#f8fafc', borderRadius: '8px', marginBottom: '15px' }}>
                    {isClarificationLoading ? <div style={{ textAlign: 'center', padding: 20 }}>Loading...</div> : clarificationHistory.map((h, i) => (
                        <div key={i} style={{
                            padding: '10px',
                            background: h.FromAddress === currentUserEmail ? '#eff6ff' : 'white',
                            border: '1px solid #e2e8f0',
                            borderRadius: '8px',
                            marginBottom: '10px',
                            alignSelf: h.FromAddress === currentUserEmail ? 'flex-end' : 'flex-start'
                        }}>
                            <div style={{ fontSize: '11px', color: '#64748b', marginBottom: '4px' }}>
                                <strong>{h.Author?.Title}</strong> • {new Date(h.Created).toLocaleString()}
                            </div>
                            <div dangerouslySetInnerHTML={{ __html: sanitizeHtml(h.MessageBody) }} style={{ fontSize: '13px' }} />
                        </div>
                    ))}
                    {clarificationHistory.length === 0 && !isClarificationLoading && <div style={{ textAlign: 'center', color: '#64748b', padding: 20 }}>No history.</div>}
                </div>
                <TextField label="Message" multiline rows={3} value={clarificationMessage} onChange={(_, v) => setClarificationMessage(v || '')} />
                <DialogFooter>
                    <PrimaryButton onClick={handleSendClarification} text="Send" disabled={!clarificationMessage || isSyncing} />
                    <DefaultButton onClick={() => setIsClarificationOpen(false)} text="Close" />
                </DialogFooter>
            </Dialog>



            <Dialog
                hidden={!isCreateMainTaskOpen}
                onDismiss={() => setIsCreateMainTaskOpen(false)}
                dialogContentProps={{
                    type: DialogType.largeHeader,
                    title: 'Create New Main Task',
                }}
                modalProps={{
                    isBlocking: true,
                    styles: { main: { minWidth: '650px !important', borderRadius: '12px' } },
                }}
            >
                <Stack tokens={{ childrenGap: 15 }}>
                    <TextField label="Task Title" required value={newMainTaskTitle} onChange={(e, v) => setNewMainTaskTitle(v || '')} />
                    <TextField label="Description" required multiline rows={3} value={newMainTaskDesc} onChange={(e, v) => setNewMainTaskDesc(v || '')} />
                    <Stack horizontal tokens={{ childrenGap: 15 }}>
                        <Dropdown
                            label="Year"
                            required
                            options={yearOptions}
                            selectedKey={newMainTaskYear}
                            onChange={(e, o) => setNewMainTaskYear(o?.key as string)}
                            styles={{ root: { width: '50%' } }}
                        />
                        <Dropdown
                            label="Month"
                            required
                            options={monthOptions}
                            selectedKey={newMainTaskMonth}
                            onChange={(e, o) => setNewMainTaskMonth(o?.key as string)}
                            styles={{ root: { width: '50%' } }}
                        />
                    </Stack>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                        <span style={{ fontWeight: 600 }}>Assign To</span>
                        <IconButton iconProps={{ iconName: 'Refresh' }} title="Reload Users" onClick={loadGlobalOptions} />
                    </Stack>
                    <ComboBox
                        required
                        autoComplete="on"
                        options={userOptions}
                        selectedKey={newMainTaskAssignee}
                        onChange={(e, o, i, v) => setNewMainTaskAssignee(o ? o.key as string : v)}
                        placeholder="Select or type a name..."
                        calloutProps={{ doNotLayer: true }}
                        dropdownMaxWidth={400}
                    />
                    <Dropdown
                        label="Department"
                        options={departmentOptions}
                        selectedKey={newMainTaskDept}
                        onChange={(e, o) => setNewMainTaskDept(o?.key as string)}
                    />
                    <TextField label="Project" value={newMainTaskProject} onChange={(e, v) => setNewMainTaskProject(v || '')} />
                    <TextField
                        label="Due Date"
                        type="date"
                        required
                        value={newMainTaskDueDate ? newMainTaskDueDate.toISOString().split('T')[0] : ''}
                        onChange={(e, v) => setNewMainTaskDueDate(v ? new Date(v) : undefined)}
                    />
                    <div>
                        <Label>Attachments</Label>
                        <input type="file" multiple onChange={(e) => setNewMainTaskFiles(Array.from(e.target.files || []))} />
                    </div>

                </Stack>
                <DialogFooter>
                    <PrimaryButton
                        onClick={handleCreateMainTaskFromDialog}
                        text={isSyncing ? "Creating..." : "Create Task"}
                        disabled={isSyncing}
                        iconProps={isSyncing ? { iconName: 'Sync' } : { iconName: 'Add' }}
                    />
                    <DefaultButton onClick={() => setIsCreateMainTaskOpen(false)} text="Cancel" disabled={isSyncing} />
                </DialogFooter>
            </Dialog>
        </div >
    );
};
