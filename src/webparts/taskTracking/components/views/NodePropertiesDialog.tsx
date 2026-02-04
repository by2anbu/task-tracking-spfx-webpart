/* eslint-disable @rushstack/no-new-null */
import * as React from 'react';
import { Dialog, DialogType, DialogFooter, TextField, Dropdown, ComboBox, PrimaryButton, DefaultButton, Stack, IComboBoxOption, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { Node, Edge } from 'react-flow-renderer';
import { CheckCircle, AlertCircle, Trash2, X, Info, AlertTriangle, ShieldAlert } from 'lucide-react';
import styles from './WorkflowDesigner.module.scss';

export interface INodePropertiesDialogProps {
    selectedNodeId: string | null;
    nodes: Node[];
    edges: Edge[];
    iconMapping: any;
    iconColorMapping: any;
    readonly: boolean;
    isSyncing: boolean;
    yearOptions: IComboBoxOption[];
    monthOptions: IComboBoxOption[];
    userOptions: IComboBoxOption[];
    categoryOptions: IComboBoxOption[];
    departmentOptions: IComboBoxOption[];
    updateSelectedNode: (key: string, value: any) => void;
    handleSyncStatus: (node: Node) => void;
    handleCreateActualTask: (node: Node) => void;
    handleOpenClarification: () => void;
    setNodes: React.Dispatch<React.SetStateAction<Node[]>>;
    setSelectedNodeId: (id: string | null) => void;
    showToast: (title: string, message: string, type?: 'success' | 'info' | 'warning' | 'error') => void;
}

export const NodePropertiesDialog: React.FC<INodePropertiesDialogProps> = (props) => {
    const {
        selectedNodeId,
        nodes,
        edges,
        iconMapping,
        iconColorMapping,
        readonly,
        isSyncing,
        yearOptions,
        monthOptions,
        userOptions,
        categoryOptions,
        departmentOptions,
        updateSelectedNode,
        handleSyncStatus,
        handleCreateActualTask,
        handleOpenClarification,
        setNodes,
        setSelectedNodeId,
        showToast
    } = props;

    if (!selectedNodeId || !nodes.find(n => n.id === selectedNodeId)) {
        return null;
    }

    const node = nodes.find(n => n.id === selectedNodeId)!;
    const hasChildren = edges.some(edge => edge.source === node.id);

    return (
        <Dialog
            hidden={!selectedNodeId}
            onDismiss={() => setSelectedNodeId(null)}
            className={styles.propertyModal}
            dialogContentProps={{
                type: DialogType.normal,
                title: (
                    <div className={styles.propHeader}>
                        <div className={styles.propIconBox} style={{
                            background: `linear-gradient(135deg, ${(iconColorMapping[node.data.type] || '#3b82f6')}15 0%, ${(iconColorMapping[node.data.type] || '#3b82f6')}25 100%)`,
                            color: iconColorMapping[node.data.type] || '#3b82f6',
                            border: `2px solid ${(iconColorMapping[node.data.type] || '#3b82f6')}30`,
                        }}>
                            {React.createElement(iconMapping[node.data.type] || CheckCircle, { size: 24 })}
                        </div>
                        <div className={styles.propTitleGroup}>
                            <div className={styles.propTitle}>
                                Node Properties
                            </div>
                            <div className={styles.propBadge} style={{
                                background: `${(iconColorMapping[node.data.type] || '#3b82f6')}15`,
                                color: iconColorMapping[node.data.type] || '#3b82f6',
                            }}>
                                {node.data.type}
                            </div>
                        </div>
                        <div className={styles.propHeaderClose} onClick={() => setSelectedNodeId(null)}>
                            <X size={20} />
                        </div>
                    </div>
                )
            }}
            modalProps={{
                isBlocking: false,
                styles: { main: { minWidth: '700px !important', maxWidth: '90vw' } }
            }}
        >
            <div className={styles.propContent}>
                <div className={styles.propGrid}>
                    {/* Left Column - Basic Info */}
                    <div className={styles.propSection}>
                        <div>
                            <div className={styles.propSectionHeader}>
                                <div className={styles.accent} />
                                Basic Information
                            </div>
                            <div className={styles.propCard}>
                                <TextField
                                    label={(node.data.type === 'Task' || node.data.type === 'Main Task') ? 'Task Title' : 'Node Label'}
                                    value={node.data.label}
                                    onChange={(_, val) => updateSelectedNode('label', val || '')}
                                    disabled={readonly || node.data.type === 'Start'}
                                    required={(node.data.type === 'Task' || node.data.type === 'Main Task')}
                                    className={styles.propTextField}
                                />
                                <TextField
                                    label="Description"
                                    value={node.data.description || ''}
                                    onChange={(_, val) => updateSelectedNode('description', val || '')}
                                    multiline
                                    rows={4}
                                    disabled={readonly}
                                    className={styles.propTextField}
                                />
                            </div>
                        </div>

                        {node.data.type === 'Main Task' && (
                            <div>
                                <div className={styles.propSectionHeader}>
                                    <div className={styles.accent} />
                                    Assignment
                                </div>
                                <div className={styles.propCard}>
                                    <Stack tokens={{ childrenGap: 12 }}>
                                        <Stack horizontal tokens={{ childrenGap: 12 }}>
                                            <Dropdown
                                                label="Year"
                                                options={yearOptions}
                                                selectedKey={node.data.year}
                                                onChange={(_, opt) => updateSelectedNode('year', opt?.key)}
                                                className={styles.propDropdown}
                                                styles={{ root: { flex: 1 } }}
                                                disabled={readonly}
                                            />
                                            <Dropdown
                                                label="Month"
                                                options={monthOptions}
                                                selectedKey={node.data.month}
                                                onChange={(_, opt) => updateSelectedNode('month', opt?.key)}
                                                className={styles.propDropdown}
                                                styles={{ root: { flex: 2 } }}
                                                disabled={readonly}
                                            />
                                        </Stack>
                                        <ComboBox
                                            label="Assign To"
                                            options={userOptions}
                                            selectedKey={node.data.assignee}
                                            onChange={(_, option, __, value) => updateSelectedNode('assignee', option ? option.key as string : value || '')}
                                            disabled={readonly}
                                            autoComplete="on"
                                            required
                                            placeholder="Select or type user..."
                                            className={styles.propComboBox}
                                        />
                                    </Stack>
                                </div>
                            </div>
                        )}

                        {node.data.type === 'Task' && (
                            <div>
                                <div className={styles.propSectionHeader}>
                                    <div className={styles.accent} />
                                    Assignment
                                </div>
                                <div className={styles.propCard}>
                                    <Stack tokens={{ childrenGap: 12 }}>
                                        <ComboBox
                                            label="Assign To"
                                            options={userOptions}
                                            selectedKey={node.data.assignee}
                                            onChange={(_, option, __, value) => updateSelectedNode('assignee', option ? option.key as string : value || '')}
                                            disabled={readonly}
                                            autoComplete="on"
                                            required
                                            placeholder="Select or type user..."
                                            className={styles.propComboBox}
                                        />
                                        <Dropdown
                                            label="Category"
                                            options={categoryOptions}
                                            selectedKey={node.data.category}
                                            onChange={(_, option) => updateSelectedNode('category', option?.key as string)}
                                            disabled={readonly}
                                            className={styles.propDropdown}
                                        />
                                    </Stack>
                                </div>
                            </div>
                        )}
                        {node.data.type === 'Condition' && (
                            <div>
                                <div className={styles.propSectionHeader}>
                                    <div className={styles.accent} />
                                    Condition Logic
                                </div>
                                <div className={styles.propCard}>
                                    <Stack tokens={{ childrenGap: 12 }}>
                                        <MessageBar messageBarType={MessageBarType.info}>
                                            Conditions currently act as &quot;Pass-through&quot; or &quot;Decision Points&quot;. Subsequent nodes will be triggered when the preceding node is completed.
                                        </MessageBar>
                                        <TextField
                                            label="Branch Name / Rule"
                                            value={node.data.label}
                                            onChange={(_, val) => updateSelectedNode('label', val || '')}
                                            placeholder="e.g. If Approved"
                                            disabled={readonly}
                                            required
                                        />
                                        <Dropdown
                                            label="Condition Type"
                                            options={[
                                                { key: 'always', text: 'Always Pass (Default)' }
                                            ]}
                                            selectedKey="always"
                                            disabled
                                        />
                                    </Stack>
                                </div>
                            </div>
                        )}
                        {node.data.type === 'Alert' && (
                            <div>
                                <div className={styles.propSectionHeader}>
                                    <div className={styles.accent} />
                                    Alert Configuration
                                </div>
                                <div className={styles.propCard}>
                                    <Stack tokens={{ childrenGap: 12 }}>
                                        <TextField
                                            label="Alert Name"
                                            value={node.data.label}
                                            onChange={(_, val) => updateSelectedNode('label', val || '')}
                                            placeholder="e.g. Overdue Reminder"
                                            disabled={readonly}
                                            required
                                        />
                                        <TextField
                                            label="Email Subject"
                                            value={node.data.emailSubject || ''}
                                            onChange={(_, val) => updateSelectedNode('emailSubject', val || '')}
                                            disabled={readonly}
                                        />
                                        <TextField
                                            label="Message Body"
                                            value={node.data.description || ''}
                                            onChange={(_, val) => updateSelectedNode('description', val || '')}
                                            multiline
                                            rows={4}
                                            disabled={readonly}
                                        />
                                        <Dropdown
                                            label="Notify Who?"
                                            options={[
                                                { key: 'assignee', text: 'Task Assignee' },
                                                { key: 'owner', text: 'Task Owner (Creator)' },
                                                { key: 'both', text: 'Both Assignee & Owner' }
                                            ]}
                                            selectedKey={node.data.notifyWho || 'assignee'}
                                            onChange={(_, opt) => updateSelectedNode('notifyWho', opt?.key)}
                                            disabled={readonly}
                                        />
                                    </Stack>
                                </div>
                            </div>
                        )}
                    </div>

                    {/* Right Column - SharePoint Integration */}
                    {(node.data.type === 'Task' || node.data.type === 'Main Task') && (
                        <div className={styles.propSection}>
                            <div className={styles.propSectionHeader}>
                                <div className={`${styles.accent} ${styles.sp}`} />
                                SharePoint Integration
                            </div>
                            <div className={styles.spIntegrationCard}>
                                {node.data.linkedTaskId ? (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                                        {node.data.isBlocked && (
                                            <MessageBar
                                                messageBarType={MessageBarType.warning}
                                                actions={
                                                    <div style={{ display: 'flex', gap: '8px' }}>
                                                        <ShieldAlert size={16} />
                                                        <strong>Task Blocked</strong>
                                                    </div>
                                                }
                                            >
                                                This task cannot be proceeded with because one or more parent tasks are not yet &quot;Completed&quot;.
                                            </MessageBar>
                                        )}
                                        <div className={styles.spTaskDetails}>
                                            <div>
                                                <div className={styles.idLabel}>Task ID</div>
                                                <div className={styles.idValue}>#{node.data.linkedTaskId}</div>
                                            </div>
                                            <div className={`${styles.propStatusBadge} ${node.data.status === 'Completed' ? styles.completed :
                                                node.data.status === 'In Progress' ? styles.inProgress :
                                                    styles.notStarted
                                                }`}>
                                                {node.data.status || 'Not Started'}
                                            </div>
                                        </div>
                                        <PrimaryButton
                                            text={isSyncing ? 'Syncing...' : 'Sync Status'}
                                            iconProps={{ iconName: 'Sync' }}
                                            onClick={() => handleSyncStatus(node)}
                                            disabled={isSyncing || node.data.isBlocked}
                                            className={styles.btnPrimary}
                                            style={{ height: '42px', width: '100%' }}
                                        />
                                        <DefaultButton
                                            text="Check Clarifications"
                                            iconProps={{ iconName: 'Comment' }}
                                            onClick={handleOpenClarification}
                                            className={styles.btnSecondary}
                                            style={{ height: '42px', width: '100%', borderColor: '#d97706', color: '#92400e' }}
                                        />
                                    </div>
                                ) : (
                                    <div style={{ textAlign: 'center' }}>
                                        <div style={{
                                            fontSize: '14px',
                                            color: '#92400e',
                                            marginBottom: '16px',
                                            fontWeight: 600
                                        }}>
                                            This node is not linked to SharePoint yet
                                        </div>
                                        <PrimaryButton
                                            text={isSyncing ? 'Creating...' : 'Create SharePoint Task'}
                                            onClick={() => handleCreateActualTask(node)}
                                            disabled={isSyncing || readonly}
                                            className={styles.btnPrimary}
                                            style={{ height: '44px', width: '100%', background: 'linear-gradient(135deg, #f59e0b 0%, #ef4444 100%)' }}
                                        />
                                    </div>
                                )}
                            </div>
                        </div>
                    )}
                </div>

                {/* Delete Button */}
                <div className={styles.propDeleteSection}>
                    {hasChildren && (
                        <div style={{
                            display: 'flex',
                            gap: '12px',
                            background: '#fff7ed',
                            border: '1px solid #ffedd5',
                            borderRadius: '12px',
                            padding: '16px',
                            marginBottom: '16px',
                            color: '#9a3412',
                            fontSize: '13px',
                            lineHeight: '1.5'
                        }}>
                            <ShieldAlert size={20} style={{ flexShrink: 0, marginTop: '2px' }} />
                            <div>
                                <strong>Cannot Delete Parent Node</strong>
                                <br />
                                This node has child connections. Please delete all outgoing connections or subsequent nodes before removing this parent.
                            </div>
                        </div>
                    )}
                    <DefaultButton
                        text="Delete Node"
                        iconProps={{ iconName: 'Delete' }}
                        disabled={hasChildren || readonly}
                        onClick={() => {
                            if (confirm("Delete this node?")) {
                                setNodes(nds => nds.filter(n => n.id !== node.id));
                                setSelectedNodeId(null);
                                showToast("Node Deleted", "The node was removed.", "info");
                            }
                        }}
                        className={styles.btnSecondary}
                        style={{
                            width: '100%',
                            color: hasChildren ? '#94a3b8' : '#ef4444',
                            borderColor: hasChildren ? '#e2e8f0' : '#fecaca',
                            height: '42px',
                            background: hasChildren ? '#f1f5f9' : 'transparent'
                        }}
                    />
                </div>
            </div>
            <DialogFooter>
                <DefaultButton
                    onClick={() => setSelectedNodeId(null)}
                    text="Close"
                    iconProps={{ iconName: 'Cancel' }}
                    className={styles.btnSecondary}
                    style={{ height: '40px' }}
                />
            </DialogFooter>
        </Dialog>
    );
};
