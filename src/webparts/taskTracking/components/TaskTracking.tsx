import * as React from 'react';
import styles from './TaskTracking.module.scss';
import { ITaskTrackingProps } from './ITaskTrackingProps';
import { Stack, Persona, PersonaSize, Nav, INavLink, INavLinkGroup, IconButton, Icon } from 'office-ui-fabric-react';
import { AdminDashboard } from './views/AdminDashboard';
import { SMTView } from './views/SMTView';
import { SubtaskView } from './views/SubtaskView';
import { UserDashboard } from './views/UserDashboard';
import { taskService } from '../../../services/sp-service';
import { CorrespondenceView } from './views/CorrespondenceView';
import { ConsolidatedReportView } from './views/ConsolidatedReportView';
import { SubtaskDashboard } from './views/SubtaskDashboard';
import { WorkflowDesigner } from './views/WorkflowDesigner';
import { GanttChartView } from './views/GanttChartView';
import { GlobalNotificationBell } from './common/GlobalNotificationBell';
import { IMainTask, ISubTask } from '../../../services/interfaces';

export interface ITaskTrackingState {
  isAdmin: boolean;
  isSMT: boolean;
  loading: boolean;
  error?: Error;
  selectedKey: string;
  navCollapsed: boolean;
  // URL-based task navigation
  initialParentTaskId?: number;
  initialChildTaskId?: number;
  initialViewTaskId?: number;
  initialViewTab?: string;
  // Data for Gantt
  mainTasks: IMainTask[];
  subTasks: ISubTask[];
}

export default class TaskTracking extends React.Component<ITaskTrackingProps, ITaskTrackingState> {
  constructor(props: ITaskTrackingProps) {
    super(props);
    this.state = {
      isAdmin: false,
      isSMT: false,
      loading: true,
      navCollapsed: false,
      selectedKey: 'subtasks', // Default to My Subtasks instead of dashboard
      initialViewTaskId: props.viewTaskId,
      mainTasks: [],
      subTasks: []
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._initAppData();
  }

  public async _initAppData(): Promise<void> {
    const email = this.props.userEmail || '';
    // Global style injection for pulse
    if (!document.getElementById('task-tracking-global-styles')) {
      const style = document.createElement('style');
      style.id = 'task-tracking-global-styles';
      style.innerText = `
            @keyframes pulse {
                0% { transform: scale(1); opacity: 1; }
                50% { transform: scale(1.15); opacity: 0.8; }
                100% { transform: scale(1); opacity: 1; }
            }
        `;
      document.head.appendChild(style);
    }

    try {
      // 1. Check permissions
      const isAdmin = await taskService.verifyAdminUser(email);
      const mainTasks = await taskService.getMainTasksForUser(email);
      const isSMT = isAdmin || mainTasks.length > 0;

      // 2. Determine default view
      let defaultKey = 'subtasks';

      // Deep link prioritization
      if (this.props.viewTaskId) {
        if (isAdmin) defaultKey = 'admin';
        else defaultKey = 'maintasks';
      } else if (this.props.parentTaskId || this.props.childTaskId) {
        defaultKey = 'maintasks';
      } else if (isAdmin) {
        defaultKey = 'admin';
      } else if (isSMT) {
        defaultKey = 'maintasks';
      }

      // 2. Fetch all tasks for Gantt/Global view
      const allMainTasks = await taskService.getAllMainTasks();
      const allSubTasks = await taskService.getAllSubTasks();

      this.setState({
        isAdmin,
        isSMT,
        loading: false,
        selectedKey: defaultKey,
        mainTasks: allMainTasks,
        subTasks: allSubTasks
      });
    } catch (e) {
      console.error("!!! ERROR loading permissions !!!", e);
      this.setState({ isAdmin: false, isSMT: false, loading: false });
    }
  }


  // Method to handle notification click – now smarter about navigation and tab selection
  private handleNotificationClick = (taskId: number, isSubtask: boolean, parentId?: number, initialTab?: string) => {
    console.log('[TaskTracking] Notification Clicked:', { taskId, isSubtask, parentId, initialTab });

    // Reset state first to ensure deep link effect re-runs if clicking same ID? 
    // Actually, setting state directly overwrites.
    // We want to force a re-navigation if needed.

    const parentTaskId = isSubtask ? parentId : taskId;

    // We clear first to ensure fresh props flow down (optional but safer for deep link triggers)
    this.setState({
      initialParentTaskId: undefined,
      initialChildTaskId: undefined,
      initialViewTaskId: undefined,
      initialViewTab: undefined
    }, () => {
      this.setState({
        // Logic Update: If Admin, ALWAYS go to Admin Dashboard
        selectedKey: this.state.isAdmin ? 'admin' : (isSubtask ? 'subtasks' : 'maintasks'),

        // For Admin Dashboard:
        initialViewTaskId: this.state.isAdmin ? (parentTaskId || taskId) : undefined,

        // For All Views: Deep link props
        initialParentTaskId: parentTaskId,
        initialChildTaskId: isSubtask ? taskId : undefined,

        initialViewTab: initialTab // <--- Store intended tab
      });
    });
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error("TaskTracking Crash:", error, errorInfo);
    this.setState({ error: error } as any);
  }

  private _onLinkClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
    if (item) {
      this.setState({ selectedKey: item.key ?? '' });
    }
  }

  private _clearMainDeepLink = (): void => {
    this.setState({
      initialParentTaskId: undefined,
      initialViewTaskId: undefined
    });
  }

  private _clearChildDeepLink = (): void => {
    this.setState({
      initialChildTaskId: undefined,
      initialViewTab: undefined
    });
  }

  private _renderContent = () => {
    const { selectedKey, isAdmin, initialParentTaskId, initialChildTaskId, initialViewTaskId, initialViewTab } = this.state;
    const { userEmail } = this.props;

    switch (selectedKey) {
      case 'admin': return isAdmin ? (
        <AdminDashboard
          initialViewTaskId={initialViewTaskId}
          initialParentTaskId={initialParentTaskId}
          initialChildTaskId={initialChildTaskId}
          initialTab={initialViewTab}
          onMainDeepLinkProcessed={this._clearMainDeepLink}
          onChildDeepLinkProcessed={this._clearChildDeepLink}
        />
      ) : <UserDashboard userEmail={userEmail} />;
      case 'maintasks': return (
        <SMTView
          userEmail={userEmail}
          initialParentTaskId={initialParentTaskId}
          initialChildTaskId={initialChildTaskId}
          initialViewTaskId={initialViewTaskId}
          onMainDeepLinkProcessed={this._clearMainDeepLink}
          onChildDeepLinkProcessed={this._clearChildDeepLink}
        />
      );
      case 'subtasks': return (
        <SubtaskView
          userEmail={userEmail}
          initialChildTaskId={initialChildTaskId}
          onDeepLinkProcessed={this._clearChildDeepLink}
        />
      );
      case 'dashboard': return <UserDashboard userEmail={userEmail} />;
      case 'subtask_dashboard': return <SubtaskDashboard userEmail={userEmail} />;
      case 'consolidated_report': return isAdmin ? <ConsolidatedReportView userEmail={userEmail} /> : <UserDashboard userEmail={userEmail} />;
      case 'correspondence': return <CorrespondenceView userEmail={userEmail} isAdmin={isAdmin} />;
      case 'workflow_designer': return (isAdmin || this.state.isSMT) ? <WorkflowDesigner userEmail={userEmail} userDisplayName={this.props.userDisplayName} /> : <UserDashboard userEmail={userEmail} />;
      case 'gantt': return <GanttChartView mainTasks={this.state.mainTasks} subTasks={this.state.subTasks} onTaskClick={this.handleNotificationClick} />;
      default: return <UserDashboard userEmail={userEmail} />;
    }
  }

  // Collapse/Expand button
  private _toggleNav = (): void => {
    this.setState({ navCollapsed: !this.state.navCollapsed });
  };

  private _renderSidebar = (navGroups: INavLinkGroup[]): React.ReactElement => {
    const { navCollapsed, selectedKey } = this.state;
    const sidebarStyle: React.CSSProperties = {
      width: navCollapsed ? 40 : 250,
      borderRight: '1px solid #eee',
      background: '#fafafa',
      transition: 'width 0.2s ease'
    };
    return (
      <div style={sidebarStyle}>
        {/* Collapse/Expand button */}
        <IconButton
          iconProps={{ iconName: navCollapsed ? 'ChevronRight' : 'ChevronLeft' }}
          title={navCollapsed ? 'Expand menu' : 'Collapse menu'}
          ariaLabel={navCollapsed ? 'Expand menu' : 'Collapse menu'}
          onClick={this._toggleNav}
          styles={{ root: { margin: 4 } }}
        />
        <Nav
          groups={navGroups}
          selectedKey={selectedKey}
          onLinkClick={this._onLinkClick}
          styles={{
            root: {
              width: '100%',
              boxSizing: 'border-box',
              overflowY: 'auto'
            }
          }}
        />
      </div>
    );
  };

  // Updated render with Correspondence handler
  private handleViewAllCorrespondence = () => {
    this.setState({ selectedKey: 'correspondence' });
  }

  // Updated render method to use _renderSidebar
  public render(): React.ReactElement<ITaskTrackingProps> {
    const { isAdmin, isSMT, loading, selectedKey, navCollapsed } = this.state;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const appState = this.state as any;

    if (appState.error) {
      return (
        <div style={{ color: 'red', padding: 20 }}>
          <h2>Something went wrong</h2>
          <pre>{appState.error.toString()}</pre>
        </div>
      );
    }
    const { userEmail, userDisplayName } = this.props;

    if (loading) return <div>Checking permissions...</div>;

    // Build Nav Groups (same as before)
    const links: INavLink[] = [];
    if (isAdmin) {
      links.push({ name: 'Admin View', url: '', key: 'admin', icon: 'ViewDashboard' });
      // NEW: Consolidated Report
      links.push({ name: 'Consolidated Task View', url: '', key: 'consolidated_report', icon: 'ReportDocument' });
      // Gantt Chart for Admin - Moved up
      links.push({ name: 'Gantt Chart', url: '', key: 'gantt', icon: 'Calendar' });
    }
    if (isSMT || isAdmin) {
      links.push({ name: 'My Main Tasks', url: '', key: 'maintasks', icon: 'TaskGroup' });
      // NEW: Workflow Designer

    }
    links.push({ name: 'My Subtasks', url: '', key: 'subtasks', icon: 'TaskManager' });
    // links.push({ name: 'Correspondence', url: '', key: 'correspondence', icon: 'Mail' });

    if (isAdmin) {
      links.push({ name: 'Workflow Designer', url: '', key: 'workflow_designer', icon: 'Workflow' });
    } else if (isSMT) {
      links.push({ name: 'Gantt Chart', url: '', key: 'gantt', icon: 'Calendar' });
    }
    console.log('TaskTracking render links:', links);
    const navGroups: INavLinkGroup[] = [{ links }];

    return (
      <div className={styles.taskTracking} style={{ minHeight: '600px', background: '#fff' }}>
        {/* Header - Professional Redesign */}
        <div style={{
          display: 'grid',
          gridTemplateColumns: '1fr auto 1fr',
          alignItems: 'center',
          padding: '10px 24px',
          background: '#ffffff',
          boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
          zIndex: 100,
          position: 'relative'
        }}>
          {/* Left: Logo/Brand */}
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <div style={{
              width: 36,
              height: 36,
              background: '#0078d4',
              borderRadius: 6,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              marginRight: 12
            }}>
              <Icon iconName="TaskManager" styles={{ root: { fontSize: 20, color: '#fff' } }} />
            </div>
            <span style={{ fontSize: '18px', fontWeight: 600, color: '#323130' }}>
              TSP
            </span>
          </div>

          {/* Center: Title */}
          <div style={{ textAlign: 'center' }}>
            <h1 style={{
              margin: 0,
              fontSize: '20px',
              fontWeight: 700, // Thicker font for title
              color: '#201f1e',
              letterSpacing: '-0.02em'
            }}>
              Task Tracking System
            </h1>
          </div>

          {/* Right: User Profile */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end' }}>
            <GlobalNotificationBell
              onNotificationClick={this.handleNotificationClick}
            />
            <Persona
              size={PersonaSize.size32} // Slightly smaller, cleaner
              text={userDisplayName}
              secondaryText={userEmail}
              styles={{
                root: { cursor: 'pointer' },
                primaryText: { fontSize: '14px', fontWeight: 600 },
                secondaryText: { fontSize: '12px' }
              }}
            />
          </div>
        </div>
        {/* Main Layout */}
        <Stack horizontal style={{ height: 'calc(100vh - 100px)' }}>
          {/* Sidebar */}
          {this._renderSidebar(navGroups)}
          {/* Content */}
          <div style={{ flexGrow: 1, padding: 20, overflowY: 'auto', overflowX: 'hidden', minWidth: 0, background: '#fff' }}>
            {this._renderContent()}
          </div>
        </Stack>
      </div>
    );
  }
}
