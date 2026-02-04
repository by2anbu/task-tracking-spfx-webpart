import * as React from 'react';
import {
    Icon,
    Callout,
    DirectionalHint,
    Stack,
    Text,
    IconButton,
    Separator,
    FocusTrapZone,
    FontWeights
} from 'office-ui-fabric-react';
import { taskService } from '../../../../services/sp-service';
import { ITaskCorrespondence } from '../../../../services/interfaces';
import styles from './GlobalNotificationBell.module.scss';
import { Bell, CheckCheck, X, Trash2, RefreshCcw, MailOpen, User, Clock } from 'lucide-react';

export interface IGlobalNotificationBellProps {
    onNotificationClick: (taskId: number, isSubtask: boolean, parentTaskId?: number, initialTab?: string) => void;
    onViewAllClick?: () => void;
}

const STORAGE_KEY = 'task_tracking_dismissed_notifications';

export const GlobalNotificationBell: React.FC<IGlobalNotificationBellProps> = (props) => {
    const [notifications, setNotifications] = React.useState<ITaskCorrespondence[]>([]);
    const [dismissedIds, setDismissedIds] = React.useState<number[]>([]);
    const [showCallout, setShowCallout] = React.useState(false);
    const menuButtonRef = React.useRef<HTMLDivElement>(null);
    const prevCountRef = React.useRef(0);
    const [viewHistory, setViewHistory] = React.useState(false);

    // Load dismissed IDs from local storage
    React.useEffect(() => {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) {
            try {
                setDismissedIds(JSON.parse(stored));
            } catch (e) {
                console.error("Failed to parse dismissed IDs", e);
            }
        }
    }, []);

    const fetchNotifications = async () => {
        try {
            const data = await taskService.getGlobalNotifications();
            const userEmail = await taskService.getCurrentUserEmail();
            const incoming = data.filter(n => (n.FromAddress || '').toLowerCase() !== userEmail.toLowerCase());

            setNotifications(incoming);

            const activeCount = incoming.filter(n => dismissedIds.indexOf(n.Id) === -1).length;
            prevCountRef.current = activeCount;
        } catch (e) {
            console.error("Failed to fetch global notifications", e);
        }
    };

    React.useEffect(() => {
        fetchNotifications();
        const interval = setInterval(fetchNotifications, 60000);
        return () => clearInterval(interval);
    }, [dismissedIds]);

    const onBellClick = () => {
        setShowCallout(!showCallout);
        if (!showCallout) {
            fetchNotifications();
            setViewHistory(false); // Reset to active view when opening
        }
    };

    const onItemClick = (item: ITaskCorrespondence) => {
        // Fix: Auto-dismiss when clicked to reduce count
        dismissNotification(null, item.Id);

        setShowCallout(false);
        const isSubtask = !!item.ChildTaskID;
        props.onNotificationClick(
            isSubtask ? item.ChildTaskID! : item.ParentTaskID,
            isSubtask,
            item.ParentTaskID,
            'Correspondence'
        );
    };

    const dismissNotification = (e: React.MouseEvent<any> | null, id: number) => {
        if (e) e.stopPropagation();
        const updated = dismissedIds.indexOf(id) === -1 ? [...dismissedIds, id] : dismissedIds;
        setDismissedIds(updated);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
    };

    const clearAll = async () => {
        const userEmail = await taskService.getCurrentUserEmail();
        await taskService.markAllNotificationsAsRead(userEmail);

        // Refresh local state from local storage (which the service updated)
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) setDismissedIds(JSON.parse(stored));
    };


    // Calculate lists
    const activeNotifications = notifications.filter(n => dismissedIds.indexOf(n.Id) === -1);
    const historyNotifications = notifications.filter(n => dismissedIds.indexOf(n.Id) !== -1);

    // Sort history by date desc (newest first)
    historyNotifications.sort((a, b) => new Date(b.Created).getTime() - new Date(a.Created).getTime());

    const displayList = viewHistory ? historyNotifications : activeNotifications;
    const count = activeNotifications.length;

    return (
        <>
            <div className={styles.notificationContainer} ref={menuButtonRef} onClick={onBellClick}>
                <Bell className={`${styles.bellIcon} ${count > 0 ? styles.bellPulse : ''}`} />
                {count > 0 && (
                    <span className={styles.badge}>
                        {count > 9 ? '9+' : count}
                    </span>
                )}
            </div>

            {showCallout && (
                <Callout
                    className={styles.callout}
                    target={menuButtonRef}
                    onDismiss={() => setShowCallout(false)}
                    directionalHint={DirectionalHint.bottomRightEdge}
                    gapSpace={8}
                    setInitialFocus
                >
                    <FocusTrapZone>
                        <header className={styles.header}>
                            <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                {viewHistory ? 'Notification History' : 'Notifications'}
                                {viewHistory && <span style={{ fontSize: '11px', background: '#e2e8f0', color: '#64748b', padding: '2px 8px', borderRadius: '10px' }}>Read Only</span>}
                            </h3>
                            <div className={styles.headerActions}>
                                {!viewHistory && activeNotifications.length > 0 && (
                                    <button onClick={clearAll} title="Dismiss All">
                                        <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                                            <Trash2 size={14} />
                                            <span>Dismiss All</span>
                                        </Stack>
                                    </button>
                                )}

                                <button
                                    onClick={() => setViewHistory(!viewHistory)}
                                    title={viewHistory ? "View Active" : "View History"}
                                    style={{ color: viewHistory ? '#10b981' : '#64748b' }}
                                >
                                    <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                                        {viewHistory ? <Bell size={14} /> : <Clock size={14} />}
                                        <span>{viewHistory ? "Active" : "History"}</span>
                                    </Stack>
                                </button>

                                <button onClick={() => fetchNotifications()} title="Refresh">
                                    <RefreshCcw size={14} />
                                </button>
                                <X
                                    className={styles.dismissBtn}
                                    size={20}
                                    style={{ cursor: 'pointer', opacity: 1 }}
                                    onClick={() => setShowCallout(false)}
                                />
                            </div>
                        </header>

                        <div className={styles.notificationList}>
                            {displayList.length === 0 ? (
                                <div className={styles.emptyState}>
                                    <CheckCheck className={styles.emptyIcon} size={48} style={{ filter: viewHistory ? 'grayscale(1)' : 'none' }} />
                                    <h4>{viewHistory ? "No history yet" : "You're all caught up!"}</h4>
                                    <p>{viewHistory ? "Dismissed notifications will appear here." : "No new replies or alerts at this time."}</p>
                                </div>
                            ) : (
                                displayList.map((item, index) => (
                                    <div
                                        key={item.Id}
                                        className={styles.notificationItem}
                                        onClick={() => onItemClick(item)}
                                        style={{ opacity: viewHistory ? 0.85 : 1 }}
                                    >
                                        <div className={styles.itemHeader}>
                                            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                                                <MailOpen size={16} color={viewHistory ? "#64748b" : "#d13438"} />
                                                <strong>{item.Title}</strong>
                                            </Stack>
                                            {!viewHistory && (
                                                <X
                                                    size={14}
                                                    className={styles.dismissBtn}
                                                    onClick={(e) => dismissNotification(e, item.Id)}
                                                />
                                            )}
                                        </div>
                                        <div
                                            className={styles.messageBody}
                                            dangerouslySetInnerHTML={{ __html: item.MessageBody }}
                                        />
                                        <div className={styles.metaInfo}>
                                            <span className={styles.author}>
                                                <User size={12} />
                                                {item.Author?.Title || 'System'}
                                            </span>
                                            <span className={styles.date}>
                                                {new Date(item.Created).toLocaleDateString(undefined, {
                                                    month: 'short',
                                                    day: 'numeric',
                                                    hour: '2-digit',
                                                    minute: '2-digit'
                                                })}
                                            </span>
                                        </div>
                                    </div>
                                ))
                            )}
                        </div>
                    </FocusTrapZone>
                </Callout>
            )}
        </>
    );
};
