import * as React from 'react';
import styles from './ModernDashboard.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export interface IStatCardProps {
    title: string;
    value: string | number;
    iconName: string;
    trend?: string; // e.g., "+5% vs last week"
    trendDirection?: 'positive' | 'negative' | 'neutral';
    colorClass?: 'iconBlue' | 'iconGreen' | 'iconOrange' | 'iconRed';
    onClick?: () => void;
}

const s = styles as any;

export const StatCard: React.FC<IStatCardProps> = ({
    title,
    value,
    iconName,
    trend,
    trendDirection = 'neutral',
    colorClass = 'iconBlue',
    onClick
}) => {
    return (
        <div className={s.statCard} onClick={onClick} role="button" tabIndex={0}>
            <div className={`${s.statIcon} ${s[colorClass]}`}>
                <Icon iconName={iconName} />
            </div>

            <div style={{ marginTop: 'auto' }}>
                <div className={s.statLabel}>{title}</div>
                <div className={s.statValue}>{value}</div>

                {trend && (
                    <div className={`${s.statTrend} ${s[trendDirection]}`}>
                        <Icon iconName={
                            trendDirection === 'positive' ? 'RiseUp' :
                                trendDirection === 'negative' ? 'FallDown' : 'Remove'
                        } style={{ fontSize: 10 }} />
                        <span>{trend}</span>
                    </div>
                )}
            </div>
        </div>
    );
};
