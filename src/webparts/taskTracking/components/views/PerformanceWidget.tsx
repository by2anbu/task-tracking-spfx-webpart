import * as React from 'react';
import styles from './ModernDashboard.module.scss';
import { Icon } from 'office-ui-fabric-react';

const s = styles as any;

export interface IPerformanceWidgetProps {
    categoryCounts: { [key: string]: number };
    totalTasks: number;
}

// Simple SVG Donut Chart Component
const DonutChart: React.FC<{ data: { label: string; value: number; color: string }[] }> = ({ data }) => {
    const total = data.reduce((acc, curr) => acc + curr.value, 0);
    let cumulativePercent = 0;

    const getCoordinatesForPercent = (percent: number) => {
        const x = Math.cos(2 * Math.PI * percent);
        const y = Math.sin(2 * Math.PI * percent);
        return [x, y];
    };

    return (
        <div style={{ position: 'relative', width: 120, height: 120 }}>
            <svg viewBox="-1 -1 2 2" style={{ transform: 'rotate(-90deg)' }}>
                {data.map((slice, i) => {
                    if (slice.value === 0) return null;
                    const startPercent = cumulativePercent;
                    const slicePercent = slice.value / total;
                    cumulativePercent += slicePercent;

                    const [startX, startY] = getCoordinatesForPercent(startPercent);
                    const [endX, endY] = getCoordinatesForPercent(cumulativePercent);
                    const largeArcFlag = slicePercent > 0.5 ? 1 : 0;

                    const pathData = [
                        `M ${startX} ${startY}`,
                        `A 1 1 0 ${largeArcFlag} 1 ${endX} ${endY}`,
                        `L 0 0`,
                    ].join(' ');

                    return (
                        <path key={i} d={pathData} fill={slice.color} stroke="white" strokeWidth="0.05" />
                    );
                })}
                {/* Inner white circle for donut effect */}
                <circle cx="0" cy="0" r="0.6" fill="white" />
            </svg>
            <div style={{
                position: 'absolute', top: 0, left: 0, width: '100%', height: '100%',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: 24, fontWeight: 'bold', color: '#333'
            }}>
                {total}
            </div>
        </div>
    );
};

export const PerformanceWidget: React.FC<IPerformanceWidgetProps> = ({ categoryCounts }) => {
    const colors = ['#0078d4', '#107c10', '#ff8c00', '#6c757d', '#881798', '#00b294'];

    const chartData = Object.keys(categoryCounts).map((cat, idx) => ({
        label: cat,
        value: categoryCounts[cat],
        color: colors[idx % colors.length]
    })).filter(d => d.value > 0);

    return (
        <div className={s.sectionCard}>
            <div className={s.cardHeader}>
                <h3>Analytics</h3>
                <Icon iconName="PieDouble" style={{ color: '#0078d4' }} />
            </div>
            <div className={s.cardBody} style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 20 }}>
                {chartData.length > 0 ? (
                    <>
                        <DonutChart data={chartData} />
                        <div style={{ width: '100%' }}>
                            {chartData.map(d => (
                                <div key={d.label} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, fontSize: 13 }}>
                                    <div style={{ width: 10, height: 10, borderRadius: '50%', background: d.color }} />
                                    <div style={{ flex: 1, color: '#666' }}>{d.label}</div>
                                    <div style={{ fontWeight: 600 }}>{d.value}</div>
                                </div>
                            ))}
                        </div>
                    </>
                ) : (
                    <div style={{ padding: 20, color: '#999' }}>No data available</div>
                )}
            </div>
        </div>
    );
};
