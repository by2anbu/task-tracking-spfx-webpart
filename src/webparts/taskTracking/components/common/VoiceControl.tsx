
import * as React from 'react';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { keyframes, mergeStyles } from 'office-ui-fabric-react/lib/Styling';

export interface IVoiceControlProps {
    isListening: boolean;
    onToggleListening: () => void;
    disabled?: boolean;
}

// Simple pulsing animation for the active state
const pulseAnimation = keyframes({
    '0%': { transform: 'scale(1)', boxShadow: '0 0 0 0 rgba(204, 0, 0, 0.7)' },
    '70%': { transform: 'scale(1.1)', boxShadow: '0 0 0 10px rgba(204, 0, 0, 0)' },
    '100%': { transform: 'scale(1)', boxShadow: '0 0 0 0 rgba(204, 0, 0, 0)' }
});

const activeMicClass = mergeStyles({
    color: 'red !important',
    animationName: pulseAnimation,
    animationDuration: '2s',
    animationIterationCount: 'infinite'
});

export const VoiceControl: React.FC<IVoiceControlProps> = ({ isListening, onToggleListening, disabled }) => {
    return (
        <TooltipHost content={isListening ? "Listening..." : "Click to speak commands"}>
            <IconButton
                iconProps={{ iconName: 'Microphone' }}
                title="Voice Control"
                ariaLabel="Voice Control"
                disabled={disabled}
                className={isListening ? activeMicClass : ''}
                onClick={onToggleListening}
                styles={{
                    root: {
                        marginLeft: 8,
                        backgroundColor: isListening ? '#ffebee' : 'transparent',
                        borderColor: isListening ? 'red' : 'transparent'
                    }
                }}
            />
        </TooltipHost>
    );
};
