
import { useState, useEffect, useCallback } from 'react';

export interface ISpeechRecognitionResult {
    transcript: string;
    isListening: boolean;
    startListening: () => void;
    stopListening: () => void;
    hasRecognitionSupport: boolean;
    error: string | undefined;
}

const useSpeechRecognition = (): ISpeechRecognitionResult => {
    const [isListening, setIsListening] = useState(false);
    const [transcript, setTranscript] = useState('');
    const [recognition, setRecognition] = useState<any>(undefined);
    const [error, setError] = useState<string | undefined>(undefined);

    useEffect(() => {
        // Check for browser support
        const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;

        if (SpeechRecognition) {
            console.log("SpeechRecognition API supported."); // DEBUG
            const recognitionInstance = new SpeechRecognition();
            recognitionInstance.continuous = false; // Stop after one sentence for commands
            recognitionInstance.interimResults = false;
            recognitionInstance.lang = 'en-US';

            recognitionInstance.onstart = () => {
                console.log("SpeechRecognition: onstart fired. Listening..."); // DEBUG
                setIsListening(true);
                setError(undefined);
            };

            recognitionInstance.onend = () => {
                console.log("SpeechRecognition: onend fired. Stopped."); // DEBUG
                setIsListening(false);
            };

            recognitionInstance.onresult = (event: any) => {
                const current = event.resultIndex;
                const transcriptText = event.results[current][0].transcript;
                console.log("SpeechRecognition: Result received:", transcriptText); // DEBUG
                setTranscript(transcriptText);
                setError(undefined);
            };

            recognitionInstance.onerror = (event: any) => {
                console.error("SpeechRecognition: Error event:", event.error); // DEBUG
                setIsListening(false);
                setError(event.error); // Set error state
            };

            setRecognition(recognitionInstance);
        } else {
            console.error("SpeechRecognition API NOT supported in this browser."); // DEBUG
            setError("Browser not supported");
        }
    }, []);

    const startListening = useCallback(() => {
        console.log("startListening called. Recognition instance exists:", !!recognition, "IsListening:", isListening); // DEBUG
        if (recognition && !isListening) {
            setTranscript(''); // Clear previous
            setError(undefined);
            try {
                recognition.start();
                console.log("recognition.start() executed."); // DEBUG
            } catch (err: any) {
                console.error("Error starting recognition:", err);
                setError("Start failed: " + err.message);
            }
        }
    }, [recognition, isListening]);

    const stopListening = useCallback(() => {
        console.log("stopListening called."); // DEBUG
        if (recognition && isListening) {
            recognition.stop();
        }
    }, [recognition, isListening]);

    return {
        transcript,
        isListening,
        startListening,
        stopListening,
        hasRecognitionSupport: !!recognition,
        error
    };
};

export default useSpeechRecognition;
