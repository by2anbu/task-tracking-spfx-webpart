export interface ITaskTrackingProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  // URL parameters for direct task navigation
  parentTaskId?: number;
  childTaskId?: number;
  viewTaskId?: number;
}
