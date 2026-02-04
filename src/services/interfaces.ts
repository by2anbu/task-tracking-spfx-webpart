
export interface IMainTask {
    Id: number;
    Title: string; // Task Title
    Task_x0020_Description: string;
    BusinessUnit: string;
    ReviewerApproverEmails: any[]; // Person field array
    TaskStartDate: string;
    TaskDueDate: string;
    Task_x0020_End_x0020_Date?: string; // End Date for Main Task
    Status: string;
    SMTYear: string;
    SMTMonth: string;
    TaskAssignedTo: any; // Person field
    TaskAssignedToId?: number; // Support for direct ID assignment
    View: { Description: string, Url: string };
    UserRemarks: string;
    Departments?: string;
    Project?: string;
    // Computed for UI
    PercentComplete?: number;
    Created?: string;
}

export interface ISubTask {
    Id: number;
    Title: string;
    Admin_Job_ID: number; // Foreign Key to Main Task
    Task_Title: string;
    Task_Description: string;
    TaskDueDate: string;
    TaskAssignedTo: any; // Person field
    TaskAssignedToId?: number; // Support for direct ID assignment
    TaskStatus: string;
    Task_Created_Date: string;
    Task_End_Date: string;
    Task_Reassign_date: string;
    User_Remarks: string;
    Category: string;
    // Computed
    Weightage?: number;
    ParentSubtaskId?: number;
    AttachmentFiles?: any[];
}

export interface ITaskCorrespondence {
    Id: number;
    Title: string; // Subject
    MessageBody: string;
    Sender: any; // Person
    ParentTaskID: number;
    ChildTaskID?: number; // Optional/Blank if Main Task level
    ToAddress: string;
    FromAddress: string;
    Created: string;
    Author?: any; // Added for expanding Author field details
}

export interface IWorkflow {
    Id: number;
    Title: string;
    WorkflowJson: string;
    IsActive: boolean;
}

export const LIST_MAIN_TASKS = "Task Tracking System";
export const LIST_SUB_TASKS = "Task Tracking System User";
export const LIST_ADMIN_CONFIG = "Task Tracking System Admin";
export const LIST_TASK_CORRESPONDENCE = "Task Correspondence";
export const LIST_WORKFLOWS = "Workflows";
