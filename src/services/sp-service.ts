
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/sputilities";
import "@pnp/sp/attachments";
import "@pnp/sp/fields";
import { IMainTask, ISubTask, ITaskCorrespondence, LIST_MAIN_TASKS, LIST_SUB_TASKS, LIST_ADMIN_CONFIG, LIST_TASK_CORRESPONDENCE, LIST_WORKFLOWS, IWorkflow } from "./interfaces";

export class TaskService {
    private _sp: SPFI;
    private _currentUserEmail: string;
    private _context: WebPartContext;

    constructor() { }

    public init(context: WebPartContext): void {
        this._context = context;
        this._sp = spfi().using(SPFx(context));
        this._currentUserEmail = context.pageContext.user.email;
    }

    public async getCurrentUserEmail(): Promise<string> {
        return this._currentUserEmail;
    }

    private _escapeHtml(text: string): string {
        if (!text) return '';
        return text
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    /**
     * Retrieves items from SharePoint using recursive paging.
     * This bypasses the 5000-item threshold for full data fetches.
     * @param query - The PnP query object (must support .getPaged())
     * @param pageSize - Number of items per page (default 2000 for optimal performance)
     * @returns Promise with all items
     */
    private async getAllItemsPaged<T>(
        query: any,
        pageSize: number = 2000
    ): Promise<T[]> {
        try {
            console.log(`[Data] Starting paged retrieval (page size: ${pageSize})...`);
            let results: T[] = [];

            // Get the first page
            let pagedResults = await query.top(pageSize).getPaged();
            results = results.concat(pagedResults.results);

            // Fetch subsequent pages if they exist
            while (pagedResults.hasNext) {
                console.log(`[Data] Fetching next page... current count: ${results.length}`);
                pagedResults = await pagedResults.getNext();
                results = results.concat(pagedResults.results);
            }

            console.log(`[Data] Completed! Total retrieved: ${results.length} items`);
            return results;
        } catch (error) {
            console.error('[Data] Error during paged retrieval:', error);
            // Fallback for queries that might not support getPaged() or for specific threshold issues
            console.log('[Data] Attempting fallback simple fetch...');
            try {
                return await query.top(pageSize)();
            } catch (fallbackError) {
                console.error('[Data] Fallback fetch also failed:', fallbackError);
                throw error;
            }
        }
    }

    public async sendEmail(to: string[], subject: string, body: string, parentTaskId?: number, childTaskId?: number): Promise<void> {
        if (!to || to.length === 0) {
            console.warn("No recipients for email.");
            return;
        }

        try {
            // Use createTaskCorrespondence for consistency
            await this.createTaskCorrespondence({
                parentTaskId: parentTaskId || 0,
                subject: subject,
                messageBody: body,
                toAddress: to.join('; '),
                fromAddress: this._currentUserEmail,
                childTaskId: childTaskId || 0
            });
            console.log(`[Email] Saved via createTaskCorrespondence for ${to.join(', ')}`);
        } catch (e) {
            console.error("[Email] Error saving to Correspondence list:", e);
            throw new Error("Could not queue email in 'Task Correspondence' list.");
        }
    }

    // --- Main Tasks ---

    public async createMainTask(task: Partial<IMainTask>, files?: File[]): Promise<number> {
        const result = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.add(task);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const itemId = (result as any).data?.Id || (result as any).Id;

        // Upload attachments if provided
        if (files && files.length > 0 && itemId) {
            await this.addAttachmentsToItem(LIST_MAIN_TASKS, itemId, files);
        }

        // Get assigned user email for main task
        // The form passes TaskAssignedToId (user ID only), so we need to fetch the email from the created item
        let mainTaskAssigneeEmail = '';

        // First, check if email was provided directly in the task object
        if (task.TaskAssignedTo) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const assignedTo = task.TaskAssignedTo as any;
            if (assignedTo.EMail) {
                mainTaskAssigneeEmail = assignedTo.EMail;
            } else if (typeof assignedTo === 'string') {
                mainTaskAssigneeEmail = assignedTo;
            }
        }

        // If no email found from task object, fetch from the created item (when TaskAssignedToId is used)
        if (!mainTaskAssigneeEmail && itemId) {
            try {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const createdTask = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(itemId)
                    .select("TaskAssignedTo/EMail,TaskAssignedTo/Title")
                    .expand("TaskAssignedTo")();
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                if ((createdTask as any).TaskAssignedTo && (createdTask as any).TaskAssignedTo.EMail) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    mainTaskAssigneeEmail = (createdTask as any).TaskAssignedTo.EMail;
                }
            } catch (e) {
                console.warn('Could not fetch assigned user email for main task', e);
            }
        }

        // Auto-create correspondence record for task creation
        if (itemId) {
            await this.createTaskCorrespondence({
                parentTaskId: itemId,
                parentTaskTitle: task.Title || 'New Main Task',
                parentTaskStatus: task.Status || 'Not Started',
                subject: `Main Task Created: ${task.Title}`,
                messageBody: `A new main task has been created.
Title: ${task.Title}
Description: ${task.Task_x0020_Description || 'N/A'}
Assigned To: ${mainTaskAssigneeEmail || 'N/A'}
Status: ${task.Status || 'Not Started'}`,
                toAddress: mainTaskAssigneeEmail,
                fromAddress: this._currentUserEmail
            });
        }

        return itemId;
    }

    public async getAllMainTasks(): Promise<IMainTask[]> {
        const query = this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items
            .select("*,TaskAssignedTo/Title,TaskAssignedTo/EMail")
            .expand("TaskAssignedTo")
            .orderBy("Created", false);

        const items = await this.getAllItemsPaged<any>(query);
        return items.map(i => this._mapToMainTask(i));
    }

    public async getMainTasksForUser(email: string): Promise<IMainTask[]> {
        // Filter by AssignedTo email
        const query = this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items
            .select("*,TaskAssignedTo/Title,TaskAssignedTo/EMail")
            .expand("TaskAssignedTo")
            .filter(`TaskAssignedTo/EMail eq '${email}'`)
            .orderBy("Created", false);

        const items = await this.getAllItemsPaged<any>(query);
        return items.map(i => this._mapToMainTask(i));
    }

    public async getMainTasksByIds(ids: number[]): Promise<IMainTask[]> {
        if (!ids || ids.length === 0) return [];

        const filter = ids.map(id => `Id eq ${id}`).join(' or ');
        const items = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items
            .select("*,TaskAssignedTo/Title,TaskAssignedTo/EMail")
            .expand("TaskAssignedTo")
            .filter(filter)();

        return items.map(i => this._mapToMainTask(i));
    }

    public async getMainTaskById(id: number): Promise<IMainTask | undefined> {
        try {
            const item = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(id)
                .select("*,TaskAssignedTo/Title,TaskAssignedTo/EMail")
                .expand("TaskAssignedTo")();
            return this._mapToMainTask(item);
        } catch (e) {
            console.error('[TaskService] Error fetching task by Id:', id, e);
            return undefined;
        }
    }

    public async updateMainTask(id: number, task: Partial<IMainTask>): Promise<void> {
        await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(id).update(task);
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private _mapToMainTask(item: any): IMainTask {
        return {
            Id: item.Id,
            Title: item.Title,
            Task_x0020_Description: item.Task_x0020_Description,
            BusinessUnit: item.Business_x0020_Unit, // Map internal name
            ReviewerApproverEmails: item.ReviewerApproverEmails,
            TaskStartDate: item.TaskStartDate,
            TaskDueDate: item.TaskDueDate,
            Task_x0020_End_x0020_Date: item.Task_x0020_End_x0020_Date,
            Status: item.Status,
            SMTYear: item.SMTYear,
            SMTMonth: item.SMTMonth,
            TaskAssignedTo: item.TaskAssignedTo, // Keep raw for now, or normalize
            View: item.View,
            UserRemarks: item.UserRemarks,
            Departments: item.Departments,
            Project: item.Project,
            Created: item.Created,
            // Computed defaults
            PercentComplete: item.PercentComplete || 0
        } as IMainTask;
    }

    // --- Subtasks ---

    public async getSubTasksForMainTask(mainTaskId: number): Promise<ISubTask[]> {
        return this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items
            .select("*,Admin_Job_ID,AttachmentFiles,TaskAssignedTo/Title,TaskAssignedTo/EMail,TaskAssignedTo/Id")
            .expand("AttachmentFiles,TaskAssignedTo")
            .filter(`Admin_Job_ID eq ${mainTaskId}`)();
    }

    public async getSubTasksForUser(email: string): Promise<ISubTask[]> {
        const query = this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items
            .select("*,Admin_Job_ID,AttachmentFiles,TaskAssignedTo/Title,TaskAssignedTo/EMail,TaskAssignedTo/Id")
            .expand("AttachmentFiles,TaskAssignedTo")
            .filter(`TaskAssignedTo/EMail eq '${email}'`);

        return this.getAllItemsPaged<ISubTask>(query);
    }

    public async getSubTasksByMainTaskIds(mainTaskIds: number[]): Promise<ISubTask[]> {
        if (!mainTaskIds || mainTaskIds.length === 0) return [];

        // Batch into chunks to avoid URL length limits
        // Each filter like "Admin_Job_ID eq 123" is ~20 chars, so 100 IDs = ~2000 chars (safe)
        const BATCH_SIZE = 100;
        const allResults: ISubTask[] = [];

        for (let i = 0; i < mainTaskIds.length; i += BATCH_SIZE) {
            const batch = mainTaskIds.slice(i, i + BATCH_SIZE);
            const filter = batch.map(id => `Admin_Job_ID eq ${id}`).join(' or ');

            const query = this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items
                .select("*,Admin_Job_ID,AttachmentFiles,TaskAssignedTo/Title,TaskAssignedTo/EMail")
                .expand("AttachmentFiles,TaskAssignedTo")
                .filter(filter);

            const batchResults = await this.getAllItemsPaged<ISubTask>(query);
            allResults.push(...batchResults);
        }

        console.log(`[Batching] Retrieved ${allResults.length} items across ${Math.ceil(mainTaskIds.length / BATCH_SIZE)} batches`);
        return allResults;
    }

    // Get all subtasks (for SMT to see their created subtasks)
    public async getAllSubTasks(): Promise<ISubTask[]> {
        const query = this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items
            .select("*,Admin_Job_ID,AttachmentFiles,TaskAssignedTo/Title,TaskAssignedTo/EMail")
            .expand("AttachmentFiles,TaskAssignedTo");

        return this.getAllItemsPaged<ISubTask>(query);
    }

    public async getSubTaskById(id: number): Promise<ISubTask | undefined> {
        try {
            return await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(id)
                .select("*,Admin_Job_ID,AttachmentFiles,TaskAssignedTo/Title,TaskAssignedTo/EMail")
                .expand("AttachmentFiles,TaskAssignedTo")();
        } catch (e) {
            console.error('[TaskService] Error fetching subtask:', id, e);
            return undefined;
        }
    }

    public async createSubTask(subTask: Partial<ISubTask>, files?: File[]): Promise<number> {
        const result = await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.add(subTask);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const itemId = (result as any).data?.Id || (result as any).Id;

        // Upload attachments if provided
        if (files && files.length > 0 && itemId) {
            await this.addAttachmentsToItem(LIST_SUB_TASKS, itemId, files);
        }

        // 1. Get assignee email
        // First check if email was provided directly in the subtask object
        let assigneeEmail = '';
        if (subTask.TaskAssignedTo) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            assigneeEmail = (subTask.TaskAssignedTo as any).EMail || '';
        }

        // If no email found from subtask object, fetch from the created item (when TaskAssignedToId is used)
        if (!assigneeEmail && itemId) {
            try {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const createdSubTask = await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(itemId)
                    .select("TaskAssignedTo/EMail,TaskAssignedTo/Title")
                    .expand("TaskAssignedTo")();
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                if ((createdSubTask as any).TaskAssignedTo && (createdSubTask as any).TaskAssignedTo.EMail) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    assigneeEmail = (createdSubTask as any).TaskAssignedTo.EMail;
                }
            } catch (e) {
                console.warn('Could not fetch assigned user email for subtask', e);
            }
        }

        // 2. Get parent task info (Main Task)
        let parentTaskTitle = '';
        let parentTaskStatus = '';

        // If it's a sub-subtask (has ParentSubtaskId), we still link it to the main Admin_Job_ID
        // But the immediate parent is the subtask.
        // For correspondence, we might want to mention the parent subtask title.

        if (subTask.Admin_Job_ID) {
            try {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const parentTask = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(subTask.Admin_Job_ID).select('Title', 'Status')();
                parentTaskTitle = parentTask.Title || '';
                parentTaskStatus = parentTask.Status || 'Not Started';
            } catch (e) {
                console.warn('Could not fetch parent task', e);
            }
        }

        // 3. Auto-create correspondence record for subtask creation
        if (itemId && subTask.Admin_Job_ID) {
            await this.createTaskCorrespondence({
                parentTaskId: subTask.Admin_Job_ID,
                parentTaskTitle: parentTaskTitle,
                parentTaskStatus: parentTaskStatus,
                childTaskId: itemId,
                childTaskTitle: subTask.Task_Title,
                childTaskStatus: subTask.TaskStatus,
                subject: `Subtask Created: ${subTask.Task_Title}`,
                messageBody: `New Subtask Created
Subtask: ${subTask.Task_Title || 'N/A'}
Description: ${subTask.Task_Description || 'N/A'}
Assigned To: ${assigneeEmail || 'N/A'}
Status: ${subTask.TaskStatus || 'Not Started'}

//Please log in to the Task Tracking System to view full details and take action.`
                ,
                toAddress: assigneeEmail,
                fromAddress: this._currentUserEmail
            });
        }

        // 4. Auto-update main task status to "In Progress" when first subtask is created
        if (subTask.Admin_Job_ID) {
            // Check if parent task status is "Not Started" - if so, update to "In Progress"
            if (parentTaskStatus === 'Not Started' || parentTaskStatus === '') {
                try {
                    await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(subTask.Admin_Job_ID).update({
                        Status: 'In Progress'
                    });
                } catch (e) {
                    console.warn('Could not update main task status to In Progress', e);
                }
            }

            // Recalculate Main Task Progress
            await this.updateMainTaskProgress(subTask.Admin_Job_ID);
        }

        return itemId;
    }

    public async updateSubTaskStatus(subTaskId: number, mainTaskId: number, status: string, remarks?: string, forceComplete?: boolean, dueDate?: string): Promise<void> {
        // If completing a parent subtask, check for incomplete children first
        if (status === 'Completed' && !forceComplete) {
            const incompleteChildren = await this.getIncompleteChildTasks(subTaskId, mainTaskId);
            if (incompleteChildren.length > 0) {
                // Throw error with incomplete task details for UI to handle
                const error = new Error('INCOMPLETE_CHILDREN');
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (error as any).incompleteTasks = incompleteChildren;
                throw error;
            }
        }

        // Fetch existing remarks to preserve [WF_NODE:...] tag
        let finalRemarks = remarks || '';
        try {
            const existing = await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(subTaskId).select('User_Remarks')();
            const existingRemarks = existing.User_Remarks || '';
            const match = existingRemarks.match(/\[WF_NODE:(node-\d+)\]/);
            if (match && !finalRemarks.includes(match[0])) {
                finalRemarks = `${finalRemarks} ${match[0]}`.trim();
            }
        } catch (e) {
            console.warn('Could not preserve node tag in remarks', e);
        }

        const updatePayload: any = {
            TaskStatus: status,
            User_Remarks: finalRemarks,
            Task_End_Date: status === 'Completed' ? new Date().toISOString() : null
        };

        if (dueDate) {
            updatePayload.TaskDueDate = dueDate;
        }

        await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(subTaskId).update(updatePayload);

        // If force completing, cascade to all children
        if (status === 'Completed' && forceComplete) {
            await this.forceCompleteAllChildren(subTaskId);
        }

        await this.updateMainTaskProgress(mainTaskId);

        // If completed, trigger workflow engine for SUBTASKS
        if (status === 'Completed') {
            await this.processWorkflowStep(mainTaskId, subTaskId, true);
        }
    }

    /**
     * Update a specific field of a subtask (DueDate or Assignee) and log the change reason
     */
    public async updateSubTaskField(options: {
        subTaskId: number;
        mainTaskId: number;
        title: string;
        field: 'DueDate' | 'Assignee' | 'Description' | 'Status';
        newValue: any;
        remark: string;
    }): Promise<void> {
        const updatePayload: any = {};
        let logMessage = '';
        let targetAssigneeEmail = '';

        if (options.field === 'DueDate') {
            updatePayload.TaskDueDate = options.newValue;
            // Remark passed is now fully formed: "(Changed from Old -> New) Reason"
            logMessage = `Due Date changed. ${options.remark}`;
        } else if (options.field === 'Assignee') {
            updatePayload.TaskAssignedToId = options.newValue; // Expecting userId

            // Fetch new assignee email for notification
            try {
                const user = await this._sp.web.siteUsers.getById(options.newValue)();
                targetAssigneeEmail = user.Email;
                logMessage = `Assignee changed to ${user.Title}. ${options.remark}`;
            } catch (e) {
                console.warn('Could not fetch new assignee details', e);
                logMessage = `Assignee changed. ${options.remark}`;
            }
        } else if (options.field === 'Description') {
            updatePayload.Task_Description = options.newValue;
            logMessage = `Description updated. ${options.remark}`;
        } else if (options.field === 'Status') {
            // For Status validation and logic, use updateSubTaskStatus
            // But we must NOT use updatePayload here if we call updateSubTaskStatus separately, or we must verify consistency
            await this.updateSubTaskStatus(options.subTaskId, options.mainTaskId, options.newValue, options.remark);
            logMessage = `Status updated. ${options.remark}`;
        }

        // Update Subtask if payload has items (Status handles its own update)
        if (Object.keys(updatePayload).length > 0) {
            await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(options.subTaskId).update(updatePayload);
        }

        // Create Correspondence / Notification
        // If assignee didn't change, we notify the current assignee.
        // If it did change, we notify the NEW assignee.
        if (!targetAssigneeEmail) {
            try {
                const subTask = await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(options.subTaskId)
                    .select("TaskAssignedTo/EMail")
                    .expand("TaskAssignedTo")();
                targetAssigneeEmail = (subTask as any).TaskAssignedTo?.EMail;
            } catch (e) {
                console.warn('Could not fetch current assignee email', e);
            }
        }

        if (targetAssigneeEmail) {
            await this.createTaskCorrespondence({
                parentTaskId: options.mainTaskId,
                childTaskId: options.subTaskId,
                childTaskTitle: options.title,
                subject: `Subtask Updated: ${options.field} Change`,
                messageBody: logMessage,
                toAddress: targetAssigneeEmail,
                fromAddress: this._currentUserEmail
            });
        }

        if (options.field !== 'Status') {
            // Status update already calls this internally
            await this.updateMainTaskProgress(options.mainTaskId);
        }
    }

    /**
     * Get all incomplete child tasks (recursive) for a given parent subtask
     * Optimized by fetching all subtasks for the Main Task ID if provided
     */
    public async getIncompleteChildTasks(parentSubtaskId: number, mainTaskId?: number): Promise<ISubTask[]> {
        let allRelatedTasks: ISubTask[] = [];

        if (mainTaskId) {
            // Optimization: Fetch all subtasks for this Main Task to avoid N+1 queries
            allRelatedTasks = await this.getSubTasksForMainTask(mainTaskId);
        } else {
            // Fallback: Fetch all subtasks from list (less efficient, but works if mainTaskId missing)
            // Or we could fetch just children recursively. For safety, let's fetch all.
            allRelatedTasks = await this.getAllSubTasks();
        }

        // Build Adjacency List
        const childrenMap = new Map<number, ISubTask[]>();
        allRelatedTasks.forEach(t => {
            const pId = t.ParentSubtaskId || 0;
            if (!childrenMap.has(pId)) childrenMap.set(pId, []);
            childrenMap.get(pId)!.push(t);
        });

        const incompleteDescendants: ISubTask[] = [];

        // Recursive traversal
        const traverse = (pId: number) => {
            const children = childrenMap.get(pId);
            if (children) {
                children.forEach(child => {
                    // Check if child is incomplete
                    if (child.TaskStatus !== 'Completed') {
                        incompleteDescendants.push(child);
                    }
                    // Continue traversal regardless of status (even if completed, it might have incomplete children?)
                    // Logic check: If B is Completed, but C is Incomplete.
                    // Should we warn? Yes, usually parent matches children. 
                    // If B is Completed, theoretically C should be Completed. 
                    // But if data is inconsistent, traversing deeper ensures we catch everything.
                    traverse(child.Id);
                });
            }
        };

        traverse(parentSubtaskId);

        return incompleteDescendants;
    }

    /**
     * Force complete all child tasks recursively
     */
    public async forceCompleteAllChildren(parentSubtaskId: number): Promise<void> {
        try {
            // Find all child subtasks (where ParentSubtaskId = this subtask's Id)
            const childSubtasks = await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items
                .filter(`ParentSubtaskId eq ${parentSubtaskId}`)();

            // Complete each child subtask
            for (const child of childSubtasks) {
                await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(child.Id).update({
                    TaskStatus: 'Completed',
                    Task_End_Date: new Date().toISOString()
                });

                // Recursively complete grandchildren
                await this.forceCompleteAllChildren(child.Id);
            }
        } catch (e) {
            console.warn('Could not cascade complete to child subtasks', e);
        }
    }

    // --- Logic ---

    public async updateMainTaskProgress(mainTaskId: number): Promise<void> {
        const subTasks = await this.getSubTasksForMainTask(mainTaskId);
        if (subTasks.length === 0) return;

        const completed = subTasks.filter(t => t.TaskStatus === 'Completed').length;
        const total = subTasks.length;
        const percent = total === 0 ? 0 : (completed / total);

        if (percent === 1) {
            await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(mainTaskId).update({
                Status: 'Completed',
                Task_x0020_End_x0020_Date: new Date().toISOString()
            });

            // Get main task info for correspondence record including Author (creator) and Assigned To
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const mainTask = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(mainTaskId)
                .select("Title,TaskAssignedTo/EMail,TaskAssignedTo/Title,Author/EMail,Author/Title")
                .expand("TaskAssignedTo,Author")();

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const assigneeEmail = (mainTask as any).TaskAssignedTo?.EMail || '';
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const creatorEmail = (mainTask as any).Author?.EMail || '';
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const mainTaskTitle = (mainTask as any).Title || 'Main Task';

            // Build ToAddress with both creator and assignee (comma separated, removing duplicates)
            const toEmails: string[] = [];
            if (creatorEmail) toEmails.push(creatorEmail);
            if (assigneeEmail && assigneeEmail !== creatorEmail) toEmails.push(assigneeEmail);
            const toAddress = toEmails.join('; ');

            // Create correspondence record - FromAddress is current user (session user), ToAddress is creator and assignee
            await this.createTaskCorrespondence({
                parentTaskId: mainTaskId,
                parentTaskTitle: mainTaskTitle,
                parentTaskStatus: 'Completed',
                subject: `Main Task Completed: ${mainTaskTitle}`,
                messageBody: `Main task has been automatically completed.<br/>All ${total} subtasks have been completed.<br/>Task: ${mainTaskTitle}`,
                toAddress: toAddress,
                fromAddress: this._currentUserEmail
            });
        }
    }

    /**
     * Update main task status directly (for tasks without subtasks)
     */
    public async updateMainTaskStatus(mainTaskId: number, status: string, userRemarks?: string): Promise<void> {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const updateData: any = {
            Status: status,
            UserRemarks: userRemarks || ''
        };

        // Set end date if completing the task, clear it if reopening
        if (status === 'Completed') {
            updateData.Task_x0020_End_x0020_Date = new Date().toISOString();
        } else {
            // eslint-disable-next-line @rushstack/no-new-null
            updateData.Task_x0020_End_x0020_Date = null;
        }

        // Fetch existing remarks to preserve [WF_NODE:...] tag
        let finalRemarks = userRemarks || '';
        try {
            const existing = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(mainTaskId).select('UserRemarks')();
            const existingRemarks = existing.UserRemarks || '';
            const match = existingRemarks.match(/\[WF_NODE:(node-\d+)\]/);
            if (match && !finalRemarks.includes(match[0])) {
                finalRemarks = `${finalRemarks} ${match[0]}`.trim();
            }
        } catch (e) {
            console.warn('Could not preserve node tag in remarks', e);
        }
        updateData.UserRemarks = finalRemarks;

        await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(mainTaskId).update(updateData);

        // If completed, trigger GLOBAL workflow engine
        if (status === 'Completed') {
            await this.processWorkflowStep(mainTaskId, mainTaskId, false);
        }
    }

    /**
     * Update a specific field of a Main Task (only DueDate supported for now) and log the change reason
     */
    public async updateMainTaskField(options: {
        mainTaskId: number;
        title: string;
        field: 'DueDate' | 'Description' | 'Status';
        newValue: any;
        remark: string;
    }): Promise<void> {
        const updatePayload: any = {};
        let logMessage = '';

        if (options.field === 'DueDate') {
            updatePayload.TaskDueDate = options.newValue;
            logMessage = `Due Date changed. ${options.remark}`;
        } else if (options.field === 'Description') {
            updatePayload.Task_x0020_Description = options.newValue;
            logMessage = `Description updated. ${options.remark}`;
        } else if (options.field === 'Status') {
            // Use dedicated status update method for Main Tasks
            await this.updateMainTaskStatus(options.mainTaskId, options.newValue, options.remark);
            logMessage = `Status updated. ${options.remark}`;
        }

        // Update Main Task if payload not empty (Status handles its own update)
        if (Object.keys(updatePayload).length > 0) {
            await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(options.mainTaskId).update(updatePayload);
        }

        // Fetch Main Task Creator/Assignee for notification
        let recipients: string[] = [];
        try {
            const task = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(options.mainTaskId)
                .select("TaskAssignedTo/EMail,Author/EMail")
                .expand("TaskAssignedTo,Author")();

            if (task.TaskAssignedTo?.EMail) recipients.push(task.TaskAssignedTo.EMail);
            if (task.Author?.EMail) recipients.push(task.Author.EMail);
        } catch (e) {
            console.warn('Could not fetch main task users for notification', e);
        }

        // De-duplicate
        recipients = recipients.filter((item, index) => recipients.indexOf(item) === index);

        // Create Correspondence
        if (recipients.length > 0) {
            await this.createTaskCorrespondence({
                parentTaskId: options.mainTaskId,
                childTaskId: 0, // Main Task Level
                childTaskTitle: options.title,
                subject: `Main Task Updated: ${options.field} Change`,
                messageBody: logMessage,
                toAddress: recipients.join('; '),
                fromAddress: this._currentUserEmail
            });
        }
    }


    public async ensureUser(email: string): Promise<number> {
        const user = await this._sp.web.ensureUser(email);
        console.log("ensureUser result:", user);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if ((user as any).data) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return (user as any).data.Id;
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (user as any).Id;
    }

    public async verifyAdminUser(email: string): Promise<boolean> {
        try {
            const currentUser = await this._sp.web.currentUser();
            const emailLower = (email || currentUser.Email || '').toLowerCase();

            const allAdmins = await this._sp.web.lists.getByTitle(LIST_ADMIN_CONFIG).items
                .select("Id", "Title", "User_Account/Title", "User_Account/EMail", "User_Account/Id")
                .expand("User_Account")();

            const isAdmin = allAdmins.some((item: any) => {
                const userField = item.User_Account;
                if (!userField) return false;

                const checkUser = (u: any) => {
                    if (u.Id && currentUser.Id && u.Id === currentUser.Id) return true;
                    if (u.EMail && emailLower && u.EMail.toLowerCase() === emailLower) return true;
                    if (u.Title && currentUser.Title && u.Title.toLowerCase() === currentUser.Title.toLowerCase()) return true;
                    return false;
                };

                if (Array.isArray(userField)) {
                    return userField.some(u => checkUser(u));
                } else {
                    return checkUser(userField);
                }
            });

            return isAdmin;
        } catch (e) {
            console.error("[AdminCheck] ERROR verifying admin:", e);
            return false;
        }
    }

    // --- Attachments ---

    public async addAttachmentsToItem(listName: string, itemId: number, files: File[]): Promise<void> {
        const item = this._sp.web.lists.getByTitle(listName).items.getById(itemId);
        for (const file of files) {
            // Check illegal chars or size if needed
            await (item as any).attachmentFiles.add(file.name, file);
        }
    }

    public async getAttachments(listName: string, itemId: number): Promise<any[]> {
        const item = this._sp.web.lists.getByTitle(listName).items.getById(itemId);
        const fileInfos = await (item as any).attachmentFiles();
        return fileInfos;
    }

    // --- Users ---

    // --- Users & Fields ---

    public async getSiteUsers(): Promise<any[]> {
        try {
            console.log("[TaskService] Attempting to fetch filtered siteUsers (PrincipalType eq 1)");
            let users = await this._sp.web.siteUsers.filter("PrincipalType eq 1")();
            if (users && users.length > 0) {
                console.log(`[TaskService] Success: Found ${users.length} users with PrincipalType eq 1`);
                return users;
            }

            console.log("[TaskService] No users found with filter. Attempting unfiltered siteUsers fetch");
            users = await this._sp.web.siteUsers();
            if (users && users.length > 0) {
                console.log(`[TaskService] Success: Found ${users.length} users in unfiltered siteUsers`);
                return users;
            }

            console.log("[TaskService] Still no users. Falling back to siteUserInfoList fetch");
            const fallbackUsers = await this._sp.web.siteUserInfoList.items
                .select("Id", "Title", "EMail", "Name")();
            console.log(`[TaskService] siteUserInfoList fetch returned ${fallbackUsers?.length || 0} items`);
            return fallbackUsers || [];
        } catch (e) {
            console.error("[TaskService] Critical error in getSiteUsers", e);
            return [];
        }
    }

    public async getChoiceFieldOptions(listName: string, fieldName: string): Promise<string[]> {
        try {
            const field = await this._sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldName)();
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return (field as any).Choices || [];
        } catch (e) {
            console.error(`Error fetching choices for ${fieldName} in ${listName}`, e);
            return [];
        }
    }

    // --- Correspondence / Email System ---

    public async getTaskCorrespondence(parentId: number, childId?: number): Promise<any[]> {
        let filter = `ParentTaskID eq ${parentId}`;
        if (childId) {
            filter += ` and ChildTaskID eq ${childId}`;
        } else {
            // If we only want Main Task level emails, ensure ChildTaskID is null or 0
            filter += ` and (ChildTaskID eq 0 or ChildTaskID eq null)`;
        }

        return this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items
            .filter(filter)
            .select("Id", "Title", "MessageBody", "Created", "Author/Title", "Author/EMail", "ToAddress", "FromAddress")
            .expand("Author")
            .orderBy("Created", false)(); // Newest first
    }

    public async processWorkflowStep(mainTaskId: number, completedTaskId: number, isSubTask: boolean): Promise<void> {
        try {
            console.log(`[WorkflowEngine] Processing step for ${isSubTask ? 'SubTask' : 'MainTask'}: ${completedTaskId}`);

            // 1. Get task instance to extract node ID and user remarks
            let nodeID = '';
            let workflowTitle = '';
            let latestRemarks = '';
            let currentAssignee = '';
            let currentOwner = '';

            if (isSubTask) {
                const subTask = await this.getSubTaskById(completedTaskId);
                if (!subTask || !subTask.User_Remarks) return;
                const match = subTask.User_Remarks.match(/\[WF_NODE:(node-\d+)\]/);
                if (match) nodeID = match[1];
                latestRemarks = subTask.User_Remarks;
                currentAssignee = (subTask.TaskAssignedTo as any)?.EMail || '';
                // Since subtasks don't explicitly store Author in interface, but SharePoint has it
                try {
                    const full = await this._sp.web.lists.getByTitle(LIST_SUB_TASKS).items.getById(completedTaskId).select('Author/EMail').expand('Author')();
                    currentOwner = (full as any).Author?.EMail || '';
                } catch (e) {
                    // ignore
                }
                workflowTitle = `TASK_WF_${mainTaskId}`;
            } else {
                const mainTask = await this.getMainTaskById(completedTaskId);
                if (!mainTask || !mainTask.UserRemarks) return;
                const match = mainTask.UserRemarks.match(/\[WF_NODE:(node-\d+)\]/);
                if (match) nodeID = match[1];
                latestRemarks = mainTask.UserRemarks;
                currentAssignee = (mainTask.TaskAssignedTo as any)?.EMail || '';
                try {
                    const full = await this._sp.web.lists.getByTitle(LIST_MAIN_TASKS).items.getById(completedTaskId).select('Author/EMail').expand('Author')();
                    currentOwner = (full as any).Author?.EMail || '';
                } catch (e) {
                    // ignore
                }
                // Global active workflow
                const active = await this.getActiveWorkflow();
                if (active) workflowTitle = active.Title;
            }

            if (!nodeID || !workflowTitle) return;

            // 2. Load the workflow
            const workflow = await this.getWorkflowByTitle(workflowTitle);
            if (!workflow || !workflow.WorkflowJson) return;

            const { nodes, edges } = JSON.parse(workflow.WorkflowJson);

            // 3. Find next nodes & handle branching/recursion
            const processNode = async (currentId: string) => {
                const nextEdges = edges.filter((e: any) => e.source === currentId);
                for (const edge of nextEdges) {
                    const nextNode = nodes.find((n: any) => n.id === edge.target);
                    if (!nextNode) continue;

                    console.log(`[WorkflowEngine] Checking node: ${nextNode.data.label} (${nextNode.data.type})`);

                    // --- Condition Processing (Branching) ---
                    if (nextNode.data.type === 'Condition') {
                        // Check for 'If [Keyword]' format
                        const labelMatch = nextNode.data.label.match(/If\s+(.+)/i);
                        if (labelMatch) {
                            const keyword = labelMatch[1].trim().toLowerCase();
                            // If keyword NOT in latest remarks, Skip this branch
                            if (!latestRemarks.toLowerCase().includes(keyword)) {
                                console.log(`[WorkflowEngine] Condition '${keyword}' not met in remarks.`);
                                continue;
                            }
                        }
                        // If logic passes or no keyword, proceed to children
                        await processNode(nextNode.id);
                        continue;
                    }

                    // --- Alert Processing ---
                    if (nextNode.data.type === 'Alert') {
                        const recipients: string[] = [];
                        const notifyWho = nextNode.data.notifyWho || 'assignee';
                        if ((notifyWho === 'assignee' || notifyWho === 'both') && currentAssignee) recipients.push(currentAssignee);
                        if ((notifyWho === 'owner' || notifyWho === 'both') && currentOwner) recipients.push(currentOwner);

                        if (recipients.length > 0) {
                            await this.sendEmail(
                                recipients,
                                nextNode.data.emailSubject || `Alert: ${nextNode.data.label}`,
                                nextNode.data.description || "System workflow alert triggered."
                            );
                        }
                        // Alert is non-blocking, so we also process its children
                        await processNode(nextNode.id);
                        continue;
                    }

                    // --- Task Action ---
                    if (nextNode.data.type === 'Task' && nextNode.data.assignee) {
                        const userId = await this.ensureUser(nextNode.data.assignee);
                        if (isSubTask) {
                            await this.createSubTask({
                                Task_Title: nextNode.data.label,
                                Task_Description: nextNode.data.description,
                                TaskAssignedToId: userId as any,
                                TaskStatus: 'Not Started',
                                Admin_Job_ID: mainTaskId,
                                TaskDueDate: new Date().toISOString(),
                                Category: 'Workflow',
                                User_Remarks: `[WF_NODE:${nextNode.id}]`
                            } as any);
                        } else {
                            await this.createMainTask({
                                Title: nextNode.data.label,
                                Task_x0020_Description: nextNode.data.description,
                                TaskAssignedToId: userId as any,
                                Status: 'Not Started',
                                Project: 'Workflow Task',
                                SMTYear: new Date().getFullYear().toString(),
                                SMTMonth: new Intl.DateTimeFormat('en-US', { month: 'long' }).format(new Date())
                                //, UserRemarks: `[WF_NODE:${nextNode.id}]`
                            } as any);
                        }
                    }
                    // --- Email Action ---
                    else if (nextNode.data.type === 'Email' && nextNode.data.assignee) {
                        await this.sendEmail(
                            [nextNode.data.assignee],
                            nextNode.data.emailSubject || "Workflow Notification",
                            nextNode.data.description || "The previous step has been completed."
                        );
                    }
                }
            };

            await processNode(nodeID);
        } catch (e) {
            console.error('[WorkflowEngine] Error processing step:', e);
        }
    }


    public async getAllCorrespondence(): Promise<any[]> {
        return this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items
            .select("Id", "Title", "MessageBody", "Created", "Author/Title", "Author/EMail", "ToAddress", "FromAddress", "ParentTaskID", "ChildTaskID")
            .expand("Author")
            .orderBy("Created", false)();
    }

    public async getPermissionsAwareCorrespondence(userEmail: string): Promise<any[]> {
        // Security Fix: Server-side filtering to prevent data leak
        // Filter: Author is user OR From is user OR To contains user
        const filter = `Author/EMail eq '${userEmail}' or FromAddress eq '${userEmail}' or substringof('${userEmail}', ToAddress)`;

        return this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items
            .filter(filter)
            .select("Id", "Title", "MessageBody", "Created", "Author/Title", "Author/EMail", "ToAddress", "FromAddress", "ParentTaskID", "ChildTaskID")
            .expand("Author")
            .orderBy("Created", false)();
    }

    /**
     * Get correspondence history for a specific task
     */
    public async getCorrespondenceByTaskId(parentId: number, childId?: number): Promise<any[]> {
        let filter = `ParentTaskID eq ${parentId}`;
        if (childId) {
            filter += ` and ChildTaskID eq ${childId}`;
        } else {
            filter += ` and (ChildTaskID eq 0 or ChildTaskID eq null)`;
        }

        try {
            return await this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items
                .filter(filter)
                .select("Id", "Title", "MessageBody", "Created", "Author/Title", "Author/EMail", "ToAddress", "FromAddress")
                .expand("Author")
                .orderBy("Created", false)(); // Newest first for chat history
        } catch (e) {
            console.error("[Correspondence] Error fetching by task ID:", e);
            return [];
        }
    }

    public async getTaskCorrespondenceMetadata(taskIds: number[], isMainTask: boolean = false, useRollup: boolean = false): Promise<Map<number, { hasCorrespondence: boolean, isReply: boolean }>> {
        if (!taskIds || taskIds.length === 0) return new Map();

        try {
            const BATCH_SIZE = 50;
            const metadata = new Map<number, { hasCorrespondence: boolean, isReply: boolean }>();

            // idField is what we join results by in the Map
            const joinField = isMainTask ? "ParentTaskID" : "ChildTaskID";
            // searchField is what we filter by in the query
            const searchField = isMainTask ? "ParentTaskID" : "ChildTaskID";

            // If we are Main Task but NO rollup, we only want ChildTaskID eq 0/null
            // If we ARE rollup, we want anything with that ParentTaskID
            const extraFilter = (isMainTask && !useRollup) ? " and (ChildTaskID eq 0 or ChildTaskID eq null)" : "";

            for (let i = 0; i < taskIds.length; i += BATCH_SIZE) {
                const batch = taskIds.slice(i, i + BATCH_SIZE);
                const filter = "(" + batch.map(id => `${searchField} eq ${id}`).join(' or ') + ")" + extraFilter;

                const items = await this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items
                    .filter(filter)
                    .select(joinField, "FromAddress", "Created")
                    .orderBy("Created", true)(); // Ascending: Newest wins in Map logic

                items.forEach((item: any) => {
                    const taskId = item[joinField];
                    const isReply = item.FromAddress && item.FromAddress.toLowerCase() !== this._currentUserEmail.toLowerCase();

                    const existing = metadata.get(taskId);
                    // Rollup logic: If ANY message in the rollup is a reply, the whole task shows a bell
                    // If NO rollup, newest message dictates the state (Map.set does this automatically)
                    if (useRollup && existing?.isReply) {
                        return; // Keep existing true
                    }

                    metadata.set(taskId, { hasCorrespondence: true, isReply: !!isReply });
                });
            }
            console.log(`[Correspondence] Metadata: ${taskIds.length} ${isMainTask ? 'Main' : 'Sub'} tasks (Rollup: ${useRollup})`);
            return metadata;
        } catch (e) {
            console.error("[Correspondence] Error fetching metadata:", e);
            return new Map();
        }
    }

    public async getGlobalNotifications(): Promise<ITaskCorrespondence[]> {
        try {
            // Robust approach: Fetch latest 100 correspondence records 
            // and filter in JS to avoid case-sensitivity issues with substringof/ToAddress.
            const items = await this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items
                .select("Id", "Title", "MessageBody", "Created", "Author/Title", "Author/EMail", "ToAddress", "FromAddress", "ParentTaskID", "ChildTaskID")
                .expand("Author")
                .orderBy("Created", false)
                .top(100)();

            const myEmail = this._currentUserEmail.toLowerCase();
            console.log(`[Notifications] Filtering ${items.length} items for ${myEmail}`);

            return items.filter((item: any) => {
                const to = (item.ToAddress || '').toLowerCase();
                const isRecipient = to.indexOf(myEmail) !== -1;
                // Note: We show both sent and received in the notification bell for debugging if needed, 
                // but standard app logic usually restricts to 'isRecipient'.
                // To match user expectation of "seeing their records", we'll stick to 'isRecipient' but ensure casing is perfect.
                return isRecipient;
            });
        } catch (e) {
            console.error("[Notifications] Error in robust fetch:", e);
            return [];
        }
    }

    /**
     * Create correspondence record with task details and view/edit links
     */
    public async createTaskCorrespondence(options: {
        parentTaskId: number;
        parentTaskTitle?: string;
        parentTaskStatus?: string;
        childTaskId?: number;
        childTaskTitle?: string;
        childTaskStatus?: string;
        subject: string;
        messageBody: string;
        toAddress: string;
        fromAddress: string;
        file?: File;
    }): Promise<number> {
        // Generate view/edit link for Power Automate
        // Robust base URL extraction to prevent repeated path segments like /SitePages/Task-Tracking.aspx/SitePages/...
        const currentHref = window.location.href;
        const aspxMatch = currentHref.match(/^(.*\.aspx)/i);
        const baseUrl = aspxMatch ? aspxMatch[1] : currentHref.split('?')[0];

        let viewEditLink = `${baseUrl}?ParentTaskID=${options.parentTaskId}`;
        if (options.childTaskId) {
            viewEditLink += `&ChildTaskID=${options.childTaskId}`;
        }

        // New link for direct viewing without filters/conditions
        const directViewLink = `${baseUrl}?ViewTaskID=${options.parentTaskId}`;

        // Enhance message body with clickable link
        // Security Fix: Sanitize content before creating HTML
        // Also convert newlines to <br/> for display since we are now using plain text with \n
        const safeBody = this._escapeHtml(options.messageBody).replace(/\n/g, '<br/>');
        const enhancedBody = `
            <div style="font-family: inherit;">
                <div style="margin-bottom: 12px; line-height: 1.5;">${safeBody}</div>
            </div>`;

        // Create Log Item in SharePoint List with all fields
        const item: any = {
            Title: options.subject,
            MessageBody: enhancedBody,
            ParentTaskID: options.parentTaskId,
            ToAddress: options.toAddress,
            FromAddress: options.fromAddress
            // Note: parent_task_status and Child_Task_Status fields removed - 
            // they don't exist in the SharePoint list. Add them to the list if needed.
        };

        if (options.childTaskId) {
            item.ChildTaskID = options.childTaskId;
        }

        const addResult = await this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items.add(item);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const itemId = (addResult as any).data?.Id || (addResult as any).Id;

        // Upload Attachment if provided
        if (options.file && itemId) {
            await (this._sp.web.lists.getByTitle(LIST_TASK_CORRESPONDENCE).items.getById(itemId) as any).attachmentFiles.add(options.file.name, options.file);
        }

        // DONE: Power Automate will pick up this new item and send email
        return itemId;
    }

    /**
     * Legacy method - kept for backward compatibility
     */
    public async sendEmailAndLog(
        toEmails: string[],
        subject: string,
        body: string,
        parentId: number,
        childId?: number,
        file?: File
    ): Promise<void> {
        const toAddressString = toEmails.join('; ');
        await this.createTaskCorrespondence({
            parentTaskId: parentId,
            childTaskId: childId,
            subject,
            messageBody: body,
            toAddress: toAddressString,
            fromAddress: this._currentUserEmail,
            file
        });
    }

    // --- Workflows ---

    public async saveWorkflow(title: string, nodes: any[], edges: any[]): Promise<number> {
        try {
            // Serialize data
            const workflowJson = JSON.stringify({ nodes, edges });

            // Create new workflow version
            const result = await this._sp.web.lists.getByTitle(LIST_WORKFLOWS).items.add({
                Title: title,
                WorkflowJson: workflowJson,
                IsActive: true
            });

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            return (result as any).data?.Id || (result as any).Id;
        } catch (e) {
            console.error('[Workflow] Error saving workflow:', e);
            throw e;
        }
    }

    public async getActiveWorkflow(): Promise<IWorkflow | undefined> {
        try {
            // Get the latest active workflow
            const items = await this._sp.web.lists.getByTitle(LIST_WORKFLOWS).items
                .filter("IsActive eq 1") // 1 for Yes
                .orderBy("Created", false)
                .top(1)();

            if (items && items.length > 0) {
                const item = items[0];
                return {
                    Id: item.Id,
                    Title: item.Title,
                    WorkflowJson: item.WorkflowJson,
                    IsActive: item.IsActive
                };
            }
            return undefined;
        } catch (e) {
            console.warn('[Workflow] No active workflow found or error fetching:', e);
            return undefined;
        }
    }

    public async getWorkflowByTitle(title: string): Promise<IWorkflow | undefined> {
        try {
            const items = await this._sp.web.lists.getByTitle(LIST_WORKFLOWS).items
                .filter(`Title eq '${title}'`)
                .orderBy("Created", false)
                .top(1)();

            if (items && items.length > 0) {
                const item = items[0];
                return {
                    Id: item.Id,
                    Title: item.Title,
                    WorkflowJson: item.WorkflowJson,
                    IsActive: item.IsActive
                };
            }
            return undefined;
        } catch (e) {
            console.error('[Workflow] Error fetching workflow by title:', title, e);
            return undefined;
        }
    }

    public async updateOrCreateWorkflow(title: string, nodes: any[], edges: any[]): Promise<void> {
        try {
            const json = JSON.stringify({ nodes, edges });
            const existing = await this.getWorkflowByTitle(title);

            if (existing) {
                await this._sp.web.lists.getByTitle(LIST_WORKFLOWS).items.getById(existing.Id).update({
                    WorkflowJson: json
                });
            } else {
                await this.saveWorkflow(title, nodes, edges);
            }
        } catch (e) {
            console.error('[Workflow] Error updateOrCreateWorkflow:', e);
            throw e;
        }
    }

    /**
     * Mark all notifications as read for current user
     * For now uses local storage but could be extended to SP storage
     */
    public async markAllNotificationsAsRead(userEmail: string): Promise<void> {
        try {
            const notifications = await this.getGlobalNotifications();
            const allIds = notifications.map(n => n.Id);
            const STORAGE_KEY = 'task_tracking_dismissed_notifications';
            const stored = localStorage.getItem(STORAGE_KEY);
            let dismissedIds: number[] = [];
            if (stored) dismissedIds = JSON.parse(stored);

            const uniqueIds = Array.from(new Set([...dismissedIds, ...allIds]));
            localStorage.setItem(STORAGE_KEY, JSON.stringify(uniqueIds));
        } catch (e) {
            console.warn("[TaskService] markAllNotificationsAsRead Error:", e);
        }
    }

    /**
     * Get incomplete parent tasks for dependency checking
     */
    public async getIncompleteParentTasks(taskId: number, mainTaskId: number): Promise<ISubTask[]> {
        const subtasks = await this.getSubTasksForMainTask(mainTaskId);
        const task = subtasks.find(t => t.Id === taskId);
        if (!task || !task.ParentSubtaskId) return [];

        const incompleteParents: ISubTask[] = [];
        const findParent = (pId: number) => {
            const parent = subtasks.find(t => t.Id === pId);
            if (parent && parent.TaskStatus !== 'Completed') {
                incompleteParents.push(parent);
            }
            if (parent && parent.ParentSubtaskId) {
                findParent(parent.ParentSubtaskId);
            }
        };

        findParent(task.ParentSubtaskId);
        return incompleteParents;
    }



}

export const taskService = new TaskService();
