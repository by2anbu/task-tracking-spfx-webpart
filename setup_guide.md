# SharePoint Environment Setup Guide

To successfully run the **Task Tracking & Gantt Orchestrator**, you must create the following SharePoint lists with the exact internal field names specified below.

## 1. Main Tasks List
**List Name:** `Task Tracking System`

| Field Display Name | Internal Name (Critical) | Type |
| --- | --- | --- |
| Title | `Title` | Single line of text |
| Task Description | `Task_x0020_Description` | Multiple lines of text |
| Business Unit | `BusinessUnit` | Choice |
| Reviewers/Approvers | `ReviewerApproverEmails` | Person or Group (Multiple) |
| Start Date | `TaskStartDate` | Date and Time |
| Due Date | `TaskDueDate` | Date and Time |
| Actual End Date | `Task_x0020_End_x0020_Date` | Date and Time |
| Status | `Status` | Choice |
| SMT Year | `SMTYear` | Single line of text |
| SMT Month | `SMTMonth` | Choice |
| Assigned To | `TaskAssignedTo` | Person or Group |
| User Remarks | `UserRemarks` | Multiple lines of text |
| Departments | `Departments` | Single line of text |
| Project | `Project` | Single line of text |

---

## 2. Sub Tasks List
**List Name:** `Task Tracking System User`

| Field Display Name | Internal Name (Critical) | Type |
| --- | --- | --- |
| Title | `Title` | Single line of text |
| Main Task ID | `Admin_Job_ID` | Number |
| Task Title | `Task_Title` | Single line of text |
| Task Description | `Task_Description` | Multiple lines of text |
| Task Due Date | `TaskDueDate` | Date and Time |
| Assigned To | `TaskAssignedTo` | Person or Group |
| Task Status | `TaskStatus` | Choice |
| Created Date | `Task_Created_Date` | Date and Time |
| End Date | `Task_End_Date` | Date and Time |
| Reassign Date | `Task_Reassign_date` | Date and Time |
| User Remarks | `User_Remarks` | Multiple lines of text |
| Category | `Category` | Single line of text |

---

## 3. Correspondence List
**List Name:** `Task Correspondence`

| Field Display Name | Internal Name (Critical) | Type |
| --- | --- | --- |
| Subject | `Title` | Single line of text |
| Message Body | `MessageBody` | Multiple lines of text |
| Sender | `Sender` | Person or Group |
| Parent Task ID | `ParentTaskID` | Number |
| Child Task ID | `ChildTaskID` | Number |
| To Address | `ToAddress` | Single line of text |
| From Address | `FromAddress` | Single line of text |

---

## 4. Workflows List
**List Name:** `Workflows`

| Field Display Name | Internal Name (Critical) | Type |
| --- | --- | --- |
| Title | `Title` | Single line of text |
| Workflow JSON | `WorkflowJson` | Multiple lines of text (Plain text) |
| Is Active | `IsActive` | Yes/No (Boolean) |

---

## 💡 Important Tips for Setup
- **Hidden Internal Names**: When creating fields in SharePoint, create them with the **Internal Name** first (no spaces), then rename them to the Display Name later.
- **Indexing**: For performance with 5000+ items, go to *List Settings > Indexed Columns* and index the `Status`, `TaskAssignedTo`, and `Admin_Job_ID` fields.
