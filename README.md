# 📊 Advanced Task Tracking & Gantt Orchestrator for SharePoint

A world-class Project Management solution built on the **SharePoint Framework (SPFx)**. This system delivers a high-performance, hierarchical task management experience with interactive Gantt visualizations and professional reporting.

![Status](https://img.shields.io/badge/Status-Production--Ready-brightgreen)
![SPFx Version](https://img.shields.io/badge/SPFx-v1.17.4-blue)
![React](https://img.shields.io/badge/UI-React--17-61dafb)
![Security](https://img.shields.io/badge/Security-A+-success)

## 🚀 Key Features

### 🌟 Interactive Gantt Orchestration
- **3-Level Hierarchy**: Manage complex projects with a nested structure (Main Task > Subtask > Sub-subtask)
- **Real-time Synchronization**: Instant data updates across all views
- **Visual Progress Tracking**: Beautiful progress bars tracked against "Due Date" vs "Actual End Date"
- **Workflow Designer**: Visual drag-and-drop workflow creation with React Flow

### 📈 Smart Reporting & Exports
- **Automated Excel Export**: Generate comprehensive project reports with a single click
- **Smart Data Mapping**: Automatically calculates task durations and identifies overdue items
- **PNG Capture**: Export the current Gantt view as a high-resolution image for presentations
- **Consolidated Reports**: Multi-employee performance tracking and analytics

### 🔐 Enterprise Architecture
- **Large List Optimization**: Custom paging logic to handle **5000+ items** seamlessly without hitting SharePoint thresholds
- **Correspondence Log**: Built-in audit trail for every task, tracking all comments and status changes
- **Deep Linking**: Navigate directly to specific tasks via URL parameters for instant collaboration
- **Role-Based Access**: Admin and user-level permissions with secure data filtering

### 🛡️ Security Features
- **XSS Protection**: All user-generated content sanitized with DOMPurify
- **SQL Injection Prevention**: Parameterized OData queries with input sanitization
- **Secure Storage**: sessionStorage instead of localStorage for temporary data
- **No Information Disclosure**: Conditional logging (development only)
- **Security Score: A+** - Comprehensive security audit passed

## 🛠️ Tech Stack

- **Frontend**: React 17 + Fluent UI (Office UI Fabric)
- **State Management**: React Component Lifecycle + Optimistic UI Updates
- **Data Layer**: PnP JS (v3) with Recursive Paging
- **Styling**: SCSS Modules with Theme awareness (Light/Dark mode supported)
- **Security**: DOMPurify for XSS protection, sanitized OData queries
- **Deployment**: SPFx Enterprise Package (.sppkg)
- **Additional Libraries**: React Flow (workflow designer), XLSX (Excel export), Lucide React (icons)

## 🏗️ Installation & Setup

### Prerequisites
- Node.js v14 or v16
- SharePoint Online tenant
- SPFx development environment

### 1. Clone the Repository
```bash
git clone https://github.com/YOUR-USERNAME/task-tracking-system.git
cd task-tracking-system
```

### 2. Install Dependencies
```bash
npm install
```

### 3. Configure SharePoint Lists
Follow the instructions in `setup_guide.md` to create the required SharePoint lists:
- Task Tracking System (Main Tasks)
- Task Tracking System User (Subtasks)
- Task Correspondence
- Workflow Designer

### 4. Development Server
```bash
npm run serve
```
*This command allocates 8GB of heap memory to ensure a smooth build process.*

### 5. Build for Production
```bash
gulp clean
gulp build
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` file will be generated in `sharepoint/solution/`

### 6. Deploy to SharePoint
1. Upload the `.sppkg` file to your App Catalog
2. Deploy the solution
3. Add the web part to a SharePoint page

## 📂 Project Structure

```
task-tracking-system/
├── src/
│   ├── services/
│   │   ├── sp-service.ts          # Core SharePoint service with paging logic
│   │   └── interfaces.ts          # TypeScript interfaces
│   ├── utils/
│   │   ├── sanitize.ts            # XSS protection utilities
│   │   └── Logger.ts              # Conditional logging
│   └── webparts/
│       └── taskTracking/
│           └── components/
│               ├── views/          # Main application views
│               │   ├── GanttChartView.tsx
│               │   ├── WorkflowDesigner.tsx
│               │   ├── TaskDetail.tsx
│               │   └── ...
│               └── common/         # Shared components
├── config/                         # SPFx configuration
├── sharepoint/                     # Package output
└── README.md
```

## 🎯 Key Components

- **Gantt Chart View**: Interactive timeline with 3-level task hierarchy
- **Workflow Designer**: Visual workflow creation with drag-and-drop
- **Task Dashboard**: Overview of all tasks with filtering and search
- **Correspondence View**: Hierarchical email/comment tracking
- **Consolidated Reports**: Multi-employee performance analytics
- **Global Notifications**: Real-time notification bell with history

## 🔒 Security

This project follows enterprise security best practices:

✅ **XSS Protection**: All HTML content sanitized with DOMPurify  
✅ **SQL Injection Prevention**: Parameterized OData queries  
✅ **Secure Storage**: sessionStorage for temporary data  
✅ **No Hardcoded Credentials**: All authentication via SharePoint  
✅ **Information Disclosure**: No console.log in production  

For security concerns, please review `final_security_audit.md` in the documentation.

## 📸 Screenshots

> Add your screenshots here after uploading to GitHub

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📝 License

This project is provided **AS IS** under the MIT License.

## 🙏 Acknowledgments

- Built with SharePoint Framework (SPFx)
- UI components from Fluent UI
- Icons from Lucide React
- Workflow visualization with React Flow

---

**⭐ If you find this project useful, please consider giving it a star on GitHub!**

*Created with ❤️ by **Anbarasan**
