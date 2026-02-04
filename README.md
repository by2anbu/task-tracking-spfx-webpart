# 📊 Advanced Task Tracking & Gantt Orchestrator for SharePoint

A world-class Project Management solution built on the **SharePoint Framework (SPFx)**. This system delivers a high-performance, hierarchical task management experience with interactive Gantt visualizations and professional reporting.

![Showcase Placeholder](https://img.shields.io/badge/Status-Production--Ready-brightgreen)
![SPFx Version](https://img.shields.io/badge/SPFx-v1.17.4-blue)
![React](https://img.shields.io/badge/UI-React--17-61dafb)

## 🚀 Key Features

### 🌟 Interactive Gantt Orchestration
- **3-Level Hierarchy**: Manage complex projects with a nested structure (Main Task > Subtask > Sub-subtask).
- **Real-time Synchronization**: Instant data updates across all views.
- **Visual Progress Tracking**: Beautiful progress bars tracked against "Due Date" vs "Actual End Date".

### 📈 Smart Reporting & Exports
- **Automated Excel Export**: Generate comprehensive project bibles with a single click.
- **Smart Data Mapping**: Automatically calculates task durations and identifies overdue items.
- **PNG Capture**: Export the current Gantt view as a high-resolution image for presentations.

### 🔐 Enterprise Architecture
- **Large List Optimization**: Custom paging logic to handle **5000+ items** seamlessly without hitting SharePoint thresholds.
- **Correspondence Log**: A built-in audit trail for every task, tracking all comments and status changes.
- **Deep Linking**: Navigate directly to specific tasks via URL parameters for instant collaboration.

## 🛠️ Tech Stack

- **Frontend**: React 17 + Fluent UI (Office UI Fabric)
- **State Management**: React Component Lifecycle + Optimistic UI Updates
- **Data Layer**: PnP JS (v3) with Recursive Paging
- **Styling**: SCSS Modules with Theme awareness (Light/Dark mode supported)
- **Deployment**: SPFx Enterprise Package (.sppkg)

## 🏗️ Minimal Path to Awesome

1. **Clone the repo**:
   ```bash
   git clone [Your-Repo-URL]
   ```
2. **Setup Dependencies**:
   ```bash
   npm install
   ```
3. **Launch with Optimized Memory**:
   I have included a specialized dev script to handle the large codebase:
   ```bash
   npm run serve
   ```
   *This command allocates 8GB of heap memory to ensure a smooth build process.*

## 📂 Project Structure

- `src/services`: Core logic including the recursive PnP JS paging system.
- `src/webparts/taskTracking/components`: High-performance React components.
- `src/webparts/taskTracking/components/views`: The Gantt engine and Dashboard views.

## 📝 License

This project is provided **AS IS**. (Add your preferred license here, e.g., MIT).

---
*Created with ❤️ by YOUR NAME*
