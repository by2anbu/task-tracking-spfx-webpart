import * as React from 'react';
import {
  DetailsList, SelectionMode, IColumn, TextField, PrimaryButton,
  Dropdown, IDropdownOption, Stack,
  DetailsListLayoutMode, ConstrainMode, IGroup,
  IconButton
} from 'office-ui-fabric-react';
import { taskService } from '../../../../services/sp-service';

export interface IConsolidatedReportViewProps {
  userEmail: string;
}

export interface IReportItem {
  key: string;            // Unique ID (M-1 or S-10)
  id: number;
  type: 'Main' | 'Sub';
  title: string;
  parentTitle?: string;
  parentId?: number;
  description: string;
  category: string;
  assignedTo: string;
  assignedToEmail: string;
  status: string;
  dueDate: string;        // ISO string
  completedDate?: string; // ISO string
  createdDate?: string;   // New Field
  overdueDays: number;

  // Hierarchy
  level: number;          // 0 for Main, 1 for Sub, 2 for Sub-Sub...
  hierarchySortKey: string; // "M-001" or "M-001-S-010"
  hasChildren: boolean;
}

export interface IReportFilter {
  startDate?: Date;
  endDate?: Date;
  status?: string;
  category?: string;
  assignedTo?: string;
  searchQuery?: string;
  isOverdue?: boolean;
}

export const ConsolidatedReportView: React.FunctionComponent<IConsolidatedReportViewProps> = (props) => {
  const [allItems, setAllItems] = React.useState<IReportItem[]>([]);
  const [displayedItems, setDisplayedItems] = React.useState<IReportItem[]>([]);
  const [groups, setGroups] = React.useState<IGroup[] | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [error, setError] = React.useState<string | undefined>(undefined);

  // Filters
  const [filters, setFilters] = React.useState<IReportFilter>({});

  // Filter Options
  const [statusOptions, setStatusOptions] = React.useState<IDropdownOption[]>([]);
  const [categoryOptions, setCategoryOptions] = React.useState<IDropdownOption[]>([]);
  const [userOptions, setUserOptions] = React.useState<IDropdownOption[]>([]);

  // Grouping
  const [groupBy, setGroupBy] = React.useState<string>('none');

  // Expansion State (Set of Item Keys that are expanded)
  const [expandedKeys, setExpandedKeys] = React.useState<Set<string>>(new Set());

  React.useEffect(() => {
    loadData();
  }, []);

  React.useEffect(() => {
    applyFiltersAndGrouping();
  }, [allItems, filters, groupBy, expandedKeys]);

  const loadData = async () => {
    setLoading(true);
    try {
      const mainTasks = await taskService.getAllMainTasks();
      const subTasks = await taskService.getAllSubTasks();

      const userSet = new Set<string>();
      const categorySet = new Set<string>();
      const statusSet = new Set<string>();
      const initialExpanded = new Set<string>();

      // 1. Map all items to a generic structure for tree building
      // Interface for Raw Tree Node
      interface ITreeNode {
        id: string; // "M-1" or "S-10"
        type: 'Main' | 'Sub';
        original: any;
        children: ITreeNode[];
        parentId: string | null; // Key of parent
      }

      const nodeMap = new Map<string, ITreeNode>();
      const roots: ITreeNode[] = [];

      // Process Main Tasks
      mainTasks.forEach(m => {
        const key = `M-${m.Id}`;
        const node: ITreeNode = {
          id: key,
          type: 'Main',
          original: m,
          children: [],
          parentId: null
        };
        nodeMap.set(key, node);
        roots.push(node); // Main tasks are always roots

        // Filters metadata
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const assigned: any = m.TaskAssignedTo;
        if (assigned && assigned.Title) userSet.add(assigned.Title);
        if (m.Status) statusSet.add(m.Status);
        if (m.BusinessUnit) categorySet.add(m.BusinessUnit);

        initialExpanded.add(key); // Default expand main tasks
      });

      // Process Subtasks (First Pass - create nodes)
      subTasks.forEach(s => {
        const key = `S-${s.Id}`;
        const node: ITreeNode = {
          id: key,
          type: 'Sub',
          original: s,
          children: [],
          parentId: null // Determined in second pass
        };
        nodeMap.set(key, node);

        // Filters metadata
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const assigned: any = s.TaskAssignedTo;
        if (assigned && assigned.Title) userSet.add(assigned.Title);
        if (s.TaskStatus) statusSet.add(s.TaskStatus);
        if (s.Category) categorySet.add(s.Category);
      });

      // Process Subtasks (Second Pass - Link to parents)
      subTasks.forEach(s => {
        const childKey = `S-${s.Id}`;
        const childNode = nodeMap.get(childKey);
        if (!childNode) return;

        let parentKey = '';
        if (s.ParentSubtaskId) {
          // It is a sub-subtask
          parentKey = `S-${s.ParentSubtaskId}`;
        } else {
          // Direct child of main task
          parentKey = `M-${s.Admin_Job_ID}`;
        }

        const parentNode = nodeMap.get(parentKey);
        if (parentNode) {
          parentNode.children.push(childNode);
          childNode.parentId = parentKey;
        } else {
          // Orphaned subtask? Treat as root or ignore. 
          // Treating as root for safety so it appears in report
          roots.push(childNode);
        }
      });

      // 3. Flatten the Tree recursively
      const processedItems: IReportItem[] = [];

      const pad = (num: number) => {
        let s = num.toString();
        while (s.length < 6) s = "0" + s;
        return s;
      };

      const processNode = (node: ITreeNode, level: number, parentSortKey: string) => {
        const isMain = node.type === 'Main';
        const item = node.original;
        let currentSortKey = '';

        // Extract Data
        let title = '', desc = '', category = '', status = '', assignedTo = '', assignedToEmail = '';
        let dueDate = '', completedDate = '', createdDate = '';
        let overdueDays = 0;
        let numId = 0;

        if (isMain) {
          numId = item.Id;
          currentSortKey = `${parentSortKey}M-${pad(numId)}`;
          title = item.Title;
          desc = item.Task_x0020_Description;
          category = item.BusinessUnit || 'N/A';
          status = item.Status || 'Not Started';
          dueDate = item.TaskDueDate;
          completedDate = item.Task_x0020_End_x0020_Date;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const a: any = item.TaskAssignedTo;
          assignedTo = a ? (a.Title || '') : 'Unassigned';
          assignedToEmail = a ? (a.EMail || '') : '';
          createdDate = item.Created || item.TaskStartDate; // Fallback
        } else {
          numId = item.Id;
          currentSortKey = `${parentSortKey}-S-${pad(numId)}`;
          title = item.Task_Title || item.Title;
          desc = item.Task_Description;
          category = item.Category || 'N/A';
          status = item.TaskStatus || 'Not Started';
          dueDate = item.TaskDueDate;
          completedDate = item.Task_End_Date;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const a: any = item.TaskAssignedTo;
          assignedTo = a ? (a.Title || '') : 'Unassigned';
          assignedToEmail = a ? (a.EMail || '') : '';
          createdDate = item.Task_Created_Date || item.Created;
        }

        overdueDays = calculateOverdue(dueDate, completedDate, status);

        processedItems.push({
          key: node.id,
          id: numId,
          type: node.type,
          title: title,
          parentId: node.parentId ? parseInt(node.parentId.split('-')[1]) : undefined,
          description: desc,
          category: category,
          assignedTo: assignedTo,
          assignedToEmail: assignedToEmail,
          status: status,
          dueDate: dueDate,
          completedDate: completedDate,
          createdDate: createdDate,
          overdueDays: overdueDays,
          level: level,
          hierarchySortKey: currentSortKey,
          hasChildren: node.children.length > 0
        });

        // Recurse children
        // Sort children by ID first? Or existing order?
        // Let's sort children by ID for consistency
        node.children.sort((a, b) => {
          const idA = a.type === 'Main' ? a.original.Id : a.original.Id;
          const idB = b.type === 'Main' ? b.original.Id : b.original.Id;
          return idA - idB;
        });

        node.children.forEach(child => {
          processNode(child, level + 1, currentSortKey);
        });
      };

      // Sort roots by ID
      roots.sort((a, b) => {
        const idA = a.type === 'Main' ? a.original.Id : a.original.Id;
        const idB = b.type === 'Main' ? b.original.Id : b.original.Id;
        return idB - idA; // Newest Main Tasks first? logic says orderBy Created Desc usually, but let's stick to ID Desc
      });

      roots.forEach(root => processNode(root, 0, ''));

      setAllItems(processedItems);
      setExpandedKeys(initialExpanded);

      // Set Options
      const sOpts: IDropdownOption[] = [];
      userSet.forEach(i => { if (i) sOpts.push({ key: i, text: i }); });
      const stOpts: IDropdownOption[] = [];
      statusSet.forEach(i => { if (i) stOpts.push({ key: i, text: i }); });
      const catOpts: IDropdownOption[] = [];
      categorySet.forEach(i => { if (i) catOpts.push({ key: i, text: i }); });

      setUserOptions(sOpts);
      setStatusOptions(stOpts);
      setCategoryOptions(catOpts);

    } catch (e) {
      console.error(e);
      setError("Failed to load task data.");
    } finally {
      setLoading(false);
    }
  };

  const toggleExpand = (key: string) => {
    const newExpanded = new Set<string>();
    expandedKeys.forEach(k => newExpanded.add(k));

    if (newExpanded.has(key)) {
      newExpanded.delete(key);
    } else {
      newExpanded.add(key);
    }
    setExpandedKeys(newExpanded);
  };

  const calculateOverdue = (dueDateStr?: string, completedDateStr?: string, status?: string): number => {
    if (!dueDateStr) return 0;
    const due = new Date(dueDateStr);
    const now = new Date();

    if (status === 'Completed' && completedDateStr) {
      const completed = new Date(completedDateStr);
      if (completed > due) {
        const diffTime = Math.abs(completed.getTime() - due.getTime());
        return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      }
      return 0;
    }

    if (status !== 'Completed' && now > due) {
      const diffTime = Math.abs(now.getTime() - due.getTime());
      return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }

    return 0;
  };

  const applyFiltersAndGrouping = () => {
    let filtered = [...allItems];

    // 1. Filtering (Applied to ALL items flatly first)
    if (filters.searchQuery) {
      const q = filters.searchQuery.toLowerCase();
      filtered = filtered.filter(i =>
        (i.title && i.title.toLowerCase().indexOf(q) !== -1) ||
        (i.description && i.description.toLowerCase().indexOf(q) !== -1) ||
        (i.assignedTo && i.assignedTo.toLowerCase().indexOf(q) !== -1)
      );
    }
    if (filters.status) {
      filtered = filtered.filter(i => i.status === filters.status);
    }
    if (filters.category) {
      filtered = filtered.filter(i => i.category === filters.category);
    }
    if (filters.assignedTo) {
      filtered = filtered.filter(i => i.assignedTo === filters.assignedTo);
    }
    if (filters.isOverdue) {
      filtered = filtered.filter(i => i.overdueDays > 0);
    }
    if (filters.startDate) {
      filtered = filtered.filter(i => i.dueDate && new Date(i.dueDate) >= filters.startDate!);
    }
    if (filters.endDate) {
      filtered = filtered.filter(i => i.dueDate && new Date(i.dueDate) <= filters.endDate!);
    }

    // 2. Sorting & Expand/Collapse Logic
    if (groupBy === 'none') {
      // Hierarchy Sort
      filtered.sort((a, b) => a.hierarchySortKey.localeCompare(b.hierarchySortKey));
      setGroups(undefined);

      // Apply Expand/Collapse Visibility (Multi-level)
      // An item is visible if its parent is expanded.
      // We can check the expandedKeys. The item needs its DIRECT parent to be expanded.
      // But what if the grandparent is collapsed?
      // Actually, we can just check if the parent Key is in expandedKeys.
      // BUT, we need to know the parent key. 
      // In the recursive build, we didn't store parentKey in IReportItem explicitly except for Main/Sub logic.
      // Let's rely on the fact that if a parent is collapsed, its children should be hidden.

      // This is tricky with a flat list filter.
      // Only show items where ALL ancestors are expanded?
      // Or simpler: Iterate and track "currently collapsed level".

      // Let's use the hierarchySortKey to determine ancestry?
      // If "M-1" is collapsed, hide "M-1-..."

      const visibleItems: IReportItem[] = [];
      // Since it's sorted by hierarchy, parents come before children.
      // We can maintain a stack or set of "collapsed prefixes".

      // Actually simpler:
      // For each item, check if its parent is expanded.
      // Root items (level 0) are always visible (if they match filter).
      // Level 1 items are visible if Level 0 parent is expanded.
      // Level 2 items are visible if Level 1 parent is expanded AND Level 0 grandparent was expanded (implied if Level 1 is visible).

      // We need to map item key to its parent key to check efficiently. 
      // Actually we can just scan.

      filtered = filtered.filter(item => {
        if (item.level === 0) return true;

        // Derive parent key from sort key? 
        // Sort Key: M-1-S-10. Parent: M-1.
        // Sort Key: M-1-S-10-S-15. Parent: M-1-S-10.
        // We can reconstruct parent Key from Sort Key by stripping the last segment.
        const parts = item.hierarchySortKey.split('-');
        // M-00001 -> [M, 00001]
        // M-00001-S-00010 -> [M, 00001, S, 00010]

        // Remove last 2 parts (Type and Id)
        if (parts.length < 3) return true; // Should be level 0 or error

        // Construct parent Key
        // We stored Key as "M-1", "S-10". 
        // Sort Key uses padded numbers "M-000001".
        // This mismatch is painful.
        // Let's use the explicit `parentId`? 
        // But `parentId` in IReportItem is just a number. We don't know if parent is Main or Sub easily without lookup.

        // Alternative: Just check `expandedKeys`.
        // If I am S-15, and my parent is S-10. Is S-10 in expandedKeys?
        // If yes, show. If no, hide.
        // This assumes S-10 is visible. If S-10 was hidden because M-1 is collapsed, S-15 matches filter "parent expanded" but effectively hidden?
        // No, if M-1 is collapsed, S-10 is hidden. User cannot expand S-10? 
        // Wait, state "expandedKeys" persists even if hidden.

        // So we need to check RECURSIVELY if all ancestors are expanded.
        // This is slow for each item.

        // Better approach:
        // Since list is sorted by hierarchy, we can simply maintain a "skip until depth X" flag?
        // No, filters scramble the order potentially? 
        // "Hierarchy Sort" sorts by key.
        // Filters might remove a parent but keep a child? 
        // e.g. Search "Subtask". Parent "Main Task" doesn't match.
        // Child "Subtask" matches.
        // Should we show the child? 
        // Usually yes, effectively flattening the list or showing orphaned child.
        // If we show filtered items, expand/collapse logic becomes weird if parent is missing.

        // User Request: "Export to Excel is not loading check" -> likely they want all data.
        // "Multi-level expand collapse".

        // Let's assume if filtered, we SHOW everything matching filter, maybe disabled hierarchy?
        // Or attempt to respect hierarchy.

        // Standard approach:
        // If filter is active (search/etc), Expand All / Ignore Collapse?
        // Or strictly follow collapse.

        // Let's strictly follow collapse logic based on immediate parent.
        // We need to find the specific Parent Key of this item.
        // We can verify this via the `nodeMap` concept, but we lost it.
        // Let's rebuild a quick map or use Sort Key string manipulation.

        // Let's try string manipulation on Sort Key, assuming consistent padding.
        // SortKey: "M-00001-S-00010"
        // Parent SortKey: "M-000001"
        // We need to convert Parent SortKey back to Item Key ("M-1").
        // Remove padding? parseInt.

        const segments = item.hierarchySortKey.match(/[A-Z]+-\d+/g);
        // Matches ["M-000001", "S-000010"]

        if (!segments || segments.length <= 1) return true; // Root

        // Parent Sort Segment is the second to last one? 
        // No, we need the whole chain.
        // Ancestors are:
        // 1. "M-000001"
        // 2. "M-000001-S-000010" -> Parent of S-15

        // Check all ancestors in the chain.
        for (let i = 0; i < segments.length - 1; i++) {
          // Reconstruct Key for this ancestor
          // Segment "M-000001" -> Key "M-1"
          const seg = segments[i]; // "M-000001" or "-S-000010" (Wait, split logic above was regex)
          // My regex above extracts "M-000001".
          // Convert "M-000001" to "M-1".
          const type = seg.split('-')[0];
          const id = parseInt(seg.split('-')[1]);
          const key = `${type}-${id}`;

          if (!expandedKeys.has(key)) return false; // An ancestor is collapsed
        }

        return true;
      });

    } else {
      // Grouping logic
      let groupKeyFn = (item: IReportItem): string => '';

      switch (groupBy) {
        case 'user': groupKeyFn = (i) => i.assignedTo || 'Unassigned'; break;
        case 'status': groupKeyFn = (i) => i.status || 'No Status'; break;
        case 'category': groupKeyFn = (i) => i.category || 'No Category'; break;
        case 'month':
          groupKeyFn = (i) => {
            if (!i.dueDate) return 'No Date';
            const d = new Date(i.dueDate);
            const monthNames = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"];
            return `${monthNames[d.getMonth()]} ${d.getFullYear()}`;
          };
          break;
      }

      filtered.sort((a, b) => groupKeyFn(a).localeCompare(groupKeyFn(b)));

      const newGroups: IGroup[] = [];
      let currentGroupKey = '';
      let currentGroupStartIndex = 0;

      filtered.forEach((item, index) => {
        const key = groupKeyFn(item);
        if (key !== currentGroupKey) {
          if (currentGroupKey !== '') {
            newGroups.push({
              key: currentGroupKey,
              name: currentGroupKey,
              startIndex: currentGroupStartIndex,
              count: index - currentGroupStartIndex,
              isCollapsed: false,
              level: 0
            });
          }
          currentGroupKey = key;
          currentGroupStartIndex = index;
        }
      });

      if (filtered.length > 0) {
        newGroups.push({
          key: currentGroupKey,
          name: currentGroupKey,
          startIndex: currentGroupStartIndex,
          count: filtered.length - currentGroupStartIndex,
          isCollapsed: false,
          level: 0
        });
      }

      setGroups(newGroups);
    }

    setDisplayedItems(filtered);
  };

  const exportToCSV = () => {
    // FIX: Filter allItems again here to ensure we get ALL matching items, ignoring collapse state
    // We can reuse the filter logic or just apply it simply.
    let exportItems = [...allItems];
    // Re-apply filters
    if (filters.searchQuery) {
      const q = filters.searchQuery.toLowerCase();
      exportItems = exportItems.filter(i =>
        (i.title && i.title.toLowerCase().indexOf(q) !== -1) ||
        (i.description && i.description.toLowerCase().indexOf(q) !== -1) ||
        (i.assignedTo && i.assignedTo.toLowerCase().indexOf(q) !== -1)
      );
    }
    if (filters.status) exportItems = exportItems.filter(i => i.status === filters.status);
    if (filters.category) exportItems = exportItems.filter(i => i.category === filters.category);
    if (filters.assignedTo) exportItems = exportItems.filter(i => i.assignedTo === filters.assignedTo);
    if (filters.isOverdue) exportItems = exportItems.filter(i => i.overdueDays > 0);
    if (filters.startDate) exportItems = exportItems.filter(i => i.dueDate && new Date(i.dueDate) >= filters.startDate!);
    if (filters.endDate) exportItems = exportItems.filter(i => i.dueDate && new Date(i.dueDate) <= filters.endDate!);

    // Sort by hierarchy for export
    exportItems.sort((a, b) => a.hierarchySortKey.localeCompare(b.hierarchySortKey));

    const headers = ['Type', 'ID', 'Title', 'Parent ID', 'Status', 'Assigned To', 'Category', 'Created Date', 'Due Date', 'End Date', 'Overdue Days', 'Description'];

    const csvContent = "data:text/csv;charset=utf-8,"
      + headers.join(",") + "\n"
      + exportItems.map(Row => {
        // Indent Title for CSV based on level
        const indent = Array(Row.level + 1).join("    ");
        const cleanTitle = (Row.title || '').replace(/"/g, '""');

        return [
          Row.type,
          Row.id,
          `"${indent}${cleanTitle}"`,
          Row.parentId || '',
          Row.status,
          `"${(Row.assignedTo || '').replace(/"/g, '""')}"`,
          Row.category,
          Row.createdDate ? new Date(Row.createdDate).toLocaleDateString() : '',
          Row.dueDate ? new Date(Row.dueDate).toLocaleDateString() : '',
          Row.completedDate ? new Date(Row.completedDate).toLocaleDateString() : '',
          Row.overdueDays,
          `"${(Row.description || '').replace(/"/g, '""')}"`
        ].join(",");
      }).join("\n");

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "Task_Report_" + new Date().toISOString() + ".csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const getSummaries = () => {
    // Summary also based on displayed items? Or filtered all items?
    // Usually summary should reflect what matches filters.
    // If displayedItems hides children, summary might be misleading.
    // Let's use displayedItems for consistency with visual grid, OR recalculate based on filters.
    // Usually user wants to see "Total Tasks" found. 
    // Let's continue using displayedItems for now.
    const total = displayedItems.length;
    const completed = displayedItems.filter(i => i.status === 'Completed').length;
    const inProgress = displayedItems.filter(i => i.status === 'In Progress').length;
    const overdue = displayedItems.filter(i => i.overdueDays > 0).length;
    return { total, completed, inProgress, overdue };
  };

  const summary = getSummaries();

  const columns: IColumn[] = [
    {
      key: 'type',
      name: 'Type',
      fieldName: 'type',
      minWidth: 50,
      maxWidth: 70,
      onRender: (item: IReportItem) => (
        <span style={{
          fontWeight: 'bold',
          color: item.type === 'Main' ? '#0078d4' : '#666'
        }}>
          {item.type}
        </span>
      )
    },
    {
      key: 'title',
      name: 'Title',
      fieldName: 'title',
      minWidth: 250,
      maxWidth: 400,
      onRender: (item: IReportItem) => {
        if (groupBy !== 'none') {
          return <span>{item.title}</span>;
        }

        // Tree View Render
        const isExpanded = expandedKeys.has(item.key);
        const indent = item.level * 20;

        return (
          <div style={{ display: 'flex', alignItems: 'center', paddingLeft: indent }}>
            {item.hasChildren ? (
              <IconButton
                iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                title={isExpanded ? "Collapse" : "Expand"}
                onClick={() => toggleExpand(item.key)}
                styles={{ root: { height: 24, margin: '0 4px 0 -8px' } }}
              />
            ) : (
              <div style={{ width: 20, display: 'inline-block' }} />
            )}

            <span style={{
              fontWeight: item.level === 0 ? 600 : 400,
              marginLeft: 4
            }}>
              {item.title}
            </span>
          </div>
        );
      }
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'status',
      minWidth: 100,
      onRender: (item: IReportItem) => {
        let color = '#333';
        let bg = 'transparent';
        let fw = 'normal';

        if (item.status === 'Completed') {
          color = '#107c10';
          bg = '#dff6dd';
          fw = '600';
        } else if (item.status === 'In Progress') {
          color = '#005a9e';
          bg = '#eff6fc';
          fw = '600';
        } else if (item.status === 'Not Started') {
          color = '#666';
        }

        // Overlap with Overdue? 
        // User asked: "IF STATUS IS Completed IN Different color", "if task is over due in different color"
        // Usually Overdue implies Incomplete.
        if (item.overdueDays > 0 && item.status !== 'Completed') {
          color = '#a80000';
          bg = '#fde7e9';
          fw = '700';
        }

        return (
          <span style={{
            color: color,
            fontWeight: fw,
            background: bg,
            padding: '4px 8px',
            borderRadius: '4px',
            fontSize: '12px'
          }}>
            {item.status}
          </span>
        );
      }
    },
    { key: 'assigned', name: 'Assigned To', fieldName: 'assignedTo', minWidth: 150 },
    { key: 'category', name: 'Category', fieldName: 'category', minWidth: 100 },
    {
      key: 'created',
      name: 'Created',
      fieldName: 'createdDate',
      minWidth: 100,
      onRender: (item: IReportItem) => item.createdDate ? new Date(item.createdDate).toLocaleDateString() : '-'
    },
    {
      key: 'due',
      name: 'Due Date',
      fieldName: 'dueDate',
      minWidth: 100,
      onRender: (item: IReportItem) => {
        const isOverdue = item.overdueDays > 0 && item.status !== 'Completed';
        return (
          <span style={{ color: isOverdue ? '#a80000' : 'inherit', fontWeight: isOverdue ? 'bold' : 'normal' }}>
            {item.dueDate ? new Date(item.dueDate).toLocaleDateString() : '-'}
          </span>
        );
      }
    },
    {
      key: 'end',
      name: 'End Date',
      fieldName: 'completedDate',
      minWidth: 100,
      onRender: (item: IReportItem) => item.completedDate ? new Date(item.completedDate).toLocaleDateString() : '-'
    },
    {
      key: 'overdue',
      name: 'Overdue (Days)',
      fieldName: 'overdueDays',
      minWidth: 100,
      onRender: (item: IReportItem) => item.overdueDays > 0 ? <span style={{ color: '#a80000', fontWeight: 'bold', background: '#fde7e9', padding: '2px 6px', borderRadius: 2 }}>{item.overdueDays}</span> : '-'
    }
  ];

  return (
    <div style={{ padding: '20px 30px', background: '#faf9f8', minHeight: '100vh' }}>
      {/* Title & Actions */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 25, background: '#fff', padding: '15px 20px', borderRadius: 4, boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
          <h2 style={{ margin: 0, color: '#201f1e', fontWeight: 600 }}>Consolidated Task Report</h2>
        </Stack>

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          {groupBy === 'none' && (
            <PrimaryButton
              text={expandedKeys.size > 0 ? "Collapse All" : "Expand All"}
              onClick={() => {
                if (expandedKeys.size > 0) {
                  setExpandedKeys(new Set());
                } else {
                  // Expand All items with children
                  const allKeys = new Set<string>();
                  allItems.forEach(i => { if (i.hasChildren) allKeys.add(i.key); });
                  setExpandedKeys(allKeys);
                }
              }}
            />
          )}
          <PrimaryButton iconProps={{ iconName: 'ExcelDocument' }} text="Export to CSV" onClick={exportToCSV} />
        </Stack>
      </Stack>

      {/* Summary Cards */}
      <Stack horizontal tokens={{ childrenGap: 20 }} style={{ marginBottom: 20 }}>
        {/* Simple summary based on displayed items (filtered) */}
        <SummaryCard title="Total Tasks" count={summary.total} color="#0078d4" />
        <SummaryCard title="In Progress" count={summary.inProgress} color="#fce100" />
        <SummaryCard title="Completed" count={summary.completed} color="#107c10" />
        <SummaryCard title="Overdue" count={summary.overdue} color="#a80000" />
      </Stack>

      {/* Filters Bar */}
      <div style={{ background: '#f9f9f9', padding: 15, borderRadius: 4, marginBottom: 20 }}>
        <Stack horizontal tokens={{ childrenGap: 15 }} wrap>
          <TextField
            placeholder="Search..."
            iconProps={{ iconName: 'Search' }}
            onChange={(e, v) => setFilters({ ...filters, searchQuery: v })}
          />
          <Dropdown
            placeholder="Filter Status"
            options={[{ key: '', text: 'All Status' }, ...statusOptions]}
            onChange={(e, o) => setFilters({ ...filters, status: o?.key as string })}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            placeholder="Filter Category"
            options={[{ key: '', text: 'All Categories' }, ...categoryOptions]}
            onChange={(e, o) => setFilters({ ...filters, category: o?.key as string })}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            placeholder="Filter User"
            options={[{ key: '', text: 'All Users' }, ...userOptions]}
            onChange={(e, o) => setFilters({ ...filters, assignedTo: o?.key as string })}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            label="Group By:"
            selectedKey={groupBy}
            options={[
              { key: 'none', text: 'Structure (Main -> Sub)' },
              { key: 'user', text: 'Assigned User' },
              { key: 'status', text: 'Status' },
              { key: 'category', text: 'Category' },
              { key: 'month', text: 'Due Month' }
            ]}
            onChange={(e, o) => setGroupBy(o?.key as string)}
            styles={{ dropdown: { width: 180 }, root: { display: 'flex', alignItems: 'center', gap: 10 } }}
          />
          <Stack horizontal verticalAlign="center">
            <span style={{ marginRight: 8 }}>Overdue Only:</span>
            <Dropdown
              selectedKey={filters.isOverdue ? 'yes' : 'no'}
              options={[{ key: 'yes', text: 'Yes' }, { key: 'no', text: 'No' }]}
              onChange={(e, o) => setFilters({ ...filters, isOverdue: o?.key === 'yes' })}
              styles={{ dropdown: { width: 80 } }}
            />
          </Stack>
        </Stack>
      </div>

      {/* Data Grid */}
      <div style={{ border: '1px solid #eee', position: 'relative', height: 'calc(100vh - 220px)', background: '#fff', boxShadow: '0 2px 8px rgba(0,0,0,0.05)' }}>
        {loading && <div style={{ padding: 20 }}>Loading...</div>}
        {!loading && displayedItems.length === 0 && <div style={{ padding: 20 }}>No tasks found matching criteria.</div>}

        {!loading && displayedItems.length > 0 && (
          <DetailsList
            items={displayedItems}
            groups={groups}
            columns={columns}
            selectionMode={SelectionMode.none}
            compact={false} // Use comfortable padding
            layoutMode={DetailsListLayoutMode.justified} // Fill width
            constrainMode={ConstrainMode.unconstrained}
            styles={{
              root: { overflowX: 'hidden' },
              headerWrapper: { selectors: { '.ms-DetailsHeader': { paddingTop: 0 } } }
            }}
          />
        )}
      </div>

    </div>
  );
};

const SummaryCard = ({ title, count, color }: { title: string, count: number, color: string }) => (
  <div style={{
    background: '#fff',
    borderLeft: `5px solid ${color}`,
    boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
    padding: '20px 30px',
    minWidth: 160,
    borderRadius: 4,
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'center'
  }}>
    <div style={{ color: '#666', fontSize: '12px', fontWeight: 600, textTransform: 'uppercase', marginBottom: 8, letterSpacing: '0.5px' }}>{title}</div>
    <div style={{ color: '#333', fontSize: '32px', fontWeight: 700 }}>{count}</div>
  </div>
);
