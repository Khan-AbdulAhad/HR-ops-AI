<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Turing Outreach V9</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://code.jquery.com/jquery-3.7.0.js"></script>
    <link href="https://cdn.datatables.net/1.13.6/css/dataTables.tailwindcss.min.css" rel="stylesheet">
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">

    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    colors: {
                        glass: { light: 'rgba(255, 255, 255, 0.9)', dark: 'rgba(17, 24, 39, 0.9)' }
                    }
                }
            }
        }
    </script>
    <style>
        body { 
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; 
            font-size: 15px;
            background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        }
        .dark body { 
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); 
        }
        
        /* Sent row styling */
        tr.sent-row td { 
            background-color: #fef2f2 !important; 
        }
        .dark tr.sent-row td { 
            background-color: #450a0a !important; 
        }
        
        /* Manual sent row styling - different shade */
        tr.manual-sent-row td { 
            background-color: #fef3c7 !important; 
        }
        .dark tr.manual-sent-row td { 
            background-color: #451a03 !important; 
        }
        
        .glass { 
            backdrop-filter: blur(16px); 
            border: 1px solid rgba(255,255,255,0.2); 
            box-shadow: 0 8px 32px rgba(0,0,0,0.08);
            background: rgba(255,255,255,0.95); 
        }
        .dark .glass { 
            background: rgba(30,41,59,0.95); 
            border-color: rgba(255,255,255,0.1); 
            box-shadow: 0 8px 32px rgba(0,0,0,0.4);
        }

        /* ========== SMART CHIPS STYLING ========== */
        .smart-chip {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.625rem 1rem;
            border-radius: 9999px;
            font-size: 0.875rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
            border: 2px solid transparent;
            position: relative;
            user-select: none;
        }
        
        /* Inactive State */
        .smart-chip.inactive {
            background: #f3f4f6;
            color: #6b7280;
            border-color: #e5e7eb;
        }
        .dark .smart-chip.inactive {
            background: #374151;
            color: #9ca3af;
            border-color: #4b5563;
        }
        .smart-chip.inactive:hover {
            background: #e5e7eb;
            border-color: #d1d5db;
            transform: translateY(-1px);
        }
        .dark .smart-chip.inactive:hover {
            background: #4b5563;
            border-color: #6b7280;
        }
        
        /* Active State - Blue Glow */
        .smart-chip.active {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
            color: white;
            border-color: #3b82f6;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.25), 0 4px 12px rgba(59, 130, 246, 0.3);
        }
        .dark .smart-chip.active {
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.35), 0 4px 16px rgba(59, 130, 246, 0.4);
        }
        .smart-chip.active:hover {
            transform: translateY(-1px);
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.3), 0 6px 16px rgba(59, 130, 246, 0.35);
        }
        
        /* Checkmark animation */
        .smart-chip .check-icon {
            width: 0;
            opacity: 0;
            transition: all 0.2s ease;
            overflow: hidden;
        }
        .smart-chip.active .check-icon {
            width: 1rem;
            opacity: 1;
        }

        /* Chip container */
        .chips-container {
            display: flex;
            flex-wrap: wrap;
            gap: 0.625rem;
            padding: 1rem;
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            border-radius: 1rem;
            border: 1px solid #e2e8f0;
        }
        .dark .chips-container {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            border-color: #374151;
        }

        /* Quick Actions for chips */
        .chip-actions {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            padding-left: 0.75rem;
            margin-left: 0.5rem;
            border-left: 2px solid #e5e7eb;
        }
        .dark .chip-actions {
            border-left-color: #4b5563;
        }
        .chip-action-btn {
            font-size: 0.75rem;
            font-weight: 600;
            color: #6b7280;
            padding: 0.375rem 0.75rem;
            border-radius: 0.5rem;
            transition: all 0.2s;
            cursor: pointer;
        }
        .chip-action-btn:hover {
            background: #e5e7eb;
            color: #374151;
        }
        .dark .chip-action-btn:hover {
            background: #4b5563;
            color: #e5e7eb;
        }

        /* Selection count badge on chip */
        .selection-count {
            background: rgba(255,255,255,0.25);
            padding: 0.125rem 0.5rem;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 700;
        }
        .smart-chip.inactive .selection-count {
            background: #e5e7eb;
            color: #6b7280;
        }
        .dark .smart-chip.inactive .selection-count {
            background: #4b5563;
            color: #9ca3af;
        }

        /* ========== GHOST ACTION TOOLBAR ========== */
        .ghost-toolbar {
            display: flex;
            align-items: center;
            gap: 0.25rem;
            padding: 0.5rem;
            background: rgba(255,255,255,0.5);
            backdrop-filter: blur(8px);
            border-radius: 0.75rem;
            border: 1px solid rgba(0,0,0,0.05);
        }
        .dark .ghost-toolbar {
            background: rgba(30,41,59,0.5);
            border-color: rgba(255,255,255,0.05);
        }

        .ghost-btn {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.5rem 0.75rem;
            font-size: 0.8125rem;
            font-weight: 600;
            color: #6b7280;
            border-radius: 0.5rem;
            transition: all 0.2s ease;
            cursor: pointer;
            border: 1px solid transparent;
            background: transparent;
            white-space: nowrap;
        }
        .ghost-btn:hover {
            background: white;
            color: #374151;
            border-color: #e5e7eb;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        .dark .ghost-btn {
            color: #9ca3af;
        }
        .dark .ghost-btn:hover {
            background: #374151;
            color: #f3f4f6;
            border-color: #4b5563;
        }

        /* Ghost button variants */
        .ghost-btn.primary {
            color: #2563eb;
        }
        .ghost-btn.primary:hover {
            background: #eff6ff;
            color: #1d4ed8;
            border-color: #bfdbfe;
        }
        .dark .ghost-btn.primary {
            color: #60a5fa;
        }
        .dark .ghost-btn.primary:hover {
            background: rgba(59, 130, 246, 0.15);
            color: #93c5fd;
            border-color: #3b82f6;
        }

        .ghost-btn.success {
            color: #16a34a;
        }
        .ghost-btn.success:hover {
            background: #f0fdf4;
            color: #15803d;
            border-color: #bbf7d0;
        }

        .ghost-btn.danger {
            color: #dc2626;
        }
        .ghost-btn.danger:hover {
            background: #fef2f2;
            color: #b91c1c;
            border-color: #fecaca;
        }
        
        .ghost-btn.warning {
            color: #d97706;
        }
        .ghost-btn.warning:hover {
            background: #fffbeb;
            color: #b45309;
            border-color: #fcd34d;
        }
        .dark .ghost-btn.warning {
            color: #fbbf24;
        }
        .dark .ghost-btn.warning:hover {
            background: rgba(251, 191, 36, 0.15);
            color: #fcd34d;
            border-color: #f59e0b;
        }

        /* Toolbar divider */
        .toolbar-divider {
            width: 1px;
            height: 1.5rem;
            background: #e5e7eb;
            margin: 0 0.25rem;
        }
        .dark .toolbar-divider {
            background: #4b5563;
        }

        /* ========== MULTI-SELECT FILTER CHIPS ========== */
        .filter-chip {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.5rem 1rem;
            border-radius: 0.625rem;
            font-size: 0.8125rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            border: 2px solid transparent;
            user-select: none;
        }
        
        /* Check icon for multi-select */
        .filter-chip .filter-check {
            width: 0;
            opacity: 0;
            transition: all 0.2s ease;
            overflow: hidden;
            margin-left: -0.25rem;
        }
        .filter-chip.active .filter-check {
            width: 1rem;
            opacity: 1;
            margin-left: 0;
        }
        
        .filter-chip.all {
            background: #f3f4f6;
            color: #374151;
            border-color: #d1d5db;
        }
        .filter-chip.all.active {
            background: #374151;
            color: white;
            border-color: #374151;
        }
        .filter-chip.all:hover:not(.active) {
            background: #e5e7eb;
            border-color: #9ca3af;
        }
        
        .filter-chip.new {
            background: #dcfce7;
            color: #166534;
            border-color: #86efac;
        }
        .filter-chip.new.active {
            background: #16a34a;
            color: white;
            border-color: #16a34a;
            box-shadow: 0 0 0 3px rgba(22, 163, 74, 0.2);
        }
        .filter-chip.new:hover:not(.active) {
            background: #bbf7d0;
            border-color: #4ade80;
        }
        
        .filter-chip.followup {
            background: #fee2e2;
            color: #991b1b;
            border-color: #fca5a5;
        }
        .filter-chip.followup.active {
            background: #dc2626;
            color: white;
            border-color: #dc2626;
            box-shadow: 0 0 0 3px rgba(220, 38, 38, 0.2);
        }
        .filter-chip.followup:hover:not(.active) {
            background: #fecaca;
            border-color: #f87171;
        }
        
        .filter-chip.manual {
            background: #fef3c7;
            color: #92400e;
            border-color: #fcd34d;
        }
        .filter-chip.manual.active {
            background: #d97706;
            color: white;
            border-color: #d97706;
            box-shadow: 0 0 0 3px rgba(217, 119, 6, 0.2);
        }
        .filter-chip.manual:hover:not(.active) {
            background: #fde68a;
            border-color: #fbbf24;
        }
        
        .filter-chip.sub {
            background: #fed7aa;
            color: #c2410c;
            border-color: #fb923c;
        }
        .filter-chip.sub.active {
            background: #ea580c;
            color: white;
            border-color: #ea580c;
            box-shadow: 0 0 0 3px rgba(234, 88, 12, 0.2);
        }
        .filter-chip.sub:hover:not(.active) {
            background: #fdba74;
            border-color: #f97316;
        }
        
        .filter-chip.agency {
            background: #f3e8ff;
            color: #7c3aed;
            border-color: #c4b5fd;
        }
        .filter-chip.agency.active {
            background: #7c3aed;
            color: white;
            border-color: #7c3aed;
            box-shadow: 0 0 0 3px rgba(124, 58, 237, 0.2);
        }
        .filter-chip.agency:hover:not(.active) {
            background: #e9d5ff;
            border-color: #a78bfa;
        }
        
        .filter-chip.direct {
            background: #e0e7ff;
            color: #4338ca;
            border-color: #a5b4fc;
        }
        .filter-chip.direct.active {
            background: #4f46e5;
            color: white;
            border-color: #4f46e5;
            box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.2);
        }
        .filter-chip.direct:hover:not(.active) {
            background: #c7d2fe;
            border-color: #818cf8;
        }

        /* Dark mode filter chips */
        .dark .filter-chip.all {
            background: #374151;
            color: #e5e7eb;
            border-color: #4b5563;
        }
        .dark .filter-chip.all.active {
            background: #6b7280;
            border-color: #6b7280;
        }
        .dark .filter-chip.new {
            background: rgba(22, 163, 74, 0.2);
            color: #86efac;
            border-color: rgba(134, 239, 172, 0.3);
        }
        .dark .filter-chip.followup {
            background: rgba(220, 38, 38, 0.2);
            color: #fca5a5;
            border-color: rgba(252, 165, 165, 0.3);
        }
        .dark .filter-chip.manual {
            background: rgba(217, 119, 6, 0.2);
            color: #fcd34d;
            border-color: rgba(252, 211, 77, 0.3);
        }
        .dark .filter-chip.sub {
            background: rgba(234, 88, 12, 0.2);
            color: #fdba74;
            border-color: rgba(253, 186, 116, 0.3);
        }
        .dark .filter-chip.agency {
            background: rgba(124, 58, 237, 0.2);
            color: #c4b5fd;
            border-color: rgba(196, 181, 253, 0.3);
        }
        .dark .filter-chip.direct {
            background: rgba(79, 70, 229, 0.2);
            color: #a5b4fc;
            border-color: rgba(165, 180, 252, 0.3);
        }
        
        /* Filter group labels */
        .filter-group-label {
            font-size: 0.65rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #9ca3af;
            margin-right: 0.5rem;
        }
        .dark .filter-group-label {
            color: #6b7280;
        }
        
        /* Active filter indicator */
        .active-filters-indicator {
            display: inline-flex;
            align-items: center;
            gap: 0.375rem;
            font-size: 0.75rem;
            font-weight: 600;
            color: #2563eb;
            background: #eff6ff;
            padding: 0.25rem 0.625rem;
            border-radius: 9999px;
            border: 1px solid #bfdbfe;
        }
        .dark .active-filters-indicator {
            background: rgba(59, 130, 246, 0.15);
            color: #60a5fa;
            border-color: #3b82f6;
        }
        
        /* Clear filters button */
        .clear-filters-btn {
            font-size: 0.75rem;
            font-weight: 600;
            color: #6b7280;
            padding: 0.375rem 0.75rem;
            border-radius: 0.5rem;
            transition: all 0.2s;
            cursor: pointer;
            border: 1px solid #e5e7eb;
            background: white;
        }
        .clear-filters-btn:hover {
            background: #fee2e2;
            color: #dc2626;
            border-color: #fecaca;
        }
        .dark .clear-filters-btn {
            background: #374151;
            border-color: #4b5563;
            color: #9ca3af;
        }
        .dark .clear-filters-btn:hover {
            background: rgba(220, 38, 38, 0.2);
            color: #fca5a5;
            border-color: #dc2626;
        }

        /* DataTables wrapper adjustments */
        .dataTables_wrapper { padding: 0; }
        
        #devTable thead th {
            font-weight: 700;
            padding: 1rem !important;
            background: #f9fafb;
            border-bottom: 2px solid #e5e7eb;
        }
        .dark #devTable thead th {
            background: #111827;
            border-bottom: 2px solid #374151;
        }
        
        #devTable tbody td {
            padding: 0.875rem 1rem !important;
            vertical-align: middle;
        }
        
        #devTable tbody tr:hover {
            background-color: #f9fafb !important;
        }
        .dark #devTable tbody tr:hover {
            background-color: #1f2937 !important;
        }

        /* IMPROVED PAGINATION STYLING */
        .dataTables_wrapper .dataTables_paginate {
            padding: 1.5rem 0 !important;
            text-align: center !important;
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            gap: 0.5rem !important;
        }
        
        .dataTables_wrapper .dataTables_paginate .paginate_button {
            padding: 0.625rem 1rem !important; 
            margin: 0 0.25rem !important; 
            border-radius: 0.5rem !important; 
            border: 1px solid #e5e7eb !important; 
            background: white !important; 
            color: #374151 !important;
            font-weight: 600 !important;
            min-width: 2.5rem !important;
            text-align: center !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            transition: all 0.2s !important;
        }
        
        .dataTables_wrapper .dataTables_paginate .paginate_button:hover {
            background: #f3f4f6 !important;
            border-color: #d1d5db !important;
            transform: translateY(-1px);
        }
        
        .dark .dataTables_wrapper .dataTables_paginate .paginate_button { 
            background: #1f2937 !important; 
            border-color: #374151 !important; 
            color: #e5e7eb !important; 
        }
        
        .dark .dataTables_wrapper .dataTables_paginate .paginate_button:hover {
            background: #374151 !important;
            border-color: #4b5563 !important;
        }
        
        .dataTables_wrapper .dataTables_paginate .paginate_button.current { 
            background: #2563eb !important; 
            color: white !important; 
            border-color: #2563eb !important;
            font-weight: 700 !important;
        }
        
        .dataTables_wrapper .dataTables_paginate .paginate_button.disabled {
            opacity: 0.5 !important;
            cursor: not-allowed !important;
        }
        
        /* Status badge */
        .status-badge {
            font-weight: 600;
            display: inline-block;
            line-height: 1.4;
        }
        
        /* Candidate Type Badge Styles */
        .type-badge {
            font-size: 0.7rem;
            text-transform: uppercase;
            padding: 2px 8px;
            border-radius: 999px;
            font-weight: 700;
            margin-left: 8px;
            border: 1px solid;
            display: inline-flex;
            align-items: center;
        }
        
        .type-badge.subcontractor {
            background: #fff7ed;
            color: #c2410c;
            border-color: #fed7aa;
        }
        .dark .type-badge.subcontractor {
            background: rgba(234, 88, 12, 0.2);
            color: #fdba74;
            border-color: #ea580c;
        }
        
        .type-badge.agency {
            background: #faf5ff;
            color: #7c3aed;
            border-color: #e9d5ff;
        }
        .dark .type-badge.agency {
            background: rgba(124, 58, 237, 0.2);
            color: #c4b5fd;
            border-color: #7c3aed;
        }
        
        .type-badge.independent {
            background: #f0fdf4;
            color: #15803d;
            border-color: #bbf7d0;
        }
        .dark .type-badge.independent {
            background: rgba(22, 163, 74, 0.2);
            color: #86efac;
            border-color: #16a34a;
        }
        
        /* RICH TEXT EDITOR */
        .formatting-toolbar {
            display: flex;
            gap: 0.5rem;
            padding: 0.75rem;
            background: #f8fafc;
            border: 2px solid #e5e7eb;
            border-bottom: none;
            border-radius: 0.75rem 0.75rem 0 0;
            flex-wrap: wrap;
            align-items: center;
        }
        .dark .formatting-toolbar {
            background: #1f2937;
            border-color: #4b5563;
        }

        .toolbar-group {
            display: flex;
            gap: 0.25rem;
            padding: 0 0.5rem;
            border-right: 1px solid #e5e7eb;
        }
        .dark .toolbar-group { border-right-color: #4b5563; }

        .toolbar-btn {
            padding: 0.5rem 0.9rem;
            border: 2px solid #e5e7eb;
            background: white;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
            color: #374151;
            font-weight: 600;
            font-size: 0.9rem;
        }
        .toolbar-btn:hover {
            background: #f3f4f6;
            border-color: #3b82f6;
            transform: translateY(-1px);
        }
        .toolbar-btn:active {
            background: #e5e7eb;
            transform: translateY(0);
        }
        .dark .toolbar-btn {
            background: #374151;
            border-color: #4b5563;
            color: #e5e7eb;
        }
        .dark .toolbar-btn:hover {
            background: #4b5563;
            border-color: #60a5fa;
        }

        .rich-text-editor {
            min-height: 300px;
            max-height: 500px;
            overflow-y: auto;
            padding: 1rem;
            border: 2px solid #e5e7eb;
            border-radius: 0 0 0.75rem 0.75rem;
            background: white;
            outline: none;
        }
        .dark .rich-text-editor {
            background: #1f2937;
            border-color: #4b5563;
            color: #e5e7eb;
        }
        .rich-text-editor:focus { border-color: #3b82f6; }

        /* Progress Bar */
        .progress-bar {
            height: 8px;
            background: #e5e7eb;
            border-radius: 999px;
            overflow: hidden;
        }
        .dark .progress-bar { background: #374151; }
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #3b82f6, #2563eb, #1d4ed8);
            transition: width 0.3s ease;
        }

        /* Section Headers */
        .section-header {
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #6b7280;
            margin-bottom: 0.75rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        .dark .section-header {
            color: #9ca3af;
        }

        /* Placeholder hint styling */
        .placeholder-hint {
            font-size: 0.875rem;
            color: #6b7280;
            margin-bottom: 0.5rem;
        }
        .placeholder-hint code {
            background: #e5e7eb;
            padding: 0.125rem 0.375rem;
            border-radius: 0.25rem;
            font-family: monospace;
            color: #2563eb;
            font-weight: 600;
        }
        .dark .placeholder-hint {
            color: #9ca3af;
        }
        .dark .placeholder-hint code {
            background: #374151;
            color: #60a5fa;
        }
        
        /* Template loading state */
        .template-loading {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.75rem;
            background: #fef3c7;
            border: 1px solid #fcd34d;
            border-radius: 0.5rem;
            color: #92400e;
            font-size: 0.875rem;
            margin-top: 0.5rem;
        }
        .dark .template-loading {
            background: rgba(217, 119, 6, 0.2);
            border-color: #d97706;
            color: #fcd34d;
        }
        
        /* Template empty state */
        .template-empty {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.75rem;
            background: #f3f4f6;
            border: 1px solid #e5e7eb;
            border-radius: 0.5rem;
            color: #6b7280;
            font-size: 0.875rem;
            margin-top: 0.5rem;
        }
        .dark .template-empty {
            background: #374151;
            border-color: #4b5563;
            color: #9ca3af;
        }
    </style>
</head>
<body class="p-4 md:p-8 min-h-screen text-gray-800 dark:text-gray-100">

    <div class="max-w-7xl mx-auto">
        <div class="flex justify-between items-center mb-6">
            <div class="flex gap-4">
                 <button onclick="openConfigModal()" class="text-sm bg-white dark:bg-gray-700 px-5 py-2.5 rounded-xl flex items-center gap-2.5 hover:bg-gray-50 dark:hover:bg-gray-600 transition shadow-sm border border-gray-200 dark:border-gray-600 font-medium">
                    <i class="fa-solid fa-database text-blue-600"></i> 
                    <span>Log Sheet Settings</span>
                </button>
                <button onclick="toggleDarkMode()" class="text-sm bg-white dark:bg-gray-700 px-5 py-2.5 rounded-xl flex items-center gap-2.5 hover:bg-gray-50 dark:hover:bg-gray-600 transition shadow-sm border border-gray-200 dark:border-gray-600 font-medium">
                    <i class="fa-solid fa-moon dark:hidden text-gray-600"></i>
                    <i class="fa-solid fa-sun hidden dark:inline text-yellow-400"></i>
                    <span class="dark:hidden">Dark Mode</span>
                    <span class="hidden dark:inline">Light Mode</span>
                </button>
            </div>
            
           <div class="flex items-center gap-3 px-4 py-2 bg-white dark:bg-gray-700 rounded-xl border border-gray-200 dark:border-gray-600 shadow-sm">
                <div class="w-8 h-8 rounded-full bg-blue-100 dark:bg-blue-900 flex items-center justify-center text-blue-600 dark:text-blue-300 font-bold">
                    <i class="fa-solid fa-user"></i>
                </div>
                <div class="flex flex-col">
                    <span class="text-xs text-gray-500 dark:text-gray-400 font-semibold uppercase tracking-wider">Logged in as</span>
                    <span id="currentUserDisplay" class="text-sm font-bold text-gray-900 dark:text-white">Loading...</span>
                </div>
           </div>
        </div>

        <div class="glass rounded-2xl shadow-xl p-8">
            <div class="mb-8 pb-6 border-b border-gray-200 dark:border-gray-700">
                <h1 class="text-3xl md:text-4xl font-bold flex items-center gap-3 text-gray-900 dark:text-white">
                    <span>Outreach Manager</span>
                </h1>
                <p class="text-sm text-gray-500 dark:text-gray-400 mt-2">Fetch and manage developer pipeline data</p>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-5 mb-8">
                <div class="lg:col-span-3">
                    <label class="block text-xs font-bold uppercase mb-2.5 text-gray-700 dark:text-gray-300 tracking-wide">Job ID</label>
                    <input type="number" id="jobId" class="w-full p-3.5 rounded-xl border-2 border-gray-200 dark:border-gray-600 bg-white dark:bg-gray-800 outline-none focus:ring-2 focus:ring-blue-500 transition font-medium" placeholder="51000">
                </div>
                
                <div class="lg:col-span-7">
                    <label class="block text-xs font-bold uppercase mb-2.5 text-gray-700 dark:text-gray-300 tracking-wide">Pipeline Stages</label>
                    <div class="chips-container">
                        <div class="smart-chip inactive" data-value="Interested" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Interested</span>
                        </div>
                        <div class="smart-chip inactive" data-value="Passed VetSmith" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Passed VetSmith</span>
                        </div>
                        <div class="smart-chip inactive" data-value="Pending Review" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Pending Review</span>
                        </div>
                        <div class="smart-chip inactive" data-value="Completed Testing" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Completed Testing</span>
                        </div>
                        <div class="smart-chip inactive" data-value="Developer Backout" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Developer Backout</span>
                        </div>
                        <div class="smart-chip inactive" data-value="On Hold - Onboarding" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>On Hold - Onboarding</span>
                        </div>
                        <div class="smart-chip inactive" data-value="Pending Onboarding" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Pending Onboarding</span>
                        </div>
                        <div class="smart-chip inactive" data-value="Ready for Selection" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Ready for Selection</span>
                        </div>
                        <div class="smart-chip inactive" data-value="Selected for Trial" onclick="toggleChip(this)">
                            <i class="fa-solid fa-check check-icon"></i>
                            <span>Selected for Trial</span>
                        </div>
                        
                        <div class="chip-actions">
                            <span class="chip-action-btn" onclick="selectAllChips()">Select All</span>
                            <span class="chip-action-btn" onclick="clearAllChips()">Clear</span>
                        </div>
                    </div>
                </div>
                
                <div class="lg:col-span-2 flex items-end">
                    <button onclick="fetchData()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3.5 px-6 rounded-xl shadow-lg transition flex items-center justify-center gap-2.5">
                        <i class="fa-solid fa-search"></i> 
                        <span>Fetch Data</span>
                    </button>
                </div>
            </div>

            <div id="smartProgress" class="hidden py-10">
                <div class="w-full max-w-md mx-auto">
                    <div class="flex items-center justify-center mb-4 gap-3">
                        <i class="fa-solid fa-circle-notch fa-spin text-blue-600 text-2xl"></i>
                        <span class="text-gray-700 dark:text-gray-300 font-semibold text-lg" id="progressText">Fetching pipeline data...</span>
                    </div>
                    <div class="progress-bar">
                        <div class="progress-fill" id="progressFill" style="width: 0%"></div>
                    </div>
                </div>
            </div>

            <div id="statusFilterContainer" class="hidden mb-6">
                <div class="section-header">
                    <i class="fa-solid fa-filter text-blue-600"></i>
                    <span>Filter by Pipeline Status</span>
                </div>
                <div id="statusButtons" class="flex flex-wrap gap-2.5"></div>
            </div>

            <div id="tableControls" class="hidden mb-5 space-y-4">
                
                <div class="flex flex-col lg:flex-row gap-4 items-start lg:items-center justify-between">
                    <div class="flex items-center gap-2 w-full lg:w-auto">
                        <div class="relative flex-grow lg:flex-grow-0 lg:w-80">
                            <!-- CHANGE: onkeyup to oninput to catch all changes -->
                            <input type="text" id="tableSearch" oninput="searchTable()" placeholder="Search developers..." class="w-full p-3 pl-10 border-2 border-gray-200 dark:border-gray-600 rounded-xl bg-white dark:bg-gray-800 outline-none focus:ring-2 focus:ring-blue-500 transition font-medium text-sm">
                            <i class="fa-solid fa-magnifying-glass absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400 text-sm"></i>
                        </div>
                        <button onclick="togglePasteFilter()" class="flex-shrink-0 bg-white dark:bg-gray-700 border-2 border-gray-200 dark:border-gray-600 hover:bg-gray-50 dark:hover:bg-gray-600 text-gray-700 dark:text-gray-200 p-3 rounded-xl transition tooltip-btn" title="Filter by Specific Emails/IDs">
                            <i class="fa-solid fa-clipboard-list"></i> Paste & Filter
                        </button>
                    </div>
                    
                    <div class="flex flex-wrap gap-2 items-center">
                        <span class="filter-group-label">Status:</span>
                        <div class="filter-chip all active" data-filter="all" data-group="status" onclick="toggleFilterChip(this)">
                            <i class="fa-solid fa-check filter-check"></i>
                            <i class="fa-solid fa-layer-group"></i>
                            <span>All</span>
                            <span id="allCount" class="text-xs opacity-75">0</span>
                        </div>
                        <div class="filter-chip new" data-filter="new" data-group="status" onclick="toggleFilterChip(this)">
                            <i class="fa-solid fa-check filter-check"></i>
                            <i class="fa-solid fa-sparkles"></i>
                            <span>New</span>
                            <span id="newCount" class="text-xs opacity-75">0</span>
                        </div>
                        <div class="filter-chip followup" data-filter="followup" data-group="status" onclick="toggleFilterChip(this)">
                            <i class="fa-solid fa-check filter-check"></i>
                            <i class="fa-solid fa-reply"></i>
                            <span>Sent</span>
                            <span id="followupCount" class="text-xs opacity-75">0</span>
                        </div>
                        <div class="filter-chip manual" data-filter="manual" data-group="status" onclick="toggleFilterChip(this)">
                            <i class="fa-solid fa-check filter-check"></i>
                            <i class="fa-solid fa-hand"></i>
                            <span>M.Sent</span>
                            <span id="manualCount" class="text-xs opacity-75">0</span>
                        </div>
                        
                        <div class="toolbar-divider"></div>
                        
                        <span class="filter-group-label">Type:</span>
                        <div class="filter-chip sub" data-filter="sub" data-group="type" onclick="toggleFilterChip(this)">
                            <i class="fa-solid fa-check filter-check"></i>
                            <i class="fa-solid fa-user-tie"></i>
                            <span>Sub-Contractor</span>
                            <span id="subCount" class="text-xs opacity-75">0</span>
                        </div>
                        <div class="filter-chip agency" data-filter="agency" data-group="type" onclick="toggleFilterChip(this)">
                            <i class="fa-solid fa-check filter-check"></i>
                            <i class="fa-solid fa-building"></i>
                            <span>Agency</span>
                            <span id="agencyCount" class="text-xs opacity-75">0</span>
                        </div>
                        <div class="filter-chip direct" data-filter="direct" data-group="type" onclick="toggleFilterChip(this)">
                            <i class="fa-solid fa-check filter-check"></i>
                            <i class="fa-solid fa-user"></i>
                            <span>Direct</span>
                            <span id="nonSubCount" class="text-xs opacity-75">0</span>
                        </div>
                        
                        <div class="toolbar-divider"></div>
                        
                        <div id="activeFiltersInfo" class="hidden items-center gap-2">
                            <span class="active-filters-indicator">
                                <i class="fa-solid fa-filter"></i>
                                <span id="activeFilterCount">0</span> filters
                            </span>
                            <button class="clear-filters-btn" onclick="clearAllFilters()">
                                <i class="fa-solid fa-xmark mr-1"></i>Clear
                            </button>
                        </div>
                    </div>
                </div>

                <!-- PASTE FILTER AREA (Hidden by default) -->
                <div id="pasteFilterArea" class="hidden bg-gray-50 dark:bg-gray-800 p-4 rounded-xl border border-gray-200 dark:border-gray-600 shadow-inner">
                    <div class="flex flex-col gap-3">
                        <label class="text-sm font-bold text-gray-700 dark:text-gray-300">Paste Comma Separated Emails or Developer IDs:</label>
                        <textarea id="pasteFilterInput" class="w-full h-24 p-3 rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 text-sm font-mono focus:ring-2 focus:ring-blue-500 outline-none" placeholder="john@example.com, jane@example.com&#10;OR&#10;12345, 67890"></textarea>
                        <div class="flex gap-3">
                            <button onclick="applyPasteFilter()" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-bold shadow transition">
                                <i class="fa-solid fa-filter mr-1"></i> Filter List
                            </button>
                            <button onclick="clearPasteFilter()" class="bg-white dark:bg-gray-700 hover:bg-gray-100 dark:hover:bg-gray-600 text-gray-700 dark:text-gray-200 border border-gray-300 dark:border-gray-500 px-4 py-2 rounded-lg text-sm font-medium transition">
                                Clear List
                            </button>
                            <button onclick="togglePasteFilter()" class="ml-auto text-gray-500 hover:text-gray-700 dark:text-gray-400 text-sm">Close</button>
                        </div>
                    </div>
                </div>

                <div class="flex flex-wrap items-center justify-between gap-3">
                    <div class="ghost-toolbar">
                        <button onclick="selectAllResults()" class="ghost-btn primary">
                            <i class="fa-solid fa-check-double"></i>
                            <span>Select All</span>
                            <span id="selectAllCount" class="text-xs opacity-60">(0)</span>
                        </button>
                        <button onclick="selectCurrentPage()" class="ghost-btn">
                            <i class="fa-solid fa-check"></i>
                            <span>Page Only</span>
                        </button>
                        <button onclick="deselectAll()" class="ghost-btn">
                            <i class="fa-solid fa-xmark"></i>
                            <span>Clear</span>
                        </button>
                        
                        <div class="toolbar-divider"></div>
                        
                        <button onclick="markAsManualSent()" class="ghost-btn warning">
                            <i class="fa-solid fa-hand"></i>
                            <span>Mark Sent</span>
                        </button>
                        
                        <div class="toolbar-divider"></div>
                        
                        <button onclick="copyToClipboard()" class="ghost-btn">
                            <i class="fa-regular fa-copy"></i>
                            <span>Copy</span>
                        </button>
                        <button onclick="downloadCSV()" class="ghost-btn success">
                            <i class="fa-solid fa-download"></i>
                            <span>Export CSV</span>
                        </button>
                    </div>
                    
                    <div id="selectionIndicator" class="hidden items-center gap-2 text-sm font-semibold text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/30 px-3 py-1.5 rounded-lg">
                        <i class="fa-solid fa-circle-check"></i>
                        <span><span id="quickSelectedCount">0</span> selected</span>
                    </div>
                </div>
            </div>
            
            <div id="tableContainer" class="hidden overflow-hidden rounded-xl border-2 border-gray-200 dark:border-gray-700 shadow-sm">
                <table id="devTable" class="w-full text-left">
                    <thead class="bg-gray-50 dark:bg-gray-900 uppercase text-sm font-bold text-gray-700 dark:text-gray-300">
                        <tr>
                            <th class="px-4 py-3 w-12 text-center">
                                <input type="checkbox" id="masterCheckbox" class="w-4 h-4 cursor-pointer" onclick="toggleMasterCheckbox()" title="Select all on current page">
                            </th>
                            <th class="px-4 py-3">Developer ID</th>
                            <th class="px-4 py-3">Candidate Name</th>
                            <th class="px-4 py-3">Email Address</th>
                            <th class="px-4 py-3">Pipeline Status</th>
                            <th class="px-4 py-3">Type</th>
                            <th class="px-4 py-3">Outreach History</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-gray-200 dark:divide-gray-700 bg-white dark:bg-gray-800"></tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="bulkActionBar" class="fixed bottom-8 left-1/2 transform -translate-x-1/2 bg-gradient-to-r from-gray-900 to-gray-800 text-white px-8 py-4 rounded-2xl shadow-2xl flex items-center gap-6 transition-all duration-300 translate-y-[200%] z-40 border border-gray-700">
        <div class="flex items-center gap-2">
            <i class="fa-solid fa-check-circle text-green-400"></i>
            <span class="font-bold text-lg"><span id="selectedCount">0</span> selected</span>
        </div>
        <div class="h-8 w-px bg-gray-600"></div>
        <button onclick="openDraftModal()" class="flex items-center gap-2.5 hover:text-blue-300 font-bold text-base px-3 py-1 hover:bg-white/10 rounded-lg transition">
            <i class="fa-solid fa-paper-plane"></i> 
            <span>Draft Email Campaign</span>
        </button>
        <div class="h-8 w-px bg-gray-600"></div>
        <button onclick="markAsManualSent()" class="flex items-center gap-2.5 hover:text-amber-300 font-bold text-base px-3 py-1 hover:bg-white/10 rounded-lg transition">
            <i class="fa-solid fa-hand"></i> 
            <span>Mark as Sent</span>
        </button>
    </div>

    <!-- CONFIG MODAL -->
    <div id="configModal" class="fixed inset-0 z-50 hidden bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl p-8 w-full max-w-md border border-gray-200 dark:border-gray-700">
            <div class="mb-6">
                <h3 class="text-2xl font-bold text-gray-900 dark:text-white">Log Sheet Configuration</h3>
                <p class="text-sm text-gray-500 dark:text-gray-400 mt-2">Enter your Google Sheets URL for email logging & templates</p>
            </div>
            <input type="text" id="sheetUrl" class="w-full p-4 border-2 border-gray-200 dark:border-gray-600 rounded-xl mb-6 dark:bg-gray-700 outline-none focus:ring-2 focus:ring-blue-500 transition font-medium" placeholder="https://docs.google.com/spreadsheets/d/...">
            <div class="flex justify-end gap-3">
                <button onclick="closeConfigModal()" class="px-6 py-2.5 text-gray-600 dark:text-gray-300 hover:bg-gray-100 rounded-lg font-semibold">Cancel</button>
                <button onclick="saveConfig()" class="px-6 py-2.5 bg-blue-600 text-white rounded-lg font-bold hover:bg-blue-700">Save Configuration</button>
            </div>
        </div>
    </div>

    <!-- EMAIL MODAL -->
    <div id="emailModal" class="fixed inset-0 z-50 hidden bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl w-full max-w-4xl h-[95vh] flex flex-col border border-gray-200 dark:border-gray-700">
            <div class="p-6 border-b-2 dark:border-gray-700 flex justify-between items-center bg-gray-50 dark:bg-gray-900/50 rounded-t-2xl">
                <div>
                    <h2 class="text-2xl font-bold text-gray-900 dark:text-white">Draft Email Campaign</h2>
                    <p class="text-sm text-gray-500 dark:text-gray-400 mt-1">Compose and send bulk emails with formatting</p>
                </div>
                <button onclick="document.getElementById('emailModal').classList.add('hidden')" class="text-gray-400 hover:text-gray-600 transition text-xl"><i class="fa-solid fa-xmark"></i></button>
            </div>
            
            <div class="p-6 overflow-y-auto flex-1 space-y-5">
                <div class="bg-gradient-to-r from-blue-50 to-indigo-50 dark:from-blue-900/20 dark:to-indigo-900/20 p-5 rounded-xl border-2 border-blue-300 dark:border-blue-700 shadow-sm">
                    <label class="block text-xs font-bold uppercase mb-3 text-blue-800 dark:text-blue-300 tracking-wide flex items-center gap-2">
                        <i class="fa-solid fa-wand-magic-sparkles text-blue-600"></i> 
                        <span>Load Saved Template</span>
                        <button onclick="refreshTemplates()" class="ml-auto text-blue-600 hover:text-blue-800 dark:text-blue-400 dark:hover:text-blue-300" title="Refresh templates">
                            <i class="fa-solid fa-arrows-rotate"></i>
                        </button>
                    </label>
                    <select id="templateLoader" onchange="loadTemplate(this.value)" class="w-full p-3.5 border-2 border-blue-400 dark:border-blue-600 rounded-lg bg-white dark:bg-gray-800 outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-base font-medium shadow-sm cursor-pointer transition hover:border-blue-500">
                        <option value="">-- Select a saved template --</option>
                    </select>
                    <div id="templateStatus" class="hidden"></div>
                </div>

                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-bold uppercase mb-2 text-gray-700 dark:text-gray-300 tracking-wide">Sender Name</label>
                        <input type="text" id="senderName" class="w-full p-3.5 border-2 border-gray-200 dark:border-gray-600 rounded-xl dark:bg-gray-700 outline-none focus:ring-2 focus:ring-blue-500 font-medium" placeholder="Your Name">
                    </div>
                    <div>
                        <label class="block text-xs font-bold uppercase mb-2 text-gray-700 dark:text-gray-300 tracking-wide">Email Subject</label>
                        <input type="text" id="emailSubject" class="w-full p-3.5 border-2 border-gray-200 dark:border-gray-600 rounded-xl dark:bg-gray-700 outline-none focus:ring-2 focus:ring-blue-500 font-medium" placeholder="Subject line">
                    </div>
                </div>

                <div>
                    <label class="block text-xs font-bold uppercase mb-2 text-gray-700 dark:text-gray-300 tracking-wide">Email Body</label>
                    <p class="placeholder-hint">Use <code>{{name}}</code> as a placeholder for the candidate's name</p>
                    <div class="formatting-toolbar">
                        <div class="toolbar-group">
                            <button type="button" class="toolbar-btn" onclick="formatText('bold')"><strong>B</strong></button>
                            <button type="button" class="toolbar-btn" onclick="formatText('italic')"><em>I</em></button>
                            <button type="button" class="toolbar-btn" onclick="formatText('underline')"><u>U</u></button>
                        </div>
                        <div class="toolbar-group">
                            <button type="button" class="toolbar-btn" onclick="formatText('h2')">H2</button>
                            <button type="button" class="toolbar-btn" onclick="formatText('insertUnorderedList')"><i class="fa-solid fa-list-ul"></i></button>
                        </div>
                        <div class="toolbar-group">
                            <button type="button" class="toolbar-btn" onclick="formatText('removeFormat')"><i class="fa-solid fa-eraser"></i></button>
                        </div>
                    </div>
                    <div id="emailBody" class="rich-text-editor" contenteditable="true" oninput="handleEditorInput()"></div>
                    <input type="hidden" id="htmlBody" name="htmlBody">
                </div>
            </div>

            <div class="p-6 border-t-2 dark:border-gray-700 bg-gray-50 dark:bg-gray-900/50">
                <div class="mb-4 flex items-center gap-3 p-3 bg-white dark:bg-gray-800 rounded-lg border border-gray-200 dark:border-gray-700 shadow-sm">
                    <input type="checkbox" id="saveTemplateCheck" class="w-5 h-5 text-blue-600 rounded focus:ring-blue-500 border-gray-300" onchange="toggleTemplateNameInput()">
                    <label for="saveTemplateCheck" class="font-medium text-gray-700 dark:text-gray-300 cursor-pointer select-none">Save this email as a new template</label>
                    <input type="text" id="newTemplateName" class="hidden ml-2 flex-1 p-2 text-sm border border-gray-300 dark:border-gray-600 rounded-md dark:bg-gray-700 focus:ring-2 focus:ring-blue-500 outline-none" placeholder="Enter template name (e.g., 'Initial Outreach V1')">
                </div>

                <!-- NEW: SENDING PROGRESS INDICATOR -->
                <div id="sendingProgress" class="hidden mb-4">
                    <div class="flex justify-between text-sm font-bold text-gray-700 dark:text-gray-300 mb-2">
                        <span id="sendingText">Sending...</span>
                        <span id="sendingCount">0/0</span>
                    </div>
                    <div class="w-full bg-gray-200 rounded-full h-3 dark:bg-gray-700 overflow-hidden">
                        <div id="sendingBar" class="bg-blue-600 h-3 rounded-full transition-all duration-300" style="width: 0%"></div>
                    </div>
                </div>

                <div id="emailActionButtons" class="flex justify-end gap-3">
                    <button onclick="document.getElementById('emailModal').classList.add('hidden')" class="px-6 py-2.5 text-gray-600 dark:text-gray-300 hover:bg-gray-200 rounded-lg font-semibold">Cancel</button>
                    <button id="sendBtn" onclick="sendEmails()" class="bg-blue-600 text-white px-8 py-2.5 rounded-lg font-bold hover:bg-blue-700 shadow-md flex items-center gap-2">
                        <i class="fa-solid fa-paper-plane"></i>
                        <span>Send Campaign</span>
                    </button>
                </div>
            </div>
        </div>
    </div>
    
    <!-- MARK SENT MODAL -->
    <div id="markSentModal" class="fixed inset-0 z-50 hidden bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
        <div class="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl p-8 w-full max-w-md border border-gray-200 dark:border-gray-700">
            <div class="mb-6">
                <div class="flex items-center gap-3 mb-3">
                    <div class="w-12 h-12 bg-amber-100 dark:bg-amber-900/30 rounded-full flex items-center justify-center">
                        <i class="fa-solid fa-hand text-amber-600 dark:text-amber-400 text-xl"></i>
                    </div>
                    <h3 class="text-2xl font-bold text-gray-900 dark:text-white">Mark as Manually Sent</h3>
                </div>
                <p class="text-sm text-gray-500 dark:text-gray-400">This will mark <strong id="markSentCount" class="text-amber-600 dark:text-amber-400">0</strong> developer(s) as manually sent (M.Sent). Use this for emails sent outside this system.</p>
            </div>
            
            <div class="mb-6">
                <label class="block text-xs font-bold uppercase mb-2 text-gray-700 dark:text-gray-300 tracking-wide">Note (Optional)</label>
                <input type="text" id="manualSentNote" class="w-full p-3 border-2 border-gray-200 dark:border-gray-600 rounded-xl dark:bg-gray-700 outline-none focus:ring-2 focus:ring-amber-500 font-medium" placeholder="e.g., Sent by colleague, LinkedIn message, etc.">
            </div>
            
            <div class="flex justify-end gap-3">
                <button onclick="closeMarkSentModal()" class="px-6 py-2.5 text-gray-600 dark:text-gray-300 hover:bg-gray-100 rounded-lg font-semibold">Cancel</button>
                <button onclick="confirmMarkAsSent()" id="confirmMarkSentBtn" class="px-6 py-2.5 bg-amber-500 text-white rounded-lg font-bold hover:bg-amber-600 flex items-center gap-2">
                    <i class="fa-solid fa-check"></i>
                    <span>Confirm</span>
                </button>
            </div>
        </div>
    </div>

    <script>
        let tableData = [];
        let fullTableData = [];
        let selectedDevs = new Set();
        let table;
        let jobTemplates = [];
        
        // MULTI-SELECT FILTER STATE
        let activeFilters = {
            status: new Set(['all']), // Default to 'all'
            type: new Set()
        };

        $(document).ready(() => {
            // Fetch User Email Immediately on Load
            google.script.run.withSuccessHandler(function(email) {
               if(email) {
                   document.getElementById('currentUserDisplay').innerText = email;
               } else {
                   document.getElementById('currentUserDisplay').innerText = "Unknown (Check deploy settings)";
               }
            }).getUserEmail();

            google.script.run.withSuccessHandler(url => { if(url) document.getElementById('sheetUrl').value = url; }).getStoredSheetUrl();
            
            table = $('#devTable').DataTable({ 
                dom: 'tp', 
                pageLength: 25,
                columnDefs: [{ orderable: false, targets: 0 }],
                // BUG FIX: Redraw callback to maintain selection state across pages/filtering
                drawCallback: function() {
                    // Re-check boxes that are in our selected set
                    $('.row-checkbox').each(function() {
                        if(selectedDevs.has(this.value)) {
                            $(this).prop('checked', true);
                        } else {
                            $(this).prop('checked', false);
                        }
                    });
                    
                    // Update master checkbox based on current page visibility
                    updateMasterCheckboxState();
                }
            });
            
            selectedDevs.clear();
            updateBulkBar();
        });

        // ========== SMART CHIP FUNCTIONS ==========
        function toggleChip(chip) {
            chip.classList.toggle('active');
            chip.classList.toggle('inactive');
        }

        function selectAllChips() {
            document.querySelectorAll('.smart-chip').forEach(chip => {
                chip.classList.remove('inactive');
                chip.classList.add('active');
            });
        }

        function clearAllChips() {
            document.querySelectorAll('.smart-chip').forEach(chip => {
                chip.classList.remove('active');
                chip.classList.add('inactive');
            });
        }

        function getSelectedStages() {
            return Array.from(document.querySelectorAll('.smart-chip.active')).map(chip => chip.dataset.value);
        }

        // ========== MULTI-SELECT FILTER CHIP FUNCTIONS ==========
        function toggleFilterChip(element) {
            const filter = element.dataset.filter;
            const group = element.dataset.group;
            
            // Special handling for "All" in status group
            if (filter === 'all') {
                // If clicking "All", clear other status filters and activate "All"
                activeFilters.status.clear();
                activeFilters.status.add('all');
                
                // Update UI - deactivate other status chips, activate "All"
                document.querySelectorAll('.filter-chip[data-group="status"]').forEach(chip => {
                    chip.classList.remove('active');
                });
                element.classList.add('active');
            } else if (group === 'status') {
                // Clicking a specific status filter
                // Remove "All" if it's active
                activeFilters.status.delete('all');
                document.querySelector('.filter-chip[data-filter="all"]').classList.remove('active');
                
                // Toggle this filter
                if (element.classList.contains('active')) {
                    element.classList.remove('active');
                    activeFilters.status.delete(filter);
                    
                    // If no status filters active, default to "All"
                    if (activeFilters.status.size === 0) {
                        activeFilters.status.add('all');
                        document.querySelector('.filter-chip[data-filter="all"]').classList.add('active');
                    }
                } else {
                    element.classList.add('active');
                    activeFilters.status.add(filter);
                }
            } else if (group === 'type') {
                // Type filters are independent toggles
                if (element.classList.contains('active')) {
                    element.classList.remove('active');
                    activeFilters.type.delete(filter);
                } else {
                    element.classList.add('active');
                    activeFilters.type.add(filter);
                }
            }
            
            applyFilters();
            updateActiveFiltersIndicator();
        }
        
        // ========== PASTE FILTER FUNCTIONS (NEW) ==========
        function togglePasteFilter() {
            const area = document.getElementById('pasteFilterArea');
            if (area.classList.contains('hidden')) {
                area.classList.remove('hidden');
                document.getElementById('pasteFilterInput').focus();
            } else {
                area.classList.add('hidden');
            }
        }

        function applyPasteFilter() {
            const rawInput = document.getElementById('pasteFilterInput').value;
            if (!rawInput || rawInput.trim() === '') {
                alert("Please paste some Emails or IDs first.");
                return;
            }

            // Split by comma or newline, clean whitespace, remove empty entries
            const filterList = rawInput.split(/[\n,]+/).map(item => item.trim()).filter(item => item !== '');
            
            if (filterList.length === 0) return;

            // Filter fullTableData to only include matches
            const filtered = fullTableData.filter(d => {
                const idMatch = filterList.includes(d.developer_id.toString());
                const emailMatch = filterList.includes(d.email);
                return idMatch || emailMatch;
            });

            // Update UI with filtered data
            tableData = filtered;
            populateTable(filtered);
            
            // Re-apply existing visual chips filters on top of this list if needed
            // For now, we assume the list overrides general filters, but we update counts
            document.getElementById('selectAllCount').textContent = `(${tableData.length})`;
            alert(`Found ${filtered.length} matches out of ${fullTableData.length} records.`);
            
            // Close the area
            togglePasteFilter();
        }

        function clearPasteFilter() {
            document.getElementById('pasteFilterInput').value = '';
            // Reset to full data
            tableData = [...fullTableData];
            populateTable(tableData);
            document.getElementById('selectAllCount').textContent = `(${tableData.length})`;
            // Re-apply active filters
            applyFilters();
        }

        // ========== FILTER LOGIC ==========
        function applyFilters() {
            // Start with the full dataset (or the specific list filtered dataset if we want to combine them, currently resetting to fullTableData for simplicity unless complex stacking is needed)
            // Ideally, we should filter on tableData if tableData was modified by Paste Filter, 
            // BUT usually filters apply to the 'source'. Let's assume Filters apply to full data for consistency.
            let filtered = [...fullTableData];
            
            // If user has pasted content, check if we should respect that subset. 
            // For this implementation, the "status/type" filters apply to the FULL set. 
            // If you want "Paste Filter" to be the master set, we would filter `filtered` against the pasted list here.
            
            // Apply status filters (OR logic within group)
            if (!activeFilters.status.has('all') && activeFilters.status.size > 0) {
                filtered = filtered.filter(d => {
                    if (activeFilters.status.has('new') && !d.is_sent && !d.is_manual_sent) return true;
                    if (activeFilters.status.has('followup') && d.is_sent && !d.is_manual_sent) return true;
                    if (activeFilters.status.has('manual') && d.is_manual_sent) return true;
                    return false;
                });
            }
            
            // Apply type filters (OR logic within group, AND with status)
            if (activeFilters.type.size > 0) {
                filtered = filtered.filter(d => {
                    const candidateType = d.candidate_status || 'Independent';
                    if (activeFilters.type.has('sub') && candidateType === 'Agency Sub-Contractor') return true;
                    if (activeFilters.type.has('agency') && candidateType === 'Agency Contractor') return true;
                    if (activeFilters.type.has('direct') && (candidateType === 'Independent' || !d.candidate_status)) return true;
                    return false;
                });
            }
            
            tableData = filtered;
            populateTable(tableData);
            document.getElementById('selectAllCount').textContent = `(${tableData.length})`;
        }
        
        function updateActiveFiltersIndicator() {
            const totalActive = (activeFilters.status.has('all') ? 0 : activeFilters.status.size) + activeFilters.type.size;
            const indicator = document.getElementById('activeFiltersInfo');
            const countEl = document.getElementById('activeFilterCount');
            
            if (totalActive > 0) {
                indicator.classList.remove('hidden');
                indicator.classList.add('flex');
                countEl.textContent = totalActive;
            } else {
                indicator.classList.add('hidden');
                indicator.classList.remove('flex');
            }
        }
        
        function clearAllFilters() {
            // Reset status to "All"
            activeFilters.status.clear();
            activeFilters.status.add('all');
            
            // Clear type filters
            activeFilters.type.clear();
            
            // Update UI
            document.querySelectorAll('.filter-chip').forEach(chip => {
                chip.classList.remove('active');
            });
            document.querySelector('.filter-chip[data-filter="all"]').classList.add('active');

            // CRITICAL FIX: Explicitly clear the text search box and DataTables search state
            document.getElementById('tableSearch').value = '';
            if(table) table.search('').draw();
            
            applyFilters();
            updateActiveFiltersIndicator();
        }

        // ========== SEARCH FUNCTION ==========
        function searchTable() {
            const searchValue = document.getElementById('tableSearch').value;
            table.search(searchValue).draw();
        }

        // ========== TABLE RENDERING ==========
        function renderTable(data) {
            completeProgress();
            tableData = data;
            fullTableData = data;
            document.getElementById('tableContainer').classList.remove('hidden');
            document.getElementById('tableControls').classList.remove('hidden');
            document.getElementById('statusFilterContainer').classList.remove('hidden');
            
            renderStatusButtons(data);
            updateCounts(data);
            selectedDevs.clear();
            document.getElementById('masterCheckbox').checked = false;
            
            // CRITICAL FIX: Clear search box value AND DataTables search state
            document.getElementById('tableSearch').value = '';
            if(table) table.search('').draw();
            
            // Reset filters logic but don't cause double render if unnecessary
            // Manually reset filter state without calling populateTable yet
            activeFilters.status.clear();
            activeFilters.status.add('all');
            activeFilters.type.clear();
            document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
            document.querySelector('.filter-chip[data-filter="all"]').classList.add('active');
            updateActiveFiltersIndicator();
            
            updateBulkBar();
            populateTable(data);

            // === CHANGE 2: Clear Draft Email Campaign Box ===
            // This ensures previous data doesn't persist when fetching new IDs
            document.getElementById('emailSubject').value = '';
            document.getElementById('emailBody').innerHTML = '';
            document.getElementById('htmlBody').value = '';
            document.getElementById('templateLoader').value = '';
            document.getElementById('saveTemplateCheck').checked = false;
            document.getElementById('newTemplateName').classList.add('hidden');
            // We do NOT clear senderName as users typically want that to persist
        }

        function updateCounts(data) {
            const allCount = data.length;
            const newCount = data.filter(d => !d.is_sent && !d.is_manual_sent).length;
            const followupCount = data.filter(d => d.is_sent && !d.is_manual_sent).length;
            const manualCount = data.filter(d => d.is_manual_sent).length;
            
            const subCount = data.filter(d => d.candidate_status === 'Agency Sub-Contractor').length;
            const agencyCount = data.filter(d => d.candidate_status === 'Agency Contractor').length;
            const nonSubCount = data.filter(d => d.candidate_status === 'Independent' || !d.candidate_status).length;
            
            document.getElementById('allCount').textContent = allCount;
            document.getElementById('newCount').textContent = newCount;
            document.getElementById('followupCount').textContent = followupCount;
            document.getElementById('manualCount').textContent = manualCount;
            document.getElementById('subCount').textContent = subCount;
            document.getElementById('agencyCount').textContent = agencyCount;
            document.getElementById('nonSubCount').textContent = nonSubCount;
            document.getElementById('selectAllCount').textContent = `(${tableData.length})`;
        }

        function populateTable(data) {
            table.clear();
            data.forEach(d => {
                let rowClass = '';
                let sentBadge = '';
                
                if (d.is_manual_sent) {
                    rowClass = 'manual-sent-row';
                    sentBadge = `<span class="status-badge bg-amber-100 text-amber-800 text-xs px-2 py-1 rounded border border-amber-200" title="${d.manual_sent_note || 'Manually marked as sent'}">
                        <i class="fa-solid fa-hand mr-1"></i>M.Sent${d.manual_sent_count > 1 ? ` (${d.manual_sent_count})` : ''}
                    </span>`;
                } else if (d.is_sent) {
                    rowClass = 'sent-row';
                    sentBadge = `<span class="status-badge bg-red-100 text-red-800 text-xs px-2 py-1 rounded border border-red-200">Sent (${d.sent_count})</span>`;
                } else {
                    sentBadge = `<span class="status-badge bg-green-100 text-green-800 text-xs px-2 py-1 rounded border border-green-200">New</span>`;
                }
                
                let typeBadge = '';
                const candidateStatus = d.candidate_status || 'Independent';
                
                if (candidateStatus === 'Agency Sub-Contractor') {
                    typeBadge = `<span class="type-badge subcontractor" title="${d.agency_name || 'Sub-Contractor'}">
                        <i class="fa-solid fa-user-tie mr-1"></i>Sub
                    </span>`;
                } else if (candidateStatus === 'Agency Contractor') {
                    typeBadge = `<span class="type-badge agency" title="${d.agency_name || 'Agency Contractor'}">
                        <i class="fa-solid fa-building mr-1"></i>Agency
                    </span>`;
                } else {
                    typeBadge = `<span class="type-badge independent">
                        <i class="fa-solid fa-user mr-1"></i>Direct
                    </span>`;
                }

                // Initial checked state based on set
                const isChecked = selectedDevs.has(d.developer_id) ? 'checked' : '';
                
                const rowNode = table.row.add([
                    `<div class="text-center"><input type="checkbox" class="row-checkbox w-4 h-4 cursor-pointer" value="${d.developer_id}" onchange="toggleRow('${d.developer_id}')" ${isChecked}></div>`,
                    `<span class="font-mono text-sm text-gray-600 dark:text-gray-400 font-semibold select-all">${d.developer_id}</span>`,
                    `<div class="font-bold text-gray-900 dark:text-white text-base">${d.full_name}</div>`,
                    `<div class="text-sm text-gray-600 dark:text-gray-400 font-mono">${d.email}</div>`,
                    `<span class="status-badge text-sm bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 px-3 py-1.5 rounded-lg border border-gray-200 dark:border-gray-600 font-semibold">${d.status}</span>`,
                    typeBadge,
                    sentBadge
                ]).node();
                $(rowNode).addClass(rowClass);
            });
            table.draw();
        }

        // ========== SELECTION FUNCTIONS ==========
        function selectAllResults() {
            // Select everything in the current filtered tableData
            tableData.forEach(d => selectedDevs.add(d.developer_id));
            
            // Visually check only visible ones immediately (others handled by drawCallback)
            $('.row-checkbox').prop('checked', true);
            updateMasterCheckboxState();
            updateBulkBar();
        }

        function selectCurrentPage() {
            $('.row-checkbox:visible').prop('checked', true);
            $('.row-checkbox:visible').each(function() {
                selectedDevs.add(this.value);
            });
            updateBulkBar();
        }

        function deselectAll() {
            selectedDevs.clear();
            $('.row-checkbox').prop('checked', false);
            updateMasterCheckboxState();
            updateBulkBar();
        }

        function toggleRow(id) {
            // DataTables returns ID as string usually, ensure type safety
            const devId = id.toString();
            selectedDevs.has(devId) ? selectedDevs.delete(devId) : selectedDevs.add(devId);
            updateBulkBar();
            updateMasterCheckboxState();
        }
        
        function toggleMasterCheckbox() {
             const isChecked = document.getElementById('masterCheckbox').checked;
             $('.row-checkbox:visible').each(function() {
                 this.checked = isChecked;
                 isChecked ? selectedDevs.add(this.value) : selectedDevs.delete(this.value);
             });
             updateBulkBar();
        }
        
        // BUG FIX: Intelligently update master checkbox based on visible rows
        function updateMasterCheckboxState() {
            const visibleCheckboxes = $('.row-checkbox:visible');
            if(visibleCheckboxes.length === 0) {
                document.getElementById('masterCheckbox').checked = false;
                return;
            }
            
            let allVisibleSelected = true;
            visibleCheckboxes.each(function() {
                if(!this.checked) allVisibleSelected = false;
            });
            
            document.getElementById('masterCheckbox').checked = allVisibleSelected;
        }

        function updateBulkBar() {
            const count = selectedDevs.size;
            document.getElementById('selectedCount').innerText = count;
            document.getElementById('quickSelectedCount').innerText = count;
            
            const bar = document.getElementById('bulkActionBar');
            const indicator = document.getElementById('selectionIndicator');
            
            if (count > 0) {
                bar.classList.remove('translate-y-[200%]');
                indicator.classList.remove('hidden');
                indicator.classList.add('flex');
            } else {
                bar.classList.add('translate-y-[200%]');
                indicator.classList.add('hidden');
                indicator.classList.remove('flex');
            }
        }
        
        // ========== MARK AS MANUAL SENT FUNCTIONS ==========
        function markAsManualSent() {
            if (selectedDevs.size === 0) {
                alert('Please select at least one developer to mark as sent.');
                return;
            }
            document.getElementById('markSentCount').textContent = selectedDevs.size;
            document.getElementById('manualSentNote').value = '';
            document.getElementById('markSentModal').classList.remove('hidden');
        }
        
        function closeMarkSentModal() {
            document.getElementById('markSentModal').classList.add('hidden');
        }
        
        function confirmMarkAsSent() {
            const note = document.getElementById('manualSentNote').value.trim();
            const jobId = document.getElementById('jobId').value;
            
            const developerIds = Array.from(selectedDevs);
            
            const btn = document.getElementById('confirmMarkSentBtn');
            const originalBtnContent = btn.innerHTML;
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin mr-2"></i>Marking...';
            btn.disabled = true;
            
            google.script.run
                .withSuccessHandler(res => {
                    btn.innerHTML = originalBtnContent;
                    btn.disabled = false;
                    closeMarkSentModal();
                    
                    if (res.success) {
                        alert(`Successfully marked ${res.marked} developer(s) as manually sent.`);
                        fetchData();
                    } else {
                        alert("Error: " + (res.error || "Unknown error occurred"));
                    }
                })
                .withFailureHandler(e => {
                    btn.innerHTML = originalBtnContent;
                    btn.disabled = false;
                    alert("Error: " + (e ? e.message : "Unknown error"));
                })
                .markAsManualSent(developerIds, jobId, note);
        }

        // ========== TEMPLATE LOGIC - IMPROVED ==========
        function openDraftModal() { 
            document.getElementById('emailModal').classList.remove('hidden'); 
            loadTemplatesForJob();
            
            // Hide progress bar when opening modal
            document.getElementById('sendingProgress').classList.add('hidden');
            document.getElementById('emailActionButtons').classList.remove('hidden');
        }
        
        function loadTemplatesForJob() {
            const jobId = document.getElementById('jobId').value;
            const statusEl = document.getElementById('templateStatus');
            
            if(!jobId) {
                showTemplateStatus('warning', 'Enter a Job ID to load saved templates');
                return;
            }
            
            // Show loading state
            showTemplateStatus('loading', 'Loading templates...');
            
            google.script.run
                .withSuccessHandler(fillTemplateDropdown)
                .withFailureHandler(err => {
                    console.error('Template load error:', err);
                    showTemplateStatus('error', 'Failed to load templates: ' + (err.message || 'Unknown error'));
                })
                .getJobTemplates(jobId);
        }
        
        function refreshTemplates() {
            loadTemplatesForJob();
        }
        
        function showTemplateStatus(type, message) {
            const statusEl = document.getElementById('templateStatus');
            statusEl.classList.remove('hidden');
            
            if (type === 'loading') {
                statusEl.className = 'template-loading';
                statusEl.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> ${message}`;
            } else if (type === 'error') {
                statusEl.className = 'template-loading';
                statusEl.innerHTML = `<i class="fa-solid fa-exclamation-triangle"></i> ${message}`;
            } else if (type === 'warning') {
                statusEl.className = 'template-empty';
                statusEl.innerHTML = `<i class="fa-solid fa-info-circle"></i> ${message}`;
            } else if (type === 'empty') {
                statusEl.className = 'template-empty';
                statusEl.innerHTML = `<i class="fa-solid fa-folder-open"></i> ${message}`;
            } else {
                statusEl.classList.add('hidden');
            }
        }

        function fillTemplateDropdown(templates) {
            jobTemplates = templates || [];
            const select = document.getElementById('templateLoader');
            select.innerHTML = '<option value="">-- Select a saved template --</option>';
            
            if (!templates || templates.length === 0) {
                showTemplateStatus('empty', 'No saved templates for this job. Send an email and check "Save as template" to create one.');
                return;
            }
            
            // Hide status message
            showTemplateStatus('hide', '');
            
            templates.forEach((t, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.text = t.name || `Template ${index + 1}`;
                select.appendChild(option);
            });
            
            console.log(`Loaded ${templates.length} templates`);
        }

        function loadTemplate(index) {
            if (index === "" || index === null || index === undefined) return;
            const t = jobTemplates[index];
            if (t) {
                if(confirm("Replace current content with this template?")) {
                    document.getElementById('emailSubject').value = t.subject || '';
                    document.getElementById('emailBody').innerHTML = t.body || '';
                    handleEditorInput();
                }
            }
        }

        function toggleTemplateNameInput() {
            const isChecked = document.getElementById('saveTemplateCheck').checked;
            const input = document.getElementById('newTemplateName');
            if (isChecked) {
                input.classList.remove('hidden');
                input.focus();
            } else {
                input.classList.add('hidden');
            }
        }

        // ========== ENCODING & SENDING LOGIC (UPDATED WITH BATCHING) ==========
        function cleanText(text) {
            if (!text) return "";
            return text
                .replace(/[\u2018\u2019]/g, "'")
                .replace(/[\u201C\u201D]/g, '"')
                .replace(/[\u2013\u2014]/g, "-")
                .replace(/\u2026/g, "...");
        }

        // Wrapper function to make google.script.run await-able
        function sendEmailBatch(recipientsBatch, senderName, subject, htmlBody, jobId, templateSettings) {
            return new Promise((resolve, reject) => {
                google.script.run
                    .withSuccessHandler(resolve)
                    .withFailureHandler(reject)
                    .sendBulkEmails(recipientsBatch, senderName, subject, htmlBody, jobId, templateSettings);
            });
        }

        async function sendEmails() {
            // 1. Gather Inputs
            const senderName = cleanText(document.getElementById('senderName').value);
            const subject = cleanText(document.getElementById('emailSubject').value);
            let htmlBody = cleanText(document.getElementById('htmlBody').value);
            const jobId = document.getElementById('jobId').value;

            // 2. Validate
            if(!htmlBody || htmlBody.trim() === '') return alert('Please enter email content');

            const shouldSave = document.getElementById('saveTemplateCheck').checked;
            const templateName = cleanText(document.getElementById('newTemplateName').value);
            
            if (shouldSave && !templateName) return alert("Please enter a name for your template.");

            // 3. Prepare Recipient List
            const recipients = [];
            selectedDevs.forEach(id => {
                const d = fullTableData.find(x => x.developer_id == id);
                if(d) recipients.push({ name: d.full_name, email: d.email });
            });
            
            if (recipients.length === 0) return alert("No recipients selected!");

            // 4. UI Setup for Sending Mode
            const btn = document.getElementById('sendBtn');
            const actionButtons = document.getElementById('emailActionButtons');
            const progressContainer = document.getElementById('sendingProgress');
            const sendingText = document.getElementById('sendingText');
            const sendingCount = document.getElementById('sendingCount');
            const sendingBar = document.getElementById('sendingBar');

            actionButtons.classList.add('hidden'); // Hide buttons
            progressContainer.classList.remove('hidden'); // Show progress bar
            
            // 5. Batching Logic
            const BATCH_SIZE = 5; // Send 5 emails at a time to allow UI updates
            const total = recipients.length;
            let sentCount = 0;
            let successCount = 0;
            let errors = [];

            try {
                // Loop through batches
                for (let i = 0; i < total; i += BATCH_SIZE) {
                    const batch = recipients.slice(i, i + BATCH_SIZE);
                    const currentBatchEnd = Math.min(i + BATCH_SIZE, total);
                    
                    // Update UI text
                    sendingText.textContent = `Sending batch ${Math.ceil((i+1)/BATCH_SIZE)}...`;
                    sendingCount.textContent = `${i} / ${total}`;
                    const percent = Math.floor((i / total) * 100);
                    sendingBar.style.width = `${percent}%`;

                    // Only save template on the first batch to avoid duplicates
                    const batchTemplateSettings = {
                        shouldSave: (i === 0) ? shouldSave : false, 
                        templateName: templateName
                    };

                    // Wait for batch to complete
                    try {
                        const result = await sendEmailBatch(batch, senderName, subject, htmlBody, jobId, batchTemplateSettings);
                        if (result.success) {
                            successCount += (result.sent || 0);
                        } else {
                            errors.push(...(result.errors || ["Unknown error in batch"]));
                        }
                    } catch (err) {
                        errors.push(err.message || "Batch failure");
                    }
                    
                    // Small delay to ensure UI renders smoothly
                    await new Promise(r => setTimeout(r, 100));
                }

                // Final Update
                sendingBar.style.width = `100%`;
                sendingCount.textContent = `${total} / ${total}`;
                sendingText.textContent = "Completed!";
                
                // Final Alert
                setTimeout(() => {
                    document.getElementById('emailModal').classList.add('hidden');
                    let msg = `Completed! Successfully sent ${successCount} out of ${total} emails.`;
                    if(errors.length > 0) msg += `\nErrors occurred: ${errors.length}`;
                    alert(msg);
                    fetchData(); // Refresh table to show updated status
                }, 500);

            } catch (globalError) {
                console.error(globalError);
                alert("Critical error during sending: " + globalError.message);
                actionButtons.classList.remove('hidden');
                progressContainer.classList.add('hidden');
            }
        }

        // ========== UTILS & BOILERPLATE ==========
        function fetchData() {
            const jobId = document.getElementById('jobId').value;
            const stages = getSelectedStages();
            if(!jobId || stages.length === 0) return alert("Enter Job ID and select at least one pipeline stage");

            startSmartProgress();
            document.getElementById('tableContainer').classList.add('hidden');
            
            google.script.run.withSuccessHandler(renderTable).withFailureHandler(handleError).getDevelopers(jobId, stages);
        }

        function handleEditorInput() {
            document.getElementById('htmlBody').value = document.getElementById('emailBody').innerHTML;
        }
        
        function formatText(cmd) {
            const editor = document.getElementById('emailBody');
            editor.focus();
            if (cmd === 'h2') {
                document.execCommand('formatBlock', false, '<h2>');
            } else {
                document.execCommand(cmd, false, null);
            }
            handleEditorInput();
            editor.focus();
        }

        function startSmartProgress() {
            const fill = document.getElementById('progressFill');
            document.getElementById('smartProgress').classList.remove('hidden');
            fill.style.width = '10%';
            setTimeout(() => fill.style.width = '40%', 500);
            setTimeout(() => fill.style.width = '70%', 1500);
        }

        function completeProgress() {
            document.getElementById('progressFill').style.width = '100%';
            setTimeout(() => document.getElementById('smartProgress').classList.add('hidden'), 500);
        }

        function handleError(e) {
            document.getElementById('smartProgress').classList.add('hidden');
            alert("Error: " + (e ? e.message : "Unknown error"));
        }

        function renderStatusButtons(data) {
            const counts = {};
            data.forEach(d => counts[d.status] = (counts[d.status] || 0) + 1);
            const container = document.getElementById('statusButtons');
            container.innerHTML = `<button onclick="filterTableByStatus('All')" class="text-sm px-3 py-1.5 bg-blue-600 text-white rounded-lg shadow font-semibold">All (${data.length})</button>`;
            Object.keys(counts).forEach(s => {
                container.innerHTML += `<button onclick="filterTableByStatus('${s}')" class="text-sm px-3 py-1.5 bg-white dark:bg-gray-700 border border-gray-200 dark:border-gray-600 rounded-lg shadow hover:bg-gray-50 dark:hover:bg-gray-600 font-medium">${s} (${counts[s]})</button>`;
            });
        }

        function filterTableByStatus(status) {
            // CRITICAL FIX: Ensure consistent state when switching views
            
            // 1. Reset visual chips
            activeFilters.status.clear();
            activeFilters.status.add('all');
            activeFilters.type.clear();
            document.querySelectorAll('.filter-chip').forEach(chip => chip.classList.remove('active'));
            document.querySelector('.filter-chip[data-filter="all"]').classList.add('active');
            updateActiveFiltersIndicator();
            
            // 2. Clear search box and DataTables state
            document.getElementById('tableSearch').value = '';
            if(table) table.search('').draw();

            // 3. Filter Data
            const filtered = status === 'All' ? fullTableData : fullTableData.filter(d => d.status === status);
            tableData = filtered;

            // 4. Update Table
            populateTable(filtered);
            document.getElementById('selectAllCount').textContent = `(${filtered.length})`;
        }
        
        function openConfigModal() { document.getElementById('configModal').classList.remove('hidden'); }
        function closeConfigModal() { document.getElementById('configModal').classList.add('hidden'); }
        function saveConfig() { google.script.run.withSuccessHandler(() => { closeConfigModal(); alert("Saved!"); }).saveLogSheetUrl(document.getElementById('sheetUrl').value); }
        
        function toggleDarkMode() {
            document.documentElement.classList.toggle('dark');
        }
        
        // ========== COPY & EXPORT FUNCTIONS ==========
        function copyToClipboard() {
            let t = "Developer ID\tName\tEmail\tPipeline Status\tType\tOutreach History\n";
            tableData.forEach(d => {
                let history = 'New';
                if (d.is_manual_sent) {
                    history = `M.Sent${d.manual_sent_count > 1 ? ` (${d.manual_sent_count})` : ''}`;
                } else if (d.is_sent) {
                    history = `Sent (${d.sent_count})`;
                }
                
                let typeDisplay = 'Direct';
                if (d.candidate_status === 'Agency Sub-Contractor') typeDisplay = 'Sub-Contractor';
                else if (d.candidate_status === 'Agency Contractor') typeDisplay = 'Agency';
                
                t += `${d.developer_id}\t${d.full_name}\t${d.email}\t${d.status}\t${typeDisplay}\t${history}\n`;
            });
            navigator.clipboard.writeText(t).then(() => alert('Copied to clipboard!'));
        }

        function downloadCSV() {
            let csv = "Developer ID,Candidate Name,Email,Pipeline Status,Type,Agency Name,Outreach History\n";
            tableData.forEach(d => {
                let history = 'New';
                if (d.is_manual_sent) {
                    history = `M.Sent${d.manual_sent_count > 1 ? ` (${d.manual_sent_count})` : ''}`;
                } else if (d.is_sent) {
                    history = `Sent (${d.sent_count})`;
                }
                
                let typeDisplay = 'Direct';
                if (d.candidate_status === 'Agency Sub-Contractor') typeDisplay = 'Sub-Contractor';
                else if (d.candidate_status === 'Agency Contractor') typeDisplay = 'Agency';
                
                const agencyName = d.agency_name || '';
                csv += `${d.developer_id},"${d.full_name}","${d.email}","${d.status}","${typeDisplay}","${agencyName}","${history}"\n`;
            });
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'candidates.csv';
            a.click();
        }
    </script>
</body>
</html>
