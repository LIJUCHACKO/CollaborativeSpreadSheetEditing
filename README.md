# Collaborative Spreadsheet & Document Editor

A real-time collaborative platform for editing spreadsheets and structured documents with **interdependencies between cells, sheets, and projects**. Multiple users can simultaneously work on the same sheets, with changes synced instantly. Cells can contain Python scripts or AI-generated content that automatically reference and react to other cells — even across different sheets — creating a live, interconnected workspace.

---

## Table of Contents

- [Purpose](#purpose)
- [Key Features](#key-features)
- [Getting Started](#getting-started)
- [User Guide](#user-guide)
  - [Projects & Sheets](#projects--sheets)
  - [DataSheet (Spreadsheet)](#datasheet-spreadsheet)
  - [Document](#document)
  - [Cell Types](#cell-types)
  - [Cell References & Interdependencies](#cell-references--interdependencies)
  - [Python Scripting](#python-scripting)
  - [AI-Generated Cells](#ai-generated-cells)
  - [Markdown Editor (Documents)](#markdown-editor-documents)
  - [Chat & Communication](#chat--communication)
  - [Timeline & Milestones](#timeline--milestones)
  - [Import & Export](#import--export)
  - [Public API](#public-api)
  - [Assets & Files](#assets--files)
  - [Help & Documentation](#help--documentation)
- [User Roles & Permissions](#user-roles--permissions)
- [Administration](#administration)
- [Architecture Overview](#architecture-overview)
- [Technology Stack](#technology-stack)

---

## Purpose

This platform is designed for teams that need to **collaboratively edit interconnected documents and spreadsheets** in real time. Unlike traditional spreadsheet tools, it focuses on:

- **Cross-sheet and cross-project dependencies** — cells can reference values from other sheets and projects, and changes automatically propagate through the dependency chain.
- **Structured documents** — in addition to flat spreadsheets, it supports tree-structured documents with hierarchical sections, titles, and rich markdown content.
- **Scriptable cells** — embed Python scripts directly in cells to compute values from other cells, even across sheets.
- **AI-powered cells** — use LLM prompts with cell references to generate content dynamically.
- **Integrity & auditability** — every edit is tracked, every file is checksum-verified, and full audit logs are available.

---

## Key Features

| Feature | Description |
|---|---|
| **Real-Time Collaboration** | Multiple users edit the same sheet simultaneously. All changes appear instantly for everyone. |
| **Two Document Types** | **DataSheet** (traditional spreadsheet grid) and **Document** (tree-structured with sections, titles, and rich content). |
| **Cell Interdependencies** | Cells can reference other cells within the same sheet, across sheets, or across projects using `{{}}` syntax. |
| **Python Scripting** | Write Python scripts in cells that read from other cells and write computed results back. Scripts auto-execute when referenced cells change. |
| **AI-Generated Content** | Cells powered by an LLM — write a prompt with cell references and get AI-generated output that stays up to date. |
| **Rich Markdown Editor** | Full markdown editing with LaTeX math, tables, images, checklists, and code blocks for document content cells. |
| **Cell Locking** | Lock cells to prevent accidental edits by other users. |
| **Named Cells** | Assign human-readable names to cells (e.g., `TotalBudget`) and reference them by name. |
| **ComboBox & MultiSelect** | Cells with dropdown options — options can be sourced dynamically from a range of cells. |
| **Audit Trail** | Full edit history per cell with diff view, plus project-level audit logs for structural changes. |
| **Project Organization** | Projects → Subfolders → Sheets. Organize work with folder hierarchies. |
| **Copy & Paste** | Copy/paste sheets, folders, and entire projects — even across different projects. |
| **Import/Export** | Import and export sheets and projects as XLSX files. |
| **Image Assets** | Upload and embed images directly into your projects. |
| **Timeline** | Track project milestones and events on a visual timeline. |
| **Chat** | Built-in real-time chat for team communication. |
| **File Integrity** | All data files are checksum-verified to detect corruption. |
| **Backup** | Administrators can download a full backup of all data. |

---

## Getting Started

### Prerequisites

- **Go** (for the backend server)
- **Node.js & npm** (for building the frontend)
- **Python 3** (optional — required only if you use Python scripting in cells)

### Running the Application

1. **Build the frontend:**
   ```bash
   cd frontend
   npm install
   npm run build
   ```

2. **Build the backend:**
   ```bash
   cd backend
   go build -o shared-spreadsheet
   ```

3. **Start the application** using the provided script:
   ```bash
   cd backend/
   ./shared-spreadsheet &
   chmod 666 ../DATA/pythonDirectory
   cd  ../frontend/serve
   ./shared-spreadsheet-serve -backend 192.168.0.100:8082 -dist ../dist
   ```
   - Replace "192.168.0.100" with your ip address.
   - The backend runs on port **8082** by default. The frontend is served via a static file server.
   - If -python-user is  provided to backend then python scripts will run as current user eg:- ./shared-spreadsheet -python-user pythonUser & . But in this case backend should run be 'root user'/sudoer.
   - Data will be save in DATA folder outside the code folder.

4. **First login:** On first launch, a default admin account is created:
   - **Username:** `admin`
   - **Password:** `admin`
   - ⚠️ **Change this password immediately** after your first login.

5. For creating a project, a user will register and admin has to grant permission to the new user to create a project.

---

## User Guide

### Projects & Sheets

The workspace is organized into a hierarchy:

```
Projects
├── Project Alpha
│   ├── Budget (DataSheet)
│   ├── Requirements (Document)
│   └── Reports/
│       ├── Q1 Report (DataSheet)
│       └── Q2 Report (DataSheet)
└── Project Beta
    ├── Specs (Document)
    └── Tracking (DataSheet)
```

- **Create a Project** from the Projects page.
- Inside a project, **create Sheets** (DataSheet or Document type) and **Folders** to organize them.
- You can **copy/paste** sheets and folders across projects.
- **Rename** or **delete** projects and sheets from their context menus.

### DataSheet (Spreadsheet)

A traditional spreadsheet grid with collaborative editing:

- Click a cell to edit its value. Press **Enter** to confirm, **Escape** to cancel.
- **Select multiple cells** by clicking and dragging, or use **Shift+Click**.
- **Copy/paste** cells with Ctrl+C / Ctrl+V (values, scripts, cell types, and styles are preserved).
- **Right-click** for context menu: insert/delete rows and columns, move rows, lock/unlock cells, change cell type.
- **Freeze rows/columns** to keep headers visible while scrolling.
- **Sort and filter** columns from the column header menu.
- **Resize** columns and rows by dragging the header borders.
- **Undo/Redo** with Ctrl+Z / Ctrl+Y.
- **Name a cell** using the Name Box (top-left input). Named cells can be referenced by name in scripts and formulas.

### Document

A structured document editor combining spreadsheet functionality with document hierarchy:

- Documents have three default columns: **Sr No.** (A), **Title** (B), and **Content** (C).
- Rows can be organized in a **tree hierarchy** — indent rows to create sections and subsections.
- **Section numbering** is automatic. Choose a numbering scheme: `1.1.1`, `I.A.1`, or custom formats.
- The **Content** column (C) supports full **Markdown editing** via a dedicated editor panel.
- Insert **child rows** to build nested document structures.
- **Move rows** up/down or reparent them to restructure the document.
- Export the entire document structure.

### Cell Types

Each cell can be one of the following types:

| Type | Description |
|---|---|
| **Value** | A plain text or number cell (default). |
| **Script** | Contains a Python script that computes the cell's value. |
| **ComboBox** | A dropdown selector. Options can be manually defined or sourced from a cell range. |
| **MultipleSelection** | Like ComboBox, but allows selecting multiple values. |
| **AI Generated** | Contains an LLM prompt that generates the cell's value using AI. |

Change a cell's type via the **right-click context menu → Set Type**.

### Cell References & Interdependencies

The `{{}}` reference syntax is the core mechanism for creating dependencies between cells:

| Reference | Example | Description |
|---|---|---|
| **Same-sheet cell** | `{{A2}}` | References cell A2 in the current sheet. |
| **Same-sheet range** | `{{A2:B5}}` | References a rectangular range of cells. |
| **Cross-sheet cell** | `{{ProjectName/SheetName/A2}}` | References cell A2 in another sheet/project. |
| **Cross-sheet range** | `{{ProjectName/SheetName/A2:C10}}` | References a range in another sheet/project. |
| **Named cell** | `{{TotalBudget}}` | References a cell by its assigned name. |
| **Cross-sheet named cell** | `{{ProjectName/SheetName/TotalBudget}}` | References a named cell in another sheet/project. |

To insert reference from the same sheet click a cell in the sheet and press 'insert reference' button in the editor. For inserting references from an another shhet open that sheet in another tab and copy the cell and then press 'insert reference' button. In preview tab the actual content after the replacement of references is visible.

**How dependencies work:**

1. When a cell's value changes, the system identifies all scripts and AI cells that reference it.
2. Those dependent cells are automatically re-executed with the updated values.
3. If their outputs change, any further downstream dependencies are triggered in turn.
4. This cascade continues until all dependent cells are up to date.
5. References **automatically adjust** when rows or columns are inserted or deleted.

### Python Scripting

Cells of type **Script** contain Python code that is executed on the server:

1. **Set a cell's type to Script** (right-click → Set Type → Script).
2. The **Script Editor Panel** opens — write your Python code there.
3. Use `{{A2}}` references in your script. They are replaced with actual cell values at execution time.
4. The script's **printed output** (`print(...)`) becomes the cell's value.
5. Scripts can output **multi-cell spans** — configure the row/column span to write results across adjacent cells.

**Example script:**
```python
# Sum values from a column range
values = {{A2:A10}}  # Replaced with actual values at runtime
total = sum([float(v) for v in values if v])
print(total)
```

**Features of the Script Editor:**
- Syntax highlighting and line numbers
- Auto-indentation and smart dedent
- Auto-closing brackets and quotes
- Insert Range button to pick cell references
- Preview tab to see resolved references before execution
- "Show Script As Output" option to display the script source in the cell

### AI-Generated Cells

Cells of type **AI Generated** use an LLM to produce content:

1. **Set a cell's type to AI Generated** (right-click → Set Type → AI Generated).
2. The **AI Prompt Editor** opens — write your prompt using natural language.
3. Include `{{}}` cell references in your prompt. They are resolved to actual values before sending to the LLM.
4. The AI response becomes the cell's value.
5. When referenced cells change, the AI prompt is re-executed automatically.

**Example prompt:**
```
Summarize the following project requirements in 3 bullet points:
{{ProjectAlpha/Requirements/C2:C20}}
```

> **Note:** AI features require an administrator to configure an LLM endpoint (OpenAI-compatible API) in the Admin panel.

### Markdown Editor (Documents)

The Content column in Documents opens a full **Markdown Editor** with:

- **Toolbar** for bold, italic, headings, lists, code blocks, links, images, quotes, checkboxes, math equations, and tables.
- **LaTeX math** support — inline with `$...$` and display blocks with `$$...$$`.
- **Table editing** — right-click tables to insert/delete rows and columns.
- **Image embedding** — browse and insert images from the project's asset library.
- **Python file embedding** — link to uploaded Python files.
- **Split view, preview, and edit** modes.

### Chat & Communication

A built-in real-time chat system accessible from any sheet:

- Send messages to all users or as direct messages.
- Messages show the sheet and project context.
- Read receipts track which messages you've seen.
- Delete your own messages.

### Timeline & Milestones

Each project has a **Timeline** feature:

- Add milestones and events with date/time and descriptions.
- View them on a visual vertical timeline, sorted newest first.
- Use milestones to track project progress and deadlines.

### Import & Export

| Action | Description |
|---|---|
| **Export Sheet** | Download a single sheet as an XLSX file. |
| **Export Project** | Download all sheets in a project as a single XLSX workbook (one sheet per tab). |
| **Import XLSX** | Import an XLSX file into a project — each worksheet becomes a new sheet. |

### Public API

Two read-only HTTP endpoints are available **without any authentication**, making it easy to integrate sheet data into external tools, dashboards, scripts, or automated pipelines using plain `curl` or `wget`.

---

#### `GET /api/public/sheet/csv`

Download the current data of a sheet as a **CSV file**.

**Query parameters:**

| Parameter | Required | Description |
|---|---|---|
| `project` | Yes | The project name (use `/` for subfolders, e.g. `MyProject/Reports`). |
| `sheet_name` | Yes | The name of the sheet to download. |

**Response:** A `text/csv` file download. Each row in the sheet becomes a row in the CSV. Columns are sorted alphabetically by their label (A, B, C, …). The cell values used are the computed/display values (script outputs, AI outputs, or plain values).

**Example:**
```bash
curl "http://localhost:8082/api/public/sheet/csv?project=MyProject&sheet_name=Budget" -o budget.csv
```

```bash
# Subfolder path example
curl "http://localhost:8082/api/public/sheet/csv?project=MyProject%2FReports&sheet_name=Q1" -o q1.csv
```

---

#### `GET /api/public/sheet/audit`

Download the **audit log** of a sheet as a CSV file — showing all recorded changes (who changed what, when, and what the old/new values were).

**Query parameters:**

| Parameter | Required | Description |
|---|---|---|
| `project` | Yes | The project name. |
| `sheet_name` | Yes | The name of the sheet. |

**Timeline-aware filtering:** If the project has a **Timeline** (milestones), only audit entries **after the most recent milestone** are returned. This makes it easy to see what has changed since the last checkpoint. If no timeline exists, all audit entries are returned.

**Response:** A `text/csv` file with the following columns:

| Column | Description |
|---|---|
| `timestamp` | ISO 8601 timestamp of the change (RFC 3339). |
| `user` | Username who made the change. |
| `action` | Type of action (e.g. `EDIT`, `INSERT_ROW`, `DELETE_COL`). |
| `details` | Human-readable description of the change. |
| `row` | Starting row number (1-based). |
| `col` | Starting column label (e.g. `A`). |
| `row2` | Ending row (for range operations). |
| `col2` | Ending column (for range operations). |
| `old_value` | The previous cell value (for edits). |
| `new_value` | The new cell value (for edits). |
| `change_reversed` | `true` if this entry records an undo/revert action. |

**Example:**
```bash
curl "http://localhost:8082/api/public/sheet/audit?project=MyProject&sheet_name=Budget" -o budget_audit.csv
```

> **Tip:** Combine the timeline and audit endpoints to build automated change reports. Add a milestone when you start a review period, then poll `/api/public/sheet/audit` to get only the changes made since that milestone.

---

#### `GET /api/public/sheet/markdown`

Export a **Document** sheet as a fully-formed **Markdown file**. This endpoint is only valid for sheets of type `document`. It is useful for publishing document content to external tools, static site generators, or documentation pipelines without requiring a login.

**Query parameters:**

| Parameter | Required | Description |
|---|---|---|
| `project` | Yes | The project name (use `/` for subfolders). |
| `sheet_name` | Yes | The name of the document sheet to export. |

**How the document structure maps to Markdown:**

| Column | Markdown role |
|---|---|
| **A** – Section number | Prepended to the heading (e.g. `## 1.2 Introduction`) |
| **B** – Title | Heading text. Heading level follows the row's depth in the tree (H2 at root, H3 one level deep, etc., capped at H6). |
| **C** – Content | Markdown body rendered verbatim beneath the heading. Supports all GFM syntax already stored in the cell. |
| Any other columns | Rendered as a small two-column table (`Column \| Value`) beneath the content. |

The document's **tree hierarchy** (parent/child rows) is fully preserved — child rows become deeper headings. The sheet name is used as the top-level H1 title.

**Response:** A `text/markdown` file download.

**Example:**
```bash
curl "http://localhost:8082/api/public/sheet/markdown?project=MyProject&sheet_name=Requirements" \
     -o requirements.md
```

```bash
# Subfolder example
curl "http://localhost:8082/api/public/sheet/markdown?project=MyProject%2FSpecs&sheet_name=Design" \
     -o design.md
```

> **Note:** If the sheet is not of a document type, the endpoint returns `400 Bad Request`.

---

### Assets & Files

- **Image Assets:** Upload images (drag-and-drop supported) to a project's asset library. Insert them into markdown content or reference their URLs.
- **Python Files:** Upload Python scripts and other files to a shared directory. Reference or link them in your documents.

---

### Help & Documentation

A built-in **Help** page is available in the application that displays this documentation rendered as a web page. Access it from the **Help** button in the top navigation bar on the Projects page. No login is required to read the help content — the documentation is served by the backend at `/api/public/readme` and rendered with full markdown formatting in the browser.

---

## User Roles & Permissions

| Role | Capabilities |
|---|---|
| **Site Admin** | Full access: manage all users, projects, and sheets. Configure LLM settings. Transfer ownership. Download backups. View integrity reports. |
| **Regular User (with Create permission)** | Create projects, create sheets within their own or administered projects, edit sheets they have permission for. |
| **Regular User (without Create permission)** | Can only edit sheets where they have been granted editor access. Cannot create projects. |
| **Project Owner** | Full control over the project: create/delete/rename sheets, manage project admins, manage editors. |
| **Project Admin** | Editor-level access to all sheets within the project. Can manage sheets. |
| **Sheet Owner** | Full control over the sheet: manage editor permissions, transfer ownership, change cell types, manage scripts. |
| **Sheet Editor** | Can edit cell values in the sheet. Cannot change cell types or manage scripts (owner-only). |

### Permission Hierarchy

```
Site Admin
  └── Can manage everything
Project Owner
  └── Full control over their project and its sheets
Project Admin
  └── Can edit all sheets in the project
Sheet Owner
  └── Full control over their sheet (types, scripts, permissions)
Sheet Editor
  └── Can edit cell values only
```

---

## Administration

Access the **Admin Panel** from the top navigation (admin users only):

- **User Management:** List all users, toggle project creation permission, reset passwords.
- **Project Transfer:** Transfer ownership of any project to another user.
- **Sheet Transfer:** Transfer ownership of any sheet to another user.
- **LLM Configuration:** Set the URL for the OpenAI-compatible LLM endpoint used by AI cells.
- **Integrity Report:** View the integrity status of all data files (intact/corrupt).
- **Backup:** Download a full ZIP backup of all application data.

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                          Users / Browsers                          │
└────────────────────────────┬────────────────────────────────────────┘
                             │  HTTP + WebSocket
                             ▼
┌─────────────────────────────────────────────────────────────────────┐
│                     Frontend (React + Vite)                        │
│                                                                     │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌───────────────────┐  │
│  │  Login   │  │ Projects │  │Dashboard │  │ DataSheet/Document│  │
│  └──────────┘  └──────────┘  └──────────┘  └───────────────────┘  │
│                                                  │                  │
│                              ┌────────────────────┤                  │
│                              ▼                    ▼                  │
│                    ┌──────────────┐   ┌─────────────────────┐      │
│                    │ Script/AI/MD │   │   Chat & Timeline   │      │
│                    │   Editors    │   │     Sidebars        │      │
│                    └──────────────┘   └─────────────────────┘      │
└────────────────────────────┬────────────────────────────────────────┘
                             │
              ┌──────────────┴──────────────┐
              │ REST API         WebSocket   │
              ▼                      ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      Backend (Go Server)                           │
│                                                                     │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │                    WebSocket Hub                              │  │
│  │  • Room-per-sheet model (one room per active sheet)          │  │
│  │  • Broadcasts edits to all connected clients in real time    │  │
│  │  • Handles cell updates, locking, row/col operations, chat  │  │
│  └──────────────────────────────────────────────────────────────┘  │
│                                                                     │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────────┐     │
│  │ Sheet Manager│  │ User Manager │  │   Chat Manager       │     │
│  │ • In-memory  │  │ • Auth/Login │  │   • Messages         │     │
│  │   sheet data │  │ • Sessions   │  │   • Read tracking    │     │
│  │ • Cell ops   │  │ • Permissions│  │   • Broadcast/DM     │     │
│  └──────┬───────┘  └──────────────┘  └──────────────────────┘     │
│         │                                                           │
│  ┌──────┴──────────────────────────────────────────────────────┐   │
│  │              Dependency & Execution Engine                   │   │
│  │                                                              │   │
│  │  Cell Changed ──► Find Dependent Scripts/AI Cells            │   │
│  │       ──► Execute Scripts (Python) / AI Prompts (LLM)        │   │
│  │       ──► Write Results Back ──► Trigger Cascading Deps      │   │
│  │                                                              │   │
│  │  • Dependency DAG tracking                                   │   │
│  │  • Auto-adjusting references on row/col insert/delete        │   │
│  │  • Cross-sheet and cross-project reference resolution        │   │
│  └──────────────────────────────────────────────────────────────┘   │
│                                                                     │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────────┐     │
│  │  Integrity   │  │ Project Meta │  │   Audit Logger       │     │
│  │  • SHA-256   │  │ & Audit      │  │   • Per-cell history │     │
│  │  • Checksums │  │ • Ownership  │  │   • Diff tracking    │     │
│  └──────────────┘  └──────────────┘  └──────────────────────┘     │
└────────────────────────────┬────────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    File-Based Storage (DATA/)                      │
│                                                                     │
│  DATA/                                                              │
│  ├── users.json + users.json.shasum                                │
│  ├── chat.json + chat.json.shasum                                  │
│  ├── project_meta.json + project_meta.json.shasum                  │
│  ├── llm_settings.json                                             │
│  ├── ProjectName/                                                   │
│  │   ├── project_audit.json                                        │
│  │   ├── SheetName.json + SheetName.json.shasum                    │
│  │   ├── Subfolder/                                                │
│  │   │   └── AnotherSheet.json + AnotherSheet.json.shasum          │
│  │   └── assets/                                                    │
│  │       └── image.png                                              │
│  └── ...                                                            │
└─────────────────────────────────────────────────────────────────────┘
```

### How Real-Time Collaboration Works

1. When a user opens a sheet, the browser establishes a **WebSocket connection** and joins a **room** for that sheet.
2. The server sends the full sheet data and chat history to the client.
3. When any user edits a cell, the change is sent via WebSocket to the **Hub**.
4. The Hub **validates permissions**, applies the edit to the in-memory sheet, and **broadcasts** the change to all other clients in the room.
5. If the changed cell is referenced by scripts or AI cells, the **Dependency Engine** triggers their re-execution.
6. Script/AI outputs are written back and broadcast as additional updates.
7. Sheets are **periodically saved** to disk with debounced writes to avoid excessive I/O.

### How Interdependencies Work

```
  Cell A1 changes
       │
       ▼
  Dependency Engine scans for references to A1
       │
       ├──► Script in B1 uses {{A1}} → Re-execute Python script
       │         │
       │         ▼
       │    B1 output changes → Scan for references to B1
       │         │
       │         └──► AI Cell in C1 uses {{B1}} → Re-execute LLM prompt
       │
       └──► ComboBox in D1 options from {{A1:A10}} → Refresh dropdown options
```

- References are tracked in a **dependency graph**.
- When a cell changes, all downstream dependents are found and executed.
- Cascading continues until no further changes propagate.
- Cross-sheet references (e.g., `{{OtherProject/OtherSheet/A1}}`) work seamlessly across the entire workspace.

---

## Technology Stack

| Component | Technology |
|---|---|
| **Frontend** | React 19, Vite, Bootstrap 5, Lucide Icons |
| **Backend** | Go (Golang) |
| **Real-Time Communication** | WebSocket (Gorilla WebSocket) |
| **Storage** | File-based JSON (no database required) |
| **Scripting** | Python 3 (server-side execution) |
| **AI Integration** | OpenAI-compatible LLM API |
| **Markdown** | Marked (GFM), MathJax (LaTeX) |
| **Export/Import** | Excelize (XLSX) |
| **Authentication** | bcrypt password hashing, token-based sessions |
| **Integrity** | Salted SHA-256 checksums |

---

## License

See the repository for license information.
