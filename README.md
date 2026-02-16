# Excel Sheet Viewer

A simple web app that displays Excel data as a browsable interface with real-time updates and inline editing.

## Features

- Reads multiple Excel files
- Combines sheets with the same name across files
- **Real-time auto-refresh** - Browser updates automatically when Excel files are saved (no manual refresh needed)
- **Inline editing** - Edit task data directly in the browser and save back to Excel
- **Add new columns** - Add new fields by typing `NewColumn: value` in the editor
- **Due Soon popup** - View all tasks with upcoming deadlines in a convenient popup window
- **Open/Close Excel** - Floating button to open or close the source Excel file directly from the browser
- File watching powered by Watchdog (monitors only `.xlsx` files in `FILE_PATHS`)
- Server-Sent Events (SSE) for instant browser updates
- Skips default sheet names (Sheet1, Sheet2, etc.)
- Preserves Excel formatting including sheet tab colors when saving

## Editing Data

Click the **Edit** button on any task to modify its data:

1. A textarea appears with the current `Key: Value` format
2. Modify existing values by changing the text after the colon
3. Add new columns by adding a new line like `NewColumn: some value`
4. Click **Save Changes** to write back to the Excel file
5. The UI auto-refreshes after saving

**Notes:**
- Changes are saved to the original source Excel file
- New columns are added to the Excel sheet for all rows (empty for other rows)
- If the Excel file was modified externally, you'll be prompted to refresh

## Customization

**All customization is done in the Excel file itself, not in the code.**

- **Add more sheets** → Each sheet becomes a separate tab
- **Add more rows** → Each row becomes a task
- **Add/rename columns** → Columns are displayed automatically
- **First column** → Used as the button/card label in the UI
- **Column names are flexible** → Use any column names you want. The app will display whatever columns you create, giving you full freedom to structure your data however you need.

### Recommended Columns

For the best experience with filtering and status badges, consider including:
- `Deadline` - Enables due date display
- `Priority` - Enables priority indicators (High/Medium/Low)
- `Status` - Enables status badges (Not Started/In Progress/Completed/Blocked)

These are optional - the app works with any columns you define.

## Setup

### Step 1: Install Python

Python is a programming language that this app is built with. You need to install it first before running the app.

1. Go to the official Python website: https://www.python.org/downloads/
2. Click the yellow **"Download Python"** button (download version 3.10 or newer)
3. Run the downloaded installer
4. **IMPORTANT:** On the first screen of the installer, check the box that says **"Add Python to PATH"** (this is crucial!)
5. Click **"Install Now"**
6. Wait for the installation to complete, then click **"Close"**

To verify Python is installed correctly, open **Command Prompt** (search "cmd" in Windows Start menu) and type:
```
python --version
```
You should see something like `Python 3.12.0`. If you see an error, try restarting your computer and running the command again.

### Step 2: Download This Project

You have two options to download this project:

#### Option A: Download as ZIP (Easiest)

1. On the GitHub page for this project, click the green **"Code"** button
2. Click **"Download ZIP"**
3. Extract the ZIP file to a folder on your computer (e.g., `C:\Users\YourName\Documents\ExcelViewer`)
4. Remember this folder location - you'll need it later

#### Option B: Using Git (Recommended for developers)

Git is a version control tool that lets you download and keep projects up to date.

**Install Git:**
1. Go to https://git-scm.com/downloads
2. Download the installer for Windows
3. Run the installer and follow the prompts (default options are fine)
4. Restart Command Prompt after installation

**Create a GitHub account (optional but recommended):**
1. Go to https://github.com/
2. Click **"Sign up"** and follow the steps to create a free account
3. This allows you to save your own projects and contribute to others

**Clone (download) the project:**
1. Open **Command Prompt**
2. Navigate to where you want to save the project:
   ```
   cd C:\Users\YourName\Documents
   ```
3. Clone the repository:
   ```
   git clone https://github.com/7ttzz9bgrk-hue/Management-PDO.git
   ```
4. A new folder called `Management-PDO` will be created with all the project files

### Step 3: Install Dependencies

Dependencies are additional code packages that this app needs to run. Open **Command Prompt** and navigate to the project folder:

```
cd path\to\this\project
```

Then run this command to install all required packages:
```
pip install -r requirements.txt
```

You'll see text scrolling as packages are downloaded and installed. Wait until it finishes.

### Step 4: Configure Your Excel Files

Open `main.py` in any text editor (Notepad works fine) and find the `FILE_PATHS` section near the top. Add the paths to your Excel files:

```python
FILE_PATHS = [
    "mockData.xlsx",
    "//server/shared/projects.xlsx",
]
```

- Use forward slashes `/` or double backslashes `\\` in file paths
- You can add as many Excel files as you need
- Network paths like `//server/shared/file.xlsx` are supported

### Step 5: Add Your Excel Files

Place your Excel data file(s) in the project directory (the same folder as `main.py`).

## Run

```
python main.py
```

Open http://localhost:8889

## Folder Structure

```
project/
├── main.py                      # Application entry point (runs Uvicorn server)
├── requirements.txt             # Python dependencies
├── mockData.xlsx                # Sample Excel data (gitignored)
├── app/                         # Main application package
│   ├── __init__.py              # FastAPI app creation, startup/shutdown events
│   ├── config.py                # Configuration (file paths, ports, debounce settings)
│   ├── models.py                # Pydantic data models (TaskUpdate, ExcelFileRequest)
│   ├── state.py                 # Global state (cached data, version, connected clients)
│   ├── routes/                  # API route handlers
│   │   ├── __init__.py          # Route registration
│   │   ├── pages.py             # HTML page serving (GET /)
│   │   ├── data.py              # Data fetch and save endpoints
│   │   ├── events.py            # Server-Sent Events for real-time updates
│   │   └── excel.py             # Excel file open/close operations
│   └── services/                # Business logic layer
│       ├── __init__.py
│       ├── data_loader.py       # Excel data loading, parsing, and reload logic
│       ├── excel_io.py          # Shared-access file reading (Windows compatible)
│       ├── excel_manager.py     # System-level Excel open/close
│       └── file_watcher.py      # Watchdog-based file change monitoring
├── templates/
│   └── index.html               # Main web interface (Jinja2 + Tailwind CSS)
└── static/
    ├── css/
    │   └── styles.css           # Custom animations, dark theme, component styles
    └── js/
        └── app.js               # Frontend logic (filtering, editing, SSE, Due Soon modal)
```

## Use Case: Team Workload Management

This application is designed for teams where individual employees manage their own workload in Excel and supervisors need visibility across the team.

### How It Works

**Step 1 - Employee manages their own tasks:**

An employee (e.g. a junior engineer) maintains their personal Excel spreadsheet on a shared network drive. They log their tasks, set priorities, update statuses, and track deadlines - all using familiar Excel columns. For example:

| Task | Project | Status | Priority | Deadline | Assignee |
|------|---------|--------|----------|----------|----------|
| Review P&ID drawings | Plant Upgrade | In Progress | High | 2026-03-01 | John |
| Submit RFI for valve specs | Maintenance Turnaround | Not Started | Medium | 2026-03-15 | John |
| Update hazard register | Safety Review | Completed | Low | 2026-02-10 | John |

**Step 2 - Employee shares their spreadsheet location:**

The employee provides their manager or supervisor with the file path to their Excel spreadsheet on the shared drive, e.g.:

```
//FILESERVER01/SharedDrive/Teams/Engineering/John_Tasks.xlsx
```

**Step 3 - Manager configures the viewer to monitor multiple people:**

The manager or supervisor adds each team member's spreadsheet path into `app/config.py`:

```python
FILE_PATHS = [
    "//FILESERVER01/SharedDrive/Teams/Engineering/John_Tasks.xlsx",
    "//FILESERVER01/SharedDrive/Teams/Engineering/Sarah_Tasks.xlsx",
    "//FILESERVER01/SharedDrive/Teams/Engineering/Mike_Tasks.xlsx",
    "//FILESERVER01/SharedDrive/Teams/Engineering/Lisa_Tasks.xlsx",
]
```

**Step 4 - Manager gets a unified overview:**

The manager launches the web app and instantly sees an aggregated view of everyone's tasks across all projects. They can:

- **Switch between team members** using the sheet tabs (sheets with the same name across files are merged, giving a combined project view)
- **Filter by status** to see what's blocked or not started
- **Filter by priority** to focus on urgent items
- **Use the Due Soon popup** to see all upcoming deadlines across every team member and project in one place
- **Edit tasks directly** in the browser if corrections are needed, which writes changes back to the source Excel file

This gives supervisors a live dashboard of team workload without requiring employees to learn new software - everyone continues working in Excel.

## Future Functionality

### Performance Analytics and Charting

Planned features for team-level performance visibility:

- **Tasks completed over time** - Plot the number of tasks each employee has completed on a weekly, daily, or monthly basis. Compare employees side-by-side to identify workload imbalances or recognise high performers.
- **Project distribution pie charts** - Visualise which projects are currently being worked on across the team as a pie chart, showing the proportion of active tasks per project.
- **Workload heatmaps** - Identify periods of high and low activity across the team to support better resource planning.

### Priority and Deadline Reminder Popup

A small HTML popup notification will appear in the corner of the screen to constantly remind users of high-priority or approaching-deadline tasks. This will function similarly to a previously built **Lesson Learnt** program, which cycled through random lesson-learnt entries using a custom HTML overlay with animations and timed transitions. The reminder popup will:

- **Persist in the bottom corner** of the browser (or a configurable position) without blocking the main interface
- **Cycle through high-priority and near-deadline tasks** at regular intervals, drawing attention to items that need action
- **Use styled HTML with animations** (fade-in/out, slide transitions) to keep reminders visible but non-intrusive, similar to the rotating lesson-learnt display
- **Highlight urgency visually** with colour-coded borders or backgrounds (red for overdue, amber for due soon, etc.)
- **Be dismissible or snooze-able** so users can temporarily acknowledge a reminder and return focus to their current work

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main web interface |
| `/api/data` | GET | Fetch all sheets data as JSON |
| `/api/save-task` | POST | Save task changes to Excel |
| `/api/open-excel` | POST | Open an Excel file with the system default app |
| `/api/close-excel` | POST | Close a previously opened Excel file |
| `/api/excel-status` | GET | Get open/close status of tracked Excel files |
| `/events` | GET | SSE endpoint for real-time updates |
