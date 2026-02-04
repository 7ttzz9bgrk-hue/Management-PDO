# Excel Sheet Viewer

A simple web app that displays Excel data as a browsable interface with real-time updates and inline editing.

## Features

- Reads multiple Excel files
- Combines sheets with the same name across files
- **Real-time auto-refresh** - Browser updates automatically when Excel files are saved (no manual refresh needed)
- **Inline editing** - Edit task data directly in the browser and save back to Excel
- **Add new columns** - Add new fields by typing `NewColumn: value` in the editor
- **Due Soon popup** - View all tasks with upcoming deadlines in a convenient popup window
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

### Step 2: Install Dependencies

Dependencies are additional code packages that this app needs to run. Open **Command Prompt** and navigate to the project folder:

```
cd path\to\this\project
```

Then run this command to install all required packages:
```
pip install -r requirements.txt
```

You'll see text scrolling as packages are downloaded and installed. Wait until it finishes.

### Step 3: Configure Your Excel Files

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

### Step 4: Add Your Excel Files

Place your Excel data file(s) in the project directory (the same folder as `main.py`).

## Run

```
python main.py
```

Open http://localhost:8889

## Folder Structure

```
project/
├── main.py              # FastAPI server with file watching
├── requirements.txt
├── mockData.xlsx        # Your Excel data (gitignored)
├── templates/
│   └── index.html       # HTML template with Jinja2
└── static/
    ├── css/
    │   └── styles.css   # Custom animations and styles
    └── js/
        └── app.js       # Application logic
```

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main web interface |
| `/api/data` | GET | Fetch all sheets data as JSON |
| `/api/save-task` | POST | Save task changes to Excel |
| `/events` | GET | SSE endpoint for real-time updates |
