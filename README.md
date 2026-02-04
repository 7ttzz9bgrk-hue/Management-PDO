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
- **First column** → Always used as the task name
- **Column names are flexible** → Use any column names you want. The app will display whatever columns you create, giving you full freedom to structure your data however you need.

### Recommended Columns

For the best experience with filtering and status badges, consider including:
- `Deadline` - Enables due date display
- `Priority` - Enables priority indicators (High/Medium/Low)
- `Status` - Enables status badges (Not Started/In Progress/Completed/Blocked)

These are optional - the app works with any columns you define.

## Setup

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Configure your Excel files in `main.py`:
   ```python
   FILE_PATHS = [
       "mockData.xlsx",
       "//server/shared/projects.xlsx",
   ]
   ```

3. Place your Excel data file(s) in the project directory.

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
