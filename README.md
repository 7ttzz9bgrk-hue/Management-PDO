# Excel Sheet Viewer

A simple web app that displays Excel data as a browsable interface with real-time updates and inline editing.

## Features

- Reads multiple Excel files
- Combines sheets with the same name across files
- **Real-time auto-refresh** - Browser updates automatically when Excel files are saved (no manual refresh needed)
- **Inline editing** - Edit task data directly in the browser and save back to Excel
- **Add new columns** - Add new fields by typing `NewColumn: value` in the editor
- File watching powered by Watchdog (monitors only `.xlsx` files in `FILE_PATHS`)
- Server-Sent Events (SSE) for instant browser updates
- Skips default sheet names (Sheet1, Sheet2, etc.)

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

### Required Columns

For full functionality, include these columns:
- `Deadline`
- `Priority`
- `Status`

Other columns are optional and will display automatically.

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
