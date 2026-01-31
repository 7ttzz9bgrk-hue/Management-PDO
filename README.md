# Excel Sheet Viewer

A simple web app that displays Excel data as a browsable interface.

## Features

- Reads multiple Excel files
- Combines sheets with the same name across files
- Auto-refreshes data on each page load

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

3. Place your `index.html` template in a `templates` folder.

## Run

```
python main.py
```

Open http://localhost:8000

## Folder Structure

```
project/
├── main.py
├── requirements.txt
├── mockData.xlsx
└── templates/
    └── index.html
```
