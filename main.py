from fastapi import FastAPI, Request
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
import pandas as pd

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# ===== CONFIGURATION =====
FILE_PATHS = [
    "mockData.xlsx",
    # Add more file paths here:
    # "//server/shared/projects.xlsx",
    # "another.xlsx",
]

def load_all_sheets_data():
    all_data = {}
    valid_sheet_names = []
    
    for file_path in FILE_PATHS:
        try:
            with pd.ExcelFile(file_path) as excel_file:
                sheet_names = excel_file.sheet_names
        except Exception as e:
            print(f"Error opening file '{file_path}': {e}")
            continue
        
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                if df.empty or len(df.columns) < 2:
                    print(f"Warning: Skipping empty sheet '{sheet_name}' in '{file_path}'")
                    continue
                
                # Rule 1: Stop at the first unnamed/undefined column
                valid_cols = []
                for col in df.columns:
                    col_str = str(col).strip()
                    if col_str.startswith('Unnamed') or col_str == '' or col_str == 'nan':
                        break
                    valid_cols.append(col)
                
                if len(valid_cols) < 2:
                    print(f"Warning: Skipping sheet '{sheet_name}' in '{file_path}' - not enough named columns")
                    continue
                
                df = df[valid_cols]
                
                # Rule 2: Stop at the first empty Task Name (first column) row
                task_name_col = df.columns[0]
                cut_index = None
                for i, value in enumerate(df[task_name_col]):
                    if pd.isna(value) or str(value).strip() == '':
                        cut_index = i
                        break
                
                if cut_index is not None:
                    df = df.iloc[:cut_index]
                
                if df.empty:
                    print(f"Warning: Skipping sheet '{sheet_name}' in '{file_path}' - no valid data")
                    continue
                
                # Initialize sheet if it doesn't exist
                if sheet_name not in all_data:
                    all_data[sheet_name] = {}
                    valid_sheet_names.append(sheet_name)
                
                for _, row in df.iterrows():
                    task_name = str(row[task_name_col])
                    
                    if task_name == 'nan' or not task_name.strip():
                        continue
                    
                    details = []
                    for col in df.columns[1:]:
                        value = row[col]
                        if pd.notna(value) and str(value).strip():
                            details.append(f"{col}: {value}")
                    
                    formatted_details = '\n'.join(details)
                    
                    # Merge: append to existing task or create new
                    if task_name not in all_data[sheet_name]:
                        all_data[sheet_name][task_name] = []
                    all_data[sheet_name][task_name].append(formatted_details)
                
                print(f"Loaded sheet '{sheet_name}' from '{file_path}'")
            
            except Exception as e:
                print(f"Error loading sheet '{sheet_name}' from '{file_path}': {e}")
                continue
    
    if not valid_sheet_names:
        print("Warning: No valid sheets found, creating default")
        all_data['Default'] = {
            'Sample Task': ['No data available\nPlease check your Excel file']
        }
        valid_sheet_names = ['Default']
    
    print(f"Total sheets loaded: {len(valid_sheet_names)}")
    return all_data, valid_sheet_names

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    all_sheets_data, sheet_names = load_all_sheets_data()
    
    return templates.TemplateResponse(
        "index.html", 
        {
            "request": request, 
            "all_sheets_data": all_sheets_data,
            "sheet_names": sheet_names
        }
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8889, reload=True)
