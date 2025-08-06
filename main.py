from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
from tempfile import NamedTemporaryFile
from datetime import datetime

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name):
    """Copy a sheet with all formatting, formulas, and dimensions"""
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy column dimensions
    for col in range(1, source_sheet.max_column + 1):
        col_letter = get_column_letter(col)
        new_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
    
    # Copy row dimensions
    for row in range(1, source_sheet.max_row + 1):
        new_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy cells with values, formulas, and formatting
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(
                row=cell.row,
                column=cell.column,
                value=cell.value
            )
            
            # Preserve formulas
            if cell.data_type == 'f':
                new_cell.value = cell.value
            
            # Copy all styling attributes
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.border = cell.border.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection.copy()
                new_cell.alignment = cell.alignment.copy()
    
    return new_sheet

def update_valuation_date(sheet):
    """Update valuation date to current date"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and "Valuation Date" in str(cell.value):
                date_cell = sheet.cell(row=cell.row, column=cell.column + 2)
                date_cell.value = datetime.now().date()
                return

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = "temp_consensus.xlsx"
    profile_path = None
    
    try:
        # Save uploaded files
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)
            
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Create new workbook for output
        output_wb = load_workbook(consensus_path)
        
        # Add Corporate Profile sheets if provided
        if profile_path:
            profile_wb = load_workbook(profile_path)
            for sheet_name in profile_wb.sheetnames:
                sheet = profile_wb[sheet_name]
                new_sheet = copy_sheet(sheet, output_wb, sheet_name)
        
        # Add DCF Model sheet
        dcf_wb = load_workbook("DCF Model.xlsx")
        dcf_sheet = dcf_wb["DCF Model"]
        new_dcf_sheet = copy_sheet(dcf_sheet, output_wb, "DCF Model")
        update_valuation_date(new_dcf_sheet)
        
        # Save combined workbook
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        
        return StreamingResponse(
            open(temp_file.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        for path in [consensus_path, profile_path]:
            if path and os.path.exists(path):
                os.remove(path)
