from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.reader.excel import InvalidFileException
import pandas as pd
import shutil
import os
from tempfile import NamedTemporaryFile
from datetime import datetime
from typing import Optional

app = FastAPI()

def cleanup_files(*file_paths):
    """Clean up temporary files."""
    for path in file_paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError as e:
                print(f"Error removing file {path}: {e}")

def update_valuation_date(sheet):
    """Update the valuation date in the DCF model to the current date."""
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "Valuation Date" in cell.value:
                date_cell = sheet.cell(row=cell.row, column=cell.column + 2)
                date_cell.value = datetime.now().date()
                date_cell.number_format = 'YYYY-MM-DD'
                return

def load_file_content(file_path):
    """Loads file content as a DataFrame, handling both XLSX and CSV."""
    try:
        # Try to load as an XLSX file first
        wb = load_workbook(file_path)
        sheet = wb.active
        data = sheet.values
        # Create a DataFrame from the sheet data
        cols = next(data)
        data = list(data)
        df = pd.DataFrame(data, columns=cols)
        return df
    except InvalidFileException:
        # If it's not a valid XLSX, fall back to CSV
        print("Invalid XLSX file, attempting to read as CSV...")
        try:
            return pd.read_csv(file_path)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Could not read the uploaded file as a CSV: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read the uploaded file: {str(e)}")

@app.post("/upload")
async def upload(
    background_tasks: BackgroundTasks, 
    consensus: UploadFile = File(...), 
    profile: Optional[UploadFile] = File(None)
):
    consensus_path = None
    profile_path = None
    output_path = None
    
    try:
        # Save uploaded files temporarily
        consensus_path = f"temp_consensus.tmp"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)
        
        if profile and profile.filename:
            profile_path = f"temp_profile.tmp"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Load the local template file
        template_wb = load_workbook("Template.xlsx")
        
        # Create a new workbook to hold the final result
        merged_wb = Workbook()
        
        # --- Process Consensus File ---
        consensus_df = load_file_content(consensus_path)
        ws_consensus = merged_wb.create_sheet("Consensus")
        for r in dataframe_to_rows(consensus_df, index=False, header=True):
            ws_consensus.append(r)
        
        # --- Process Profile File (Optional) ---
        if profile_path and os.path.exists(profile_path):
            profile_df = load_file_content(profile_path)
            ws_profile = merged_wb.create_sheet("Public Company")
            for r in dataframe_to_rows(profile_df, index=False, header=True):
                ws_profile.append(r)

        # --- Copy DCF Model from Template, preserving formulas ---
        dcf_sheet = template_wb["DCF Model"]
        new_dcf_sheet = merged_wb.create_sheet("DCF Model Output")
        
        for row in dcf_sheet.iter_rows():
            for cell in row:
                new_cell = new_dcf_sheet.cell(
                    row=cell.row, 
                    column=cell.column
                )
                # Copy formula if it exists. This is the key fix.
                if cell.data_type == 'f':
                    new_cell.value = cell.formula
                else:
                    new_cell.value = cell.value

                # Copy styling (font, border, fill, etc.)
                if cell.has_style:
                    new_cell.font = cell.font.copy()
                    new_cell.border = cell.border.copy()
                    new_cell.fill = cell.fill.copy()
                    new_cell.number_format = cell.number_format
                    new_cell.protection = cell.protection.copy()
                    new_cell.alignment = cell.alignment.copy()
        
        # --- Finalize Workbook ---
        if "Sheet" in merged_wb.sheetnames:
            merged_wb.remove(merged_wb["Sheet"])
            
        update_valuation_date(new_dcf_sheet)
        
        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        merged_wb.save(temp_file.name)
        
        # Add cleanup to run in the background
        cleanup_files_list = [consensus_path, temp_file.name]
        if profile_path:
            cleanup_files_list.append(profile_path)
        background_tasks.add_task(cleanup_files, *cleanup_files_list)
        
        # Return the generated file using a StreamingResponse
        def file_iterator():
            with open(temp_file.name, "rb") as f:
                yield from f

        return StreamingResponse(
            file_iterator(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Merged Solution.xlsx"},
        )
    
    except Exception as e:
        cleanup_files(consensus_path, profile_path, output_path)
        raise HTTPException(status_code=500, detail=f"Error processing files: {str(e)}")
