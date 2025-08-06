from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
import copy
from tempfile import NamedTemporaryFile
from openpyxl.utils import get_column_letter

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_sheet):
    """Copy all cell values, styles, dimensions and merged cells between sheets"""
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = target_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.border = copy.copy(cell.border)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy.copy(cell.protection)
                new_cell.alignment = copy.copy(cell.alignment)

    # Copy column dimensions
    for col_letter, col_dim in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col_letter] = copy.copy(col_dim)

    # Copy row dimensions
    for row_idx, row_dim in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[row_idx] = copy.copy(row_dim)

    # Copy merged cells
    for merge_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merge_range))

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded files temporarily
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        profile_path = None
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Load base DCF model
        base_wb = load_workbook("DCF Model.xlsx", data_only=False)
        
        # Load consensus file and copy sheet
        consensus_wb = load_workbook(consensus_path, data_only=False)
        if "Consensus" not in consensus_wb.sheetnames:
            raise HTTPException(status_code=400, detail="Consensus sheet not found in uploaded file")
        
        # Create new Consensus sheet in base workbook
        if "Consensus" in base_wb.sheetnames:
            base_wb.remove(base_wb["Consensus"])
        new_consensus = base_wb.create_sheet("Consensus")
        copy_sheet(consensus_wb["Consensus"], new_consensus)
        consensus_wb.close()

        # Load profile file and copy sheet if provided
        if profile_path:
            profile_wb = load_workbook(profile_path, data_only=False)
            if "Public Company" not in profile_wb.sheetnames:
                raise HTTPException(status_code=400, detail="Public Company sheet not found in uploaded file")
            
            if "Public Company" in base_wb.sheetnames:
                base_wb.remove(base_wb["Public Company"])
            new_public = base_wb.create_sheet("Public Company")
            copy_sheet(profile_wb["Public Company"], new_public)
            profile_wb.close()

        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        base_wb.save(temp_file.name)
        base_wb.close()

        # Return the generated file
        return StreamingResponse(
            open(temp_file.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Clean up temporary files
        if os.path.exists(consensus_path):
            os.remove(consensus_path)
        if profile_path and os.path.exists(profile_path):
            os.remove(profile_path)
