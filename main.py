from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
import tempfile
from copy import copy

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name=None):
    """Copy a worksheet to another workbook preserving all formatting"""
    target_sheet = target_wb.create_sheet(sheet_name or source_sheet.title)
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_range))
    
    # Copy row dimensions
    for idx, dim in source_sheet.row_dimensions.items():
        new_dim = copy(dim)
        target_sheet.row_dimensions[idx] = new_dim
    
    # Copy column dimensions
    for col_letter, dim in source_sheet.column_dimensions.items():
        new_dim = copy(dim)
        target_sheet.column_dimensions[col_letter] = new_dim
    
    # Copy all cell values and styles
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = target_sheet.cell(
                row=cell.row, 
                column=cell.column, 
                value=cell.value
            )
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    
    # Copy conditional formatting
    for cf in source_sheet.conditional_formatting:
        target_sheet.conditional_formatting.add(cf)
    
    return target_sheet

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Load DCF template
        template_wb = load_workbook("DCF Model.xlsx")
        template_sheet = template_wb["DCF Model"]
        
        # Create new workbook
        output_wb = load_workbook("DCF Model.xlsx")
        output_wb.remove(output_wb.active)  # Remove default sheet
        
        # Process consensus file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_consensus:
            shutil.copyfileobj(consensus.file, tmp_consensus)
            consensus_wb = load_workbook(tmp_consensus.name)
            consensus_sheet = consensus_wb["Consensus"]
            copy_sheet(consensus_sheet, output_wb, "Consensus")
            os.unlink(tmp_consensus.name)
        
        # Process profile file if provided
        if profile:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_profile:
                shutil.copyfileobj(profile.file, tmp_profile)
                profile_wb = load_workbook(tmp_profile.name)
                public_sheet = profile_wb["Public Company"]
                copy_sheet(public_sheet, output_wb, "Public Company")
                os.unlink(tmp_profile.name)
        
        # Add DCF Model sheet
        copy_sheet(template_sheet, output_wb, "DCF Model")
        
        # Save to temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
            output_wb.save(tmp_output.name)
            tmp_path = tmp_output.name
        
        # Return the generated file
        return StreamingResponse(
            open(tmp_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except KeyError as e:
        raise HTTPException(status_code=400, detail=f"Missing required sheet: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Clean up temporary files
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
