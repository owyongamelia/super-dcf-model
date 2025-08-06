from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from copy import copy
import shutil
import os
from tempfile import NamedTemporaryFile

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded files
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        profile_path = None
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        output_path = generate_output_file(consensus_path, profile_path)

        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        for path in [consensus_path, profile_path]:
            if path and os.path.exists(path):
                os.remove(path)

def copy_sheet(source_sheet, target_sheet):
    # Merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_range))

    # Row dimensions
    for idx, dim in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[idx] = copy(dim)

    # Column dimensions
    for col_letter, dim in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col_letter] = copy(dim)

    # Cell values and styles
    for row in source_sheet.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            new_cell = target_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

def generate_output_file(consensus_path, profile_path):
    # Load template workbook
    template_wb = load_workbook("DCF Model.xlsx")
    output_wb = Workbook()
    output_wb.remove(output_wb.active)  # Remove default sheet

    # Copy DCF Model sheet
    dcf_sheet = template_wb["DCF Model"]
    new_dcf = output_wb.create_sheet("DCF Model")
    copy_sheet(dcf_sheet, new_dcf)

    # Copy Consensus sheet
    cons_wb = load_workbook(consensus_path, data_only=False)
    if "Consensus" in cons_wb.sheetnames:
        cons_sheet = cons_wb["Consensus"]
        new_cons = output_wb.create_sheet("Consensus")
        copy_sheet(cons_sheet, new_cons)
    else:
        raise Exception("Consensus sheet not found.")

    # Copy Public Company sheet if available
    if profile_path:
        prof_wb = load_workbook(profile_path, data_only=False)
        if "Public Company" in prof_wb.sheetnames:
            pub_sheet = prof_wb["Public Company"]
            new_pub = output_wb.create_sheet("Public Company")
            copy_sheet(pub_sheet, new_pub)

    # Save output
    temp = NamedTemporaryFile(delete=False, suffix=".xlsx")
    output_wb.save(temp.name)
