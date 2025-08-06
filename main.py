"""
FastAPI application to merge a DCF model with user‑supplied consensus and
public company worksheets.

This module exposes a single `/upload` endpoint that accepts two Excel
files: a required `consensus` workbook and an optional `profile`
workbook.  It combines the `DCF Model` sheet from the packaged
template with the `Consensus` and `Public Company` sheets from the
uploads.  The merging preserves cell values, formulas, styles,
row/column dimensions and merged ranges so that the resulting workbook
behaves like the Aspose.Cells merger.

Note: To enable file uploads, ensure the dependency ``python-multipart``
is listed in your `requirements.txt` file.
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from copy import copy
from tempfile import NamedTemporaryFile
import shutil
import os
from typing import Optional


app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: Optional[UploadFile] = File(None)):
    """Receive uploaded files, merge them, and return the combined workbook.

    Parameters
    ----------
    consensus : UploadFile
        The primary workbook containing at least a ``Consensus`` sheet and
        optionally a ``Public Company`` sheet.
    profile : UploadFile, optional
        A secondary workbook containing a ``Public Company`` sheet if it
        is not present in the consensus workbook.

    Returns
    -------
    StreamingResponse
        A streaming response carrying the merged Excel file.
    """
    # Write the uploaded files to temporary locations on disk.  This
    # allows openpyxl to open them repeatedly and simplifies cleanup.
    consensus_path = "temp_consensus.xlsx"
    with open(consensus_path, "wb") as f:
        shutil.copyfileobj(consensus.file, f)

    profile_path = None
    if profile is not None:
        profile_path = "temp_profile.xlsx"
        with open(profile_path, "wb") as f:
            shutil.copyfileobj(profile.file, f)

    try:
        output_path = generate_output_file(consensus_path, profile_path)
        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Merged_DCF_Model.xlsx"},
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
    finally:
        # Always remove the uploaded files once processing is done.
        if os.path.exists(consensus_path):
            os.remove(consensus_path)
        if profile_path and os.path.exists(profile_path):
            os.remove(profile_path)


def generate_output_file(consensus_path: str, profile_path: Optional[str]) -> str:
    """Combine the DCF template with consensus and profile worksheets.

    This function implements the logic to merge the template with the
    uploaded workbooks.  It uses openpyxl to copy sheets between
    workbooks while preserving formulas, formatting, dimensions and
    merged cells.

    Parameters
    ----------
    consensus_path : str
        Path to the consensus workbook on disk.
    profile_path : str or None
        Path to the profile workbook on disk, if supplied.

    Returns
    -------
    str
        The path to a temporary XLSX file containing the merged workbook.
    """
    # Load the DCF model template.  Using data_only=False preserves
    # formulas instead of cached values.
    template_wb = load_workbook("DCF Model.xlsx", data_only=False)
    output_wb = template_wb

    # Define a helper to copy one sheet into the output workbook.
    def copy_sheet(source_sheet, target_title: str) -> None:
        """Copy an entire worksheet, including formatting, into the output.

        Parameters
        ----------
        source_sheet : Worksheet
            The worksheet to copy from.
        target_title : str
            The name of the new sheet in the output workbook.  If a
            sheet with the same name already exists, it will be
            removed before copying.
        """
        # Remove existing sheet of the same name to avoid stale data.
        if target_title in output_wb.sheetnames:
            std = output_wb[target_title]
            output_wb.remove(std)
        # Create the sheet at the end of the workbook.
        target_sheet = output_wb.create_sheet(title=target_title)
        # Copy merged cells first so their structure is established.
        for merged_range in source_sheet.merged_cells.ranges:
            target_sheet.merge_cells(str(merged_range))
        # Copy row dimensions (height, style, etc.).  We copy the
        # dimension objects directly to preserve all attributes.
        for idx, dim in source_sheet.row_dimensions.items():
            target_sheet.row_dimensions[idx] = copy(dim)
        # Copy column dimensions similarly.
        for col_letter, dim in source_sheet.column_dimensions.items():
            target_sheet.column_dimensions[col_letter] = copy(dim)
        # Copy each cell's value and style.  Skip merged cells that
        # represent the children of a merged range, as writing to
        # MergedCell objects is not allowed.
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
                if cell.hyperlink:
                    new_cell.hyperlink = copy(cell.hyperlink)
                if cell.comment:
                    new_cell.comment = copy(cell.comment)
        # Copy conditional formatting rules if present.  Each entry in
        # source_sheet.conditional_formatting yields a rule object.  We
        # clone and add it to the target sheet.
        try:
            for cf in source_sheet.conditional_formatting:
                target_sheet.conditional_formatting.add(copy(cf))
        except Exception:
            # Conditional formatting copy may fail for complex rules; skip gracefully.
            pass

    # Open the consensus workbook and determine which sheets to import.
    consensus_wb = load_workbook(consensus_path, data_only=False)
    sheets_to_copy = []
    # Copy the Consensus sheet if present.
    if "Consensus" in consensus_wb.sheetnames:
        sheets_to_copy.append("Consensus")
    # If no separate profile is provided and the consensus workbook
    # includes a Public Company sheet, copy it from the consensus file.
    if profile_path is None and "Public Company" in consensus_wb.sheetnames:
        sheets_to_copy.append("Public Company")
    # Copy the selected sheets from the consensus workbook.
    for name in sheets_to_copy:
        copy_sheet(consensus_wb[name], name)
    # If a profile workbook was supplied, copy its Public Company sheet.
    if profile_path:
        profile_wb = load_workbook(profile_path, data_only=False)
        if "Public Company" in profile_wb.sheetnames:
            copy_sheet(profile_wb["Public Company"], "Public Company")
    # Save the merged workbook to a temporary file.  We don't set
    # delete=True because FastAPI will read this file when streaming it
    # back to the client.
    temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
    output_wb.save(temp_file.name)
    return temp_file.name
