import openpyxl
from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from copy import copy

def copy_sheet(source_sheet, target_sheet):
    # Copy merged ranges first
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_range))
    # Copy row heights and other row properties
    for idx, dim in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[idx] = copy(dim)
    # Copy column widths and other column properties
    for col_letter, dim in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col_letter] = copy(dim)
    # Copy cell values and styles
    for row in source_sheet.iter_rows():
        for cell in row:
            # Skip merged cell placeholders; only copy the master cell
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
    # Copy conditional formatting rules, if any
    try:
        for cf in source_sheet.conditional_formatting:
            target_sheet.conditional_formatting.add(copy(cf))
    except Exception:
        pass  # Some complex rules may not copy directly; skip them gracefully

def build_final_excel(consensus_path, profile_path, template_path, output_path):
    # Load your template (DCF Model)
    template_wb = openpyxl.load_workbook(template_path, data_only=False)
    dcf_sheet = template_wb["DCF Model"]

    # Load the consensus workbook
    cons_wb = openpyxl.load_workbook(consensus_path, data_only=False)

    # Create the output workbook and remove the default blank sheet
    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    # Copy the DCF Model sheet
    new_dcf = out_wb.create_sheet("DCF Model")
    copy_sheet(dcf_sheet, new_dcf)

    # Copy the Consensus sheet from the consensus workbook (preserve name)
    if "Consensus" in cons_wb.sheetnames:
        new_consensus = out_wb.create_sheet("Consensus")
        copy_sheet(cons_wb["Consensus"], new_consensus)
    else:
        raise ValueError("Consensus sheet not found in consensus file")

    # Copy the Public Company sheet, from consensus file or profile file if supplied
    if "Public Company" in cons_wb.sheetnames:
        new_pc = out_wb.create_sheet("Public Company")
        copy_sheet(cons_wb["Public Company"], new_pc)
    elif profile_path:
        prof_wb = openpyxl.load_workbook(profile_path, data_only=False)
        if "Public Company" in prof_wb.sheetnames:
            new_pc = out_wb.create_sheet("Public Company")
            copy_sheet(prof_wb["Public Company"], new_pc)

    # Save the merged workbook
    out_wb.save(output_path)
