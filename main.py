from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
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
        # Save uploaded files temporarily
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        profile_path = None
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Generate the output file
        output_path = generate_output_file(consensus_path, profile_path)

        # Return the generated file
        return StreamingResponse(
            open(output_path, "rb"),
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

def generate_output_file(consensus_path, profile_path):
    try:
        # Create a new Excel writer
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            output_path = temp_file.name
        
        # Load consensus file
        consensus_df = pd.read_excel(consensus_path, sheet_name=None)
        
        # Load profile file if exists
        profile_dfs = {}
        if profile_path:
            profile_dfs = pd.read_excel(profile_path, sheet_name=None)
        
        # Load DCF Model
        dcf_model_df = pd.read_excel("DCF_Model.xlsx", sheet_name="DCF Model")
        
        # Create new workbook with all sheets
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write all sheets from consensus
            for sheet_name, df in consensus_df.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Write all sheets from profile
            for sheet_name, df in profile_dfs.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Write DCF Model sheet
            dcf_model_df.to_excel(writer, sheet_name="DCF Model", index=False)
        
        return output_path
        
    except Exception as e:
        raise Exception(f"Error generating output file: {str(e)}")
