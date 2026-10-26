import os
import uuid
from typing import Dict
from fastapi import FastAPI, File, UploadFile, WebSocket, WebSocketDisconnect, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

app = FastAPI(title="Flick Presentation API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

sessions: Dict[str, dict] = {}

@app.post("/api/sessions")
async def create_session(file: UploadFile = File(...)):
    session_id = str(uuid.uuid4())[:8]
    presenter_token = str(uuid.uuid4())
    
    file_ext = file.filename.split(".")[-1].lower()
    input_filename = f"{session_id}_{file.filename}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)
    
    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)
        
    pdf_path = input_path
    
    if file_ext in ["ppt", "pptx"]:
        output_pdf_filename = f"{session_id}.pdf"
        output_pdf_path = os.path.join(UPLOAD_DIR, output_pdf_filename)
        try:
            import comtypes.client
            import pythoncom
            pythoncom.CoInitialize()
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1
            deck = powerpoint.Presentations.Open(os.path.abspath(input_path))
            deck.SaveAs(os.path.abspath(output_pdf_path), 32)
            deck.Close()
            powerpoint.Quit()
            pdf_path = output_pdf_path
        except Exception as e:
            print("PPT to PDF conversion error:", e)
            pass
    elif file_ext == "pdf":
        output_pdf_filename = f"{session_id}.pdf"
        output_pdf_path = os.path.join(UPLOAD_DIR, output_pdf_filename)
        os.rename(input_path, output_pdf_path)
        pdf_path = output_pdf_path

    sessions[session_id] = {
        "presenter_ws": None,
        "remotes": [],
        "presenter_token": presenter_token,
        "pdf_path": pdf_path
    }
    
    return {
        "sessionId": session_id,
        "presenterToken": presenter_token
    }

@app.get("/api/sessions/{session_id}/file")
async def get_session_file(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    path = sessions[session_id]["pdf_path"]
    if os.path.exists(path):
        return FileResponse(path, media_type="application/pdf")
    raise HTTPException(status_code=404, detail="File not found")

@app.websocket("/ws/presenter/{session_id}")
async def presenter_websocket(websocket: WebSocket, session_id: str, token: str = None):
    await websocket.accept()
    if session_id not in sessions:
        await websocket.send_text('{"type": "ERROR", "message": "Session not found!"}')
        await websocket.close(code=4004, reason="Session not found")
        return

    # Strict Token Validation to prevent unauthorized or multiple active access
    stored_token = sessions[session_id].get("presenter_token")
    if not token or token != stored_token:
        await websocket.send_text('{"type": "ERROR", "message": "Unauthorized access! Invalid or missing token."}')
        await websocket.close(code=4003, reason="Unauthorized")
        return

    if sessions[session_id]["presenter_ws"] is not None:
        await websocket.send_text('{"type": "ERROR", "message": "Session already active on another screen!"}')
        await websocket.close(code=4001, reason="Already connected")
        return

    sessions[session_id]["presenter_ws"] = websocket
    await broadcast_status(session_id)

    try:
        while True:
            data = await websocket.receive_text()
            for remote in sessions[session_id]["remotes"]:
                try:
                    await remote.send_text(data)
                except:
                    pass
    except WebSocketDisconnect:
        if session_id in sessions:
            sessions[session_id]["presenter_ws"] = None
            await broadcast_status(session_id)

@app.websocket("/api/sessions/{session_id}/ws")
async def remote_websocket(websocket: WebSocket, session_id: str, role: str = "remote", token: str = None):
    await websocket.accept()
    if session_id not in sessions:
        await websocket.close(code=4004, reason="Session not found")
        return

    if websocket not in sessions[session_id]["remotes"]:
        sessions[session_id]["remotes"].append(websocket)
        
    await broadcast_status(session_id)

    try:
        while True:
            data = await websocket.receive_text()
            presenter = sessions[session_id].get("presenter_ws")
            if presenter:
                try:
                    await presenter.send_text(data)
                except:
                    pass
    except WebSocketDisconnect:
        if session_id in sessions and websocket in sessions[session_id]["remotes"]:
            sessions[session_id]["remotes"].remove(websocket)
            await broadcast_status(session_id)

async def broadcast_status(session_id: str):
    if session_id not in sessions:
        return
    remotes_count = len(sessions[session_id]["remotes"])
    presenter = sessions[session_id].get("presenter_ws")
    status_msg = f'{{ "type": "STATUS", "remotesCount": {remotes_count} }}'
    if presenter:
        try:
            await presenter.send_text(status_msg)
        except:
            pass