from fastapi import FastAPI, WebSocket
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from ppt_controller import PPTController
import uvicorn
import asyncio
import os
import sys

def get_base_path():
    try:
        return sys._MEIPASS
    except Exception:
        return os.path.dirname(__file__)

app = FastAPI()

# Mount static files
static_dir = os.path.join(get_base_path(), "static")
# Create static dir if it doesn't exist
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")

ppt = PPTController()

class Command(BaseModel):
    action: str
    index: Optional[int] = None

@app.get("/")
def read_root():
    return FileResponse(os.path.join(static_dir, "index.html"))

@app.post("/command")
def send_command(cmd: Command):
    if cmd.action == "next":
        ppt.next_slide()
    elif cmd.action == "prev":
        ppt.prev_slide()
    elif cmd.action == "start":
        ppt.start_show()
    elif cmd.action == "stop":
        ppt.stop_show()
    elif cmd.action == "blank":
        ppt.toggle_blank()
    elif cmd.action == "goto" and cmd.index is not None:
        ppt.goto_slide(cmd.index)
    return {"status": "ok"}

@app.get("/slides_info")
def get_slides_info():
    return ppt.get_slides_info()

@app.get("/slide_thumb/{idx}")
def get_slide_thumb(idx: int):
    path = ppt.get_slide_thumb(idx)
    if path and os.path.exists(path):
        return FileResponse(path)
    return {"error": "Thumbnail not found"}

shutdown_requested = False

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await websocket.accept()
    last_sent_hash = None
    try:
        while not shutdown_requested:
            state = await asyncio.to_thread(ppt.get_state)
            
            # Only send if state changed to save network bandwidth
            current_hash = f"{state.get('is_running')}-{state.get('current_slide')}-{state.get('is_blank')}-{state.get('pres_name')}"
            
            if current_hash != last_sent_hash:
                await websocket.send_json(state)
                last_sent_hash = current_hash
                
            await asyncio.sleep(0.1)
            
        await websocket.close()
    except Exception as e:
        print("WebSocket disconnected:", e)

server_instance = None

def run_server(host="0.0.0.0", port=5432):
    global server_instance, shutdown_requested
    shutdown_requested = False
    config = uvicorn.Config(app, host=host, port=port)
    server_instance = uvicorn.Server(config)
    server_instance.run()

def stop_server():
    global server_instance, shutdown_requested
    shutdown_requested = True
    if server_instance:
        server_instance.should_exit = True
        server_instance = None
