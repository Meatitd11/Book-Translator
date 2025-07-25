from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from app.api.routes import router as api_router
from typing import List, Dict

app = FastAPI(title="Book Translation API")

# Store active WebSocket connections
class ConnectionManager:
    def __init__(self):
        self.active_connections: Dict[str, WebSocket] = {}

    async def connect(self, websocket: WebSocket, client_id: str):
        await websocket.accept()
        self.active_connections[client_id] = websocket

    def disconnect(self, client_id: str):
        self.active_connections.pop(client_id, None)

    async def send_progress(self, client_id: str, data: dict):
        if client_id in self.active_connections:
            await self.active_connections[client_id].send_json(data)

manager = ConnectionManager()

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://book-translator.com.equationofintelligence.com","https://book-translator-ui.vercel.app","http://localhost:3000", "http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"]
)

@app.websocket("/ws/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    print(f"New WebSocket connection request for client: {client_id}")
    await manager.connect(websocket, client_id)
    print(f"WebSocket connection established for client: {client_id}")
    try:
        while True:
            data = await websocket.receive_text()
            print(f"Received WebSocket message from {client_id}: {data}")
    except WebSocketDisconnect:
        print(f"WebSocket connection closed for client: {client_id}")
        manager.disconnect(client_id)

# Include API routes
app.include_router(api_router, prefix="/api")
