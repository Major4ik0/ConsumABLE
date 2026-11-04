from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import os

from api.roster import router as roster_router

app = FastAPI(
    title="ConsumABLE API",
    description="Автоматизация строевого учета и формирования ежедневного расхода личного состава",
    version="2.0.0"
)

# CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(roster_router)

# Mount web directory if it exists
web_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "web")
if os.path.exists(web_dir):
    app.mount("/static", StaticFiles(directory=web_dir), name="static")

@app.get("/api/health")
def health_check():
    return {"status": "ok", "app": "ConsumABLE", "version": "2.0.0"}
