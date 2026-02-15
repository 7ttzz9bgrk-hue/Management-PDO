from app import create_app
from app.config import APP_HOST, APP_PORT

app = create_app()

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host=APP_HOST, port=APP_PORT, reload=False)
