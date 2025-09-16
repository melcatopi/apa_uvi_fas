from fastapi.responses import JSONResponse
import uvicorn

app = FastAPI(title=\"FastAPI Sample\", version=\"1.0.0\")

@app.get(\"/\")
async def root():
    return {\"message\": \"Hello World from FastAPI!\", \"status\": \"success\"}

@app.get(\"/health\")
async def health_check():
    return {\"status\": \"healthy\", \"service\": \"fastapi-sample\"}

@app.get(\"/api/users/{user_id}\")
async def get_user(user_id: int):
    return {
        \"user_id\": user_id,
        \"name\": f\"User_{user_id}\",
        \"email\": f\"user{user_id}@example.com\"
    }

@app.post(\"/api/users\")
async def create_user(user_data: dict):
    return {
        \"message\": \"User created successfully\",
        \"user\": user_data,
        \"id\": 123
    }

if __name__ == \"__main__\":
    # 開発用の直接実行
    uvicorn.run(\"main:app\", host=\"127.0.0.1\", port=8000, reload=True)
