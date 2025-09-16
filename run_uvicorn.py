import uvicorn
from main import app

if __name__ == \"__main__\":
    # プロダクション用のUvicorn設定
    uvicorn.run(
        \"main:app\",
        host=\"127.0.0.1\",
        port=8000,
        workers=4,  # ワーカープロセス数
        log_level=\"info\",
        access_log=True,
        reload=False  # プロダクションでは無効
    )
