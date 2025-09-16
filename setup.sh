echo \"FastAPI + Uvicorn + Apache セットアップスクリプト\"
echo \"================================================\"

# Python仮想環境の作成
echo \"Python仮想環境を作成中...\"
python3 -m venv venv
source venv/bin/activate

# 必要なパッケージのインストール
echo \"必要なパッケージをインストール中...\"
pip install -r requirements.txt

echo \"\"
echo \"セットアップ完了！🎉\"
echo \"\"
echo \"起動方法:\"
echo \"1. FastAPIアプリを直接起動 (開発用):\"
echo \"   python main.py\"
echo \"\"
echo \"2. Uvicornでプロダクション起動:\"
echo \"   python run_uvicorn.py\"
echo \"\"
echo \"3. Apacheの設定:\"
echo \"   - apache_config.conf を /etc/apache2/sites-available/ にコピー\"
echo \"   - sudo a2ensite fastapi-sample\"
echo \"   - sudo a2enmod proxy proxy_http headers\"
echo \"   - sudo systemctl reload apache2\"
echo \"\"
echo \"アクセス先:\"
echo \"- http://localhost/ (Apache経由)\"
echo \"- http://localhost:8000/ (直接アクセス)\"
echo \"- http://localhost/health (ヘルスチェック)\"
echo \"- http://localhost/api/users/1 (API例)\"
