Apache → Uvicorn → FastAPI の構成でのWebアプリケーションサンプルです。

## 構成

```
Client → Apache (リバースプロキシ) → Uvicorn (ASGIサーバー) → FastAPI (Webフレームワーク)
```

## ファイル構成

```
fastapi_sample/
├── main.py              # FastAPIメインアプリケーション
├── run_uvicorn.py       # Uvicorn起動スクリプト
├── apache_config.conf   # Apache設定ファイル
├── requirements.txt     # Python依存関係
├── setup.sh            # セットアップスクリプト
└── README.md           # このファイル
```

## セットアップ

1. セットアップスクリプトの実行
```bash
chmod +x setup.sh
./setup.sh
```

2. Apacheの設定
```bash
# 必要なモジュールの有効化
sudo a2enmod proxy proxy_http headers

# 設定ファイルのコピーと有効化
sudo cp apache_config.conf /etc/apache2/sites-available/fastapi-sample.conf
sudo a2ensite fastapi-sample

# Apache再起動
sudo systemctl reload apache2
```

## 起動方法

### 開発環境
```bash
# 仮想環境のアクティベート
source venv/bin/activate

# FastAPI直接起動（オートリロード有効）
python main.py
```

### プロダクション環境
```bash
# 仮想環境のアクティベート
source venv/bin/activate

# Uvicornでプロダクション起動
python run_uvicorn.py
```

## API エンドポイント

- `GET /` - Hello World
- `GET /health` - ヘルスチェック
- `GET /api/users/{user_id}` - ユーザー取得
- `POST /api/users` - ユーザー作成

## アクセス方法

- Apache経由: http://localhost/
- 直接アクセス: http://localhost:8000/
- API例: http://localhost/api/users/1

## 技術スタック

- **Python 3.8+**
- **FastAPI**: 高速なWeb APIフレームワーク
- **Uvicorn**: ASGI サーバー
- **Apache**: リバースプロキシサーバー

## メリット

1. **Apache**: 静的ファイル配信、SSL終端、ロードバランシング
2. **Uvicorn**: 高性能なASGI処理
3. **FastAPI**: 型安全、自動ドキュメント生成、高パフォーマンス
