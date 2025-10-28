using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace OutlookMailMonitor
{
    class Program
    {
        private static Application outlookApp;
        private static NameSpace outlookNamespace;
        private static MAPIFolder inboxFolder;
        private static Items inboxItems;
        private static readonly HttpClient httpClient = new HttpClient();
        
        // FastAPIサーバーのベースURL（環境に合わせて変更してね）
        private const string FASTAPI_BASE_URL = "http://localhost:8000";

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            Console.WriteLine("🚀 Outlook→FastAPI 連携プログラム起動中...");
            Console.WriteLine("メール受信したらチケット番号抽出してAPIたたくよ！");
            Console.WriteLine($"📡 接続先: {FASTAPI_BASE_URL}");
            Console.WriteLine("終了するには Ctrl+C を押してね～\n");

            try
            {
                // Outlookアプリケーションに接続
                outlookApp = new Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
                
                Console.WriteLine("✅ Outlookに接続成功！");
                
                // 受信トレイを取得
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                inboxItems = inboxFolder.Items;
                
                Console.WriteLine("📬 受信トレイを監視中...");
                Console.WriteLine("💡 準備完了！メール待ってるよ～\n");
                
                // イベントハンドラーを登録
                inboxItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(OnNewMailReceived);
                
                // Ctrl+C のハンドラー
                Console.CancelKeyPress += (sender, e) =>
                {
                    e.Cancel = true;
                    Console.WriteLine("\n👋 終了するね～");
                    Cleanup();
                    Environment.Exit(0);
                };
                
                // プログラムを実行し続ける
                Console.WriteLine("待機中... (何かキーを押すと終了するよ)");
                Console.ReadLine();
                
                Cleanup();
            }
            catch (COMException ex)
            {
                Console.WriteLine($"❌ COM エラー発生: {ex.Message}");
                Console.WriteLine("Outlookが起動してるか確認してね！");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"❌ エラー発生: {ex.Message}");
            }
            
            Console.WriteLine("\n✨ お疲れ様でした～！");
        }

        // 新着メールのイベントハンドラー
        private static void OnNewMailReceived(object item)
        {
            try
            {
                MailItem mailItem = item as MailItem;
                if (mailItem != null)
                {
                    Console.WriteLine("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                    Console.WriteLine("🎉 新しいメールきたよー！");
                    Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                    
                    string subject = mailItem.Subject ?? "";
                    Console.WriteLine($"📧 差出人: {mailItem.SenderName}");
                    Console.WriteLine($"📌 件名: {subject}");
                    Console.WriteLine($"📅 受信日時: {mailItem.ReceivedTime}");
                    
                    // 件名からチケット番号を抽出
                    string ticketNumber = ExtractTicketNumber(subject);
                    
                    if (!string.IsNullOrEmpty(ticketNumber))
                    {
                        Console.WriteLine($"🎫 チケット番号発見: {ticketNumber}");
                        Console.WriteLine($"📡 FastAPIにリクエスト送信中...");
                        
                        // FastAPIを非同期で呼び出し
                        CallFastApiAsync(ticketNumber).Wait();
                    }
                    else
                    {
                        Console.WriteLine("⚠️  チケット番号が見つからなかったよ...");
                        Console.WriteLine("   件名に以下のパターンが含まれてるか確認してね：");
                        Console.WriteLine("   - [#12345] や (#12345)");
                        Console.WriteLine("   - TICKET-12345 や TKT-12345");
                        Console.WriteLine("   - INC12345 や REQ12345");
                    }
                    
                    Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
                    
                    // メールオブジェクトを解放
                    Marshal.ReleaseComObject(mailItem);
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"❌ メール処理中にエラー: {ex.Message}");
            }
        }

        // 件名からチケット番号を抽出する
        // 複数のパターンに対応してるよ！
        private static string ExtractTicketNumber(string subject)
        {
            if (string.IsNullOrEmpty(subject))
                return null;

            // パターン1: [#12345] や (#12345) のような形式
            Match match = Regex.Match(subject, @"[(\[](#?\s*(\d{4,})[)\]]");
            if (match.Success)
            {
                return match.Groups[2].Value;
            }

            // パターン2: TICKET-12345, TKT-12345, INC-12345 のような形式
            match = Regex.Match(subject, @"\b(TICKET|TKT|INC|REQ|CASE|SR|CHG)-?(\d{4,})\b", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[2].Value;
            }

            // パターン3: #12345 のような形式（単独）
            match = Regex.Match(subject, @"#(\d{4,})\b");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // パターン4: 4桁以上の数字（最後の手段）
            match = Regex.Match(subject, @"\b(\d{4,})\b");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            return null;
        }

        // FastAPIを呼び出す
        private static async Task CallFastApiAsync(string ticketNumber)
        {
            try
            {
                string url = $"{FASTAPI_BASE_URL}/test/{ticketNumber}";
                Console.WriteLine($"🌐 URL: {url}");
                
                // GETリクエストを送信
                HttpResponseMessage response = await httpClient.GetAsync(url);
                
                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"✅ API呼び出し成功！");
                    Console.WriteLine($"📥 レスポンス: {responseBody}");
                }
                else
                {
                    Console.WriteLine($"⚠️  API呼び出し失敗... ステータスコード: {response.StatusCode}");
                    string errorBody = await response.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(errorBody))
                    {
                        Console.WriteLine($"📥 エラー内容: {errorBody}");
                    }
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"❌ HTTP通信エラー: {ex.Message}");
                Console.WriteLine("   FastAPIサーバーが起動してるか確認してね！");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"❌ エラー発生: {ex.Message}");
            }
        }

        // クリーンアップ処理
        private static void Cleanup()
        {
            try
            {
                if (inboxItems != null)
                {
                    Marshal.ReleaseComObject(inboxItems);
                }
                if (inboxFolder != null)
                {
                    Marshal.ReleaseComObject(inboxFolder);
                }
                if (outlookNamespace != null)
                {
                    Marshal.ReleaseComObject(outlookNamespace);
                }
                if (outlookApp != null)
                {
                    Marshal.ReleaseComObject(outlookApp);
                }
                httpClient?.Dispose();
            }
            catch
            {
                // エラーは無視
            }
        }
    }
}
