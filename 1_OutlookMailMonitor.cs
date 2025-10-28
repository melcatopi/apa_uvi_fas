using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

namespace OutlookMailMonitor
{
    class Program
    {
        private static Application outlookApp;
        private static NameSpace outlookNamespace;
        private static MAPIFolder inboxFolder;
        private static Items inboxItems;

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            Console.WriteLine("🚀 Outlookメール監視プログラム起動中...");
            Console.WriteLine("監視してるから、メール来たら教えるね！");
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
                
                // イベントハンドラーを登録（これが真のイベントドリブン！）
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
                    
                    Console.WriteLine($"📧 差出人: {mailItem.SenderName}");
                    Console.WriteLine($"📧 メールアドレス: {mailItem.SenderEmailAddress}");
                    Console.WriteLine($"📌 件名: {mailItem.Subject}");
                    Console.WriteLine($"📅 受信日時: {mailItem.ReceivedTime}");
                    
                    // 未読/既読状態
                    Console.WriteLine($"👁️  状態: {(mailItem.UnRead ? "未読" : "既読")}");
                    
                    // 重要度
                    string importance = mailItem.Importance switch
                    {
                        OlImportance.olImportanceHigh => "⭐ 高",
                        OlImportance.olImportanceLow => "📉 低",
                        _ => "📊 普通"
                    };
                    Console.WriteLine($"⚡ 重要度: {importance}");
                    
                    // 添付ファイル
                    if (mailItem.Attachments.Count > 0)
                    {
                        Console.WriteLine($"📎 添付ファイル: {mailItem.Attachments.Count}個");
                        for (int i = 1; i <= mailItem.Attachments.Count; i++)
                        {
                            Console.WriteLine($"   - {mailItem.Attachments[i].FileName}");
                        }
                    }
                    
                    // 本文のプレビュー（最初の200文字）
                    string body = mailItem.Body;
                    if (body.Length > 200)
                    {
                        body = body.Substring(0, 200) + "...";
                    }
                    Console.WriteLine($"💌 本文プレビュー:\n{body}");
                    
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
            }
            catch
            {
                // エラーは無視
            }
        }
    }
}
