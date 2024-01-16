using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Arx.DocSearch.Client
{
	static class Program
	{
        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
		static void Main()
		{
            //ミューテックス作成
            Mutex app_mutex = new Mutex(false, "Arx_DocSearch_Client");
            //ミューテックスの所有権を要求する
            if (app_mutex.WaitOne(0, false) == false)
            {
                MessageBox.Show("このアプリケーションは複数起動できません。");
                return;
            }
            // Console表示
            AllocConsole();
            // コンソールとstdoutの紐づけを行う。無くても初回は出力できるが、表示、非表示を繰り返すとエラーになる。
            Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
            // ThreadExceptionイベント・ハンドラを登録する
            Application.ThreadException += new
              ThreadExceptionEventHandler(Application_ThreadException);

            // UnhandledExceptionイベント・ハンドラを登録する
            Thread.GetDomain().UnhandledException += new
              UnhandledExceptionEventHandler(Application_UnhandledException);

            // メイン・スレッド以外の例外はUnhandledExceptionでハンドル
            //string buffer = "1"; char error = buffer[2];
            Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}

        // 未処理例外をキャッチするイベント・ハンドラ
        // （Windowsアプリケーション用）
        public static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            ShowErrorMessage(e.Exception, "Application_ThreadExceptionによる例外通知です。");
        }

        // 未処理例外をキャッチするイベント・ハンドラ
        // （主にコンソール・アプリケーション用）
        public static void Application_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            if (ex != null)
            {
                ShowErrorMessage(ex, "Application_UnhandledExceptionによる例外通知です。");
            }
        }

        // ユーザー・フレンドリなダイアログを表示するメソッド
        public static void ShowErrorMessage(Exception ex, string message)
        {
            MessageBox.Show(message + " n――――――――nn" +
              "エラーが発生しました。開発元にお知らせくださいnn" +
              "【エラー内容】n" + ex.Message + "nn" +
              "【スタックトレース】n" + ex.StackTrace);
        }
    }
}
