using System;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using Xyn.Util;

namespace Arx.DocSearch
{
	static class Program
	{
		/// <summary>
		/// アプリケーションのメイン エントリ ポイントです。
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}

		/// <summary>
		/// 例外発生時の処理を指定します。
		/// </summary>
		/// <param name="sender">イベントのソース。</param>
		/// <param name="e">イベントデータの格納。</param>
		static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
		{
			string pathname = Path.Combine(FileEx.GetFileSystemPath(Environment.SpecialFolder.ApplicationData), "logs");
			WriteErrorLog(pathname, e.Exception.StackTrace);
			string filename = Path.Combine(pathname,
 "error" + DateTime.Now.ToString("yyyy-MM") + ".log");
			MessageBox.Show(string.Format(
@"ご迷惑をおかけしております。アプリケーションに障害が発生いたしました。
障害内容は以下のログファイルをご参照ください。
{0}",
filename));
			Application.Exit();
		}

		/// <summary>
		/// エラーログを記録します。
		/// </summary>
		/// <param name="Message">エラーメッセージ文。</param>
		static void WriteErrorLog(string pathname, string message)
		{
			if (!Directory.Exists(pathname)) Directory.CreateDirectory(pathname);
			FileEx.WriteErrorLog(pathname, message);
		}
	}
}
