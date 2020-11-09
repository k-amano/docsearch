using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace Arx.DocSearch.Util
{
	public class ErrorLog
	{

		private ErrorLog() { }
		// マルチスレッド対応のシングルトンパターンを採用
		private static volatile ErrorLog instance;
		private static object syncRoot = new Object();
		// ロック用のインスタンス
		private static ReaderWriterLock rwl = new ReaderWriterLock();

		// シングルトンパターンのインスタンス取得メソッド
		public static ErrorLog Instance
		{
			get
			{
				if (instance == null)
				{
					lock (syncRoot)
					{
						if (instance == null)
							instance = new ErrorLog();
					}
				}
				return instance;
			}
		}

		/// <summary>
		/// エラーログを記録します。
		/// </summary>
		/// <param name="pathname">ログを格納するディレクトリ。</param>
		/// <param name="message">エラーメッセージ文。</param>
		public void WriteErrorLog(string pathname, string message)
		{
			if (!Directory.Exists(pathname)) Directory.CreateDirectory(pathname);
			string filename = Path.Combine(pathname,
			 "error" + DateTime.Now.ToString("yyyyMMdd") + ".log");
			// ここからロック
			rwl.AcquireWriterLock(Timeout.Infinite);
			// ファイルオープン
			try
			{
				using (FileStream fs = File.Open(filename, FileMode.Append))
				using (StreamWriter writer = new StreamWriter(fs))
				{
					// 1 行書き込み
					writer.WriteLine(string.Format("[{0}] {1}", DateTime.Now, message));
				}
			}
			finally
			{
				// ロック解除は finally の中で行う
				rwl.ReleaseWriterLock();
			}
		}
	}
}
