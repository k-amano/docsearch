using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using HCInterface;
using Arx.DocSearch.Util;

namespace Arx.DocSearch.Agent
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			InitializeComponent();
		}

		private SearchJob job;
		private const string CRLF = "\r\n";
		delegate void AppendTextCallback(string text);

		private void onLoad(object sender, EventArgs e)
		{
			int userIndex = this.GetUserIndexFromCommandLine();
			this.CleanFolder();
			this.job = new SearchJob(this, userIndex);
		}

		private void onFormClosing(object sender, FormClosingEventArgs e)
		{
			this.job.Dispose();
			this.CleanFolder();
		}

		private void onResize(object sender, EventArgs e)
		{
			this.textBox.Width = this.ClientSize.Width - 16;
			this.textBox.Height = this.ClientSize.Height - 13;
		}

		private int GetUserIndexFromCommandLine()
		{
			int userIndex = 1;
			string[] commandLine = System.Environment.GetCommandLineArgs();
			string paramStr1 = string.Empty;
			if (commandLine.Length > 1) paramStr1 = commandLine[1];
			if (paramStr1.Length > 14) userIndex = Convert.ToInt32(paramStr1.Substring(13, 2));
			if (userIndex < 1 || userIndex > 16) userIndex = 1;
			return userIndex;
		}

		public void WriteLog(string Log)
		{
			// 呼び出し元のコントロールのスレッドが異なるか確認をする
			if (this.textBox.InvokeRequired)
			{
				// 同一メソッドへのコールバックを作成する
				AppendTextCallback delegateMethod = new AppendTextCallback(WriteLog);
				// コントロールの親のInvoke()メソッドを呼び出すことで、呼び出し元の
				// コントロールのスレッドでこのメソッドを実行する
				this.Invoke(delegateMethod, new object[] { Log });
			}
			else
			{
				this.textBox.AppendText(string.Format("[{0}] {1}{2}", DateTime.Now, Log, CRLF));
				if (this.textBox.Lines.Length > 400)
				{
					this.textBox.Select(0, this.textBox.Lines[0].Length + 1);
					this.textBox.SelectedText = string.Empty;
				}
			}
			string path = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "log");
			ErrorLog.Instance.WriteErrorLog(path, Log);
		}

		private void CleanFolder()
		{
			string path = Path.GetDirectoryName(Application.ExecutablePath);
			// 数字だけのディレクトリ(クライアントから配布されたドキュメントを格納)を削除
			var dirs = Directory.EnumerateDirectories(path, "*", SearchOption.AllDirectories);
			foreach (string dir in dirs)
			{
				string dirname = Path.GetFileName(dir);
				if (Regex.IsMatch(dirname, @"^[0-9]+$")) Directory.Delete(dir, true);
			}
			// テキストおよびインデックスファイル(クライアントから配布されたドキュメント)を削除
			List<string> exts = new List<string> { ".txt", ".idx" };
			string[] files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
			foreach (string file in files)
			{
				if (exts.Contains(Path.GetExtension(file).ToLower()))
				{
					File.Delete(file);
				}
			}
		}
	}
}
