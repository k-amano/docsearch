using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Arx.DocSearch.Util;
using System.Reflection;
using System.Diagnostics;
using Xyn.Util;
using System.Text;
using System.Windows.Forms.VisualStyles;
//using static System.Net.WebRequestMethods;

namespace Arx.DocSearch.Agent
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			this.logs = new List<string>();
            this.subPrograms = new List<Process>();
            InitializeComponent();
			this.timer1.Interval = 5000;
			this.mainProgram = "";
            this.fileName = "";
            this.programNo = 0;
			this.StartSubPrograms();
			this.GetProgramNo();
            this.Text = string.Format("{0}(Agent{1})", this.Text, this.programNo);
        }

		private SearchJob job;
		private const string CRLF = "\r\n";
		private List<string> logs;
		private string mainProgram;
        private string fileName;
        private List<Process> subPrograms;
        private int programNo;
		private int mainPid;
		delegate void AppendTextCallback(string text);

		private void onLoad(object sender, EventArgs e)
		{
            try
            {
                int userIndex = this.GetUserIndexFromCommandLine();
				this.mainPid = this.GetPidFromCommandLine();
                //this.CleanFolder();
                this.GetProgramNo();
                this.timer1.Start();
                this.job = new SearchJob(this, userIndex);
                if (1 != this.programNo)
                {
                    this.WindowState = FormWindowState.Minimized;
                }

            }
            catch (Exception ex)
            {
                this.WriteLog(ex.Message + ex.StackTrace);
            }

        }

		private void onFormClosing(object sender, FormClosingEventArgs e)
		{
			this.WriteErrorLog();
			this.job.Dispose();
			//this.CleanFolder();
			this.timer1.Stop();
			/*if (1 == this.programNo) {
                foreach (Process process in this.subPrograms) { process.Kill(); }
            }*/
		}

		private void onResize(object sender, EventArgs e)
		{
			this.textBox.Width = this.ClientSize.Width - 16;
			this.textBox.Height = this.ClientSize.Height - 13;
		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			this.WriteErrorLog();
			if (1 < this.programNo && !string.IsNullOrEmpty(this.mainProgram))
			{
                Process p = null;
                try
                {
                    p = Process.GetProcessById(this.mainPid);

                }
                catch { }
                if (null == p) this.Close();
            }
		}

		private int GetUserIndexFromCommandLine()
		{
			int userIndex = 1;
			string[] commandLine = System.Environment.GetCommandLineArgs();
			string paramStr1 = string.Empty;
			if (commandLine.Length > 1) paramStr1 = commandLine[1];
			if (paramStr1.StartsWith("/IndexOfUser=")) {
                if (paramStr1.Length > 14) userIndex = Convert.ToInt32(paramStr1.Substring(13, 2));
                if (userIndex < 1 || userIndex > 16) userIndex = 1;
            }
			return userIndex;
		}

        private int GetPidFromCommandLine()
        {
            int pid = 1;
            string[] commandLine = System.Environment.GetCommandLineArgs();
            string paramStr1 = string.Empty;
            if (commandLine.Length > 1) paramStr1 = commandLine[1];
            if (paramStr1.StartsWith("/pid="))
            {
                if (5 < paramStr1.Length) pid = ConvertEx.GetInt(paramStr1.Substring(5));
            }
            return pid;
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
				string text = string.Format("[{0:yyyyMMdd HHmmss FFF}] {1}{2}", DateTime.Now, Log, CRLF);
				this.textBox.AppendText(text);
				this.logs.Add(text);
				if (this.textBox.Lines.Length > 400)
				{
					this.textBox.Select(0, this.textBox.Lines[0].Length + 1);
					this.textBox.SelectedText = string.Empty;
				}
			}
		}

		private void WriteErrorLog()
		{
			if (0 == this.logs.Count) return;
			string Log = string.Join("\r\n", this.logs.ToArray());
			string appName = Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location);
			string path = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "log_" + appName);
			ErrorLog.Instance.WriteErrorLog(path, Log);
			this.logs.Clear();
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

		private void StartSubPrograms(){ 
			string pid = ConvertEx.GetString(Process.GetCurrentProcess().Id);
            var assembly = System.Reflection.Assembly.GetEntryAssembly();
			// Get the full path of the assembly
			string filePath = assembly.Location;
			string dir = Path.GetDirectoryName(filePath);
			this.fileName = Path.GetFileNameWithoutExtension(filePath);
			string[] parts = this.fileName.Split('_');
			if (1 < parts.Length) {
				int len = parts.Length;
				this.programNo  = ConvertEx.GetInt(parts[len - 2]);
                int total = ConvertEx.GetInt(parts[len - 1]);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < len - 2; i++)
                {
                    sb.Append(parts[i]);
                }
                this.mainProgram = string.Format("{0}_1_{1}.exe", sb.ToString(), total);
				List<string> files = new List<string>();
                if (1 == programNo && programNo < total)
                {
                    for (int i = 2; i <= total; i++)
                    {
                        string pname = Path.Combine(dir, string.Format("{0}_{1}_{2}.exe", sb.ToString(), i, total));
                        if (File.Exists(pname)) files.Add(pname);
                        else 
                        {
                            MessageBox.Show(string.Format("プログラム'{0}'が存在しません。", pname),
                                "エラー",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return;
                        }
                    }
					foreach(string file in files) {
                        Process p = Process.Start(file, "/pid=" + pid);
						this.subPrograms.Add(p);
                    }
                }
            }
		}
		private void GetProgramNo()
		{
			var assembly = System.Reflection.Assembly.GetEntryAssembly();
			// Get the full path of the assembly
			string filePath = assembly.Location;
			string dir = Path.GetDirectoryName(filePath);
			this.fileName = Path.GetFileNameWithoutExtension(filePath);
			string[] parts = this.fileName.Split('_');
			if (1 < parts.Length)
			{
				int len = parts.Length;
				this.programNo = ConvertEx.GetInt(parts[len - 2]);
			}
		}


        /// <summary>
        /// 指定した実行ファイル名のプロセスをすべて取得する。
        /// </summary>
        /// <param name="searchFileName">検索する実行ファイル名。</param>
        /// <returns>プロセスが存在の有無。</returns>
        private Process GetProcessByFileName(string searchFileName)
        {
            searchFileName = searchFileName.ToLower();
            //すべてのプロセスを列挙する
            foreach (Process p in Process.GetProcesses())
            {
                string fileName;
                try
                {
                    //メインモジュールのパスを取得する
                    fileName = p.MainModule.FileName;
                }
                catch (System.ComponentModel.Win32Exception)
                {
                    //MainModuleの取得に失敗
                    fileName = "";
                }
                if (0 < fileName.Length)
                {
                    //ファイル名の部分を取得する
                    fileName = System.IO.Path.GetFileName(fileName);
                    //探しているファイル名と一致した時、真を返す
                    if (searchFileName.Equals(fileName.ToLower()))
                    {
                        return p;
                    }
                }
            }
            return null;
        }
    }
}
