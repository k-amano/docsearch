using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;
using System.Text;
using Arx.DocSearch.Util;


namespace Arx.DocSearch.SpecialChars
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			InitializeComponent();
			this.word = null;
		}

		private Application word;
		// ロック用のインスタンス
		private static ReaderWriterLock rwl = new ReaderWriterLock();

		private void MainForm_Load(object sender, EventArgs e)
		{
			this.folderBrowserDialog1.Description = "検索先のフォルダを指定してください。";
			this.folderBrowserDialog2.Description = "検索結果を出力するフォルダを指定してください。";
			this.label1.Text = "検索先と検索結果を出力するフォルダを指定して開始ボタンをクリックしてください。";
			this.button2.Visible = false;
		}

		private void button1_Click(object sender, EventArgs e)
		{
			if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
			{
				if (this.folderBrowserDialog2.ShowDialog() == DialogResult.OK)
				{
					this.FindMatchLinesFromWord(this.folderBrowserDialog1.SelectedPath, this.folderBrowserDialog2.SelectedPath);
				}
			}
		}

		private void FindMatchLinesFromWord(string srcPath, string targetPath) {
			this.button1.Visible = false;
			string textDir = Path.Combine(srcPath, ".adsidx");
			if (Directory.Exists(textDir)) { }
			List<string> srcFiles = this.FindDocuments(textDir);
			foreach (string srcFile in srcFiles)
			{
				string fileName = Path.GetFileName(srcFile);
				// 拡張子を除いたファイル名を取得
				string docFile = Path.Combine(srcPath, Path.GetFileNameWithoutExtension(fileName));
				this.label1.Text = docFile + "を処理中。";
				Application.DoEvents();
				Debug.WriteLine(docFile);
				this.EditWord(srcFile, docFile, targetPath);
			}
			this.label1.Text = "終了しました。";
			this.button2.Visible = true;

		}

		private List<string> FindDocuments(string textDir)
		{
			List<string> exts = new List<string> { ".txt" };
			List<string> srcFiles = new List<string>();
			if (!Directory.Exists(textDir)) return srcFiles;
			string[] files = Directory.GetFiles(textDir, "*", SearchOption.AllDirectories);
			foreach (string file in files)
			{
				if (exts.Contains(Path.GetExtension(file).ToLower()))
				{
					srcFiles.Add(file);
				}
			}
			return srcFiles;
		}

		private void EditWord(string srcFile, string docFile, string targetDir)
		{
			string targetPath = Path.Combine(targetDir, Path.GetFileName(docFile));
			File.Copy(docFile, targetPath, true);
			try
			{
				string text = string.Empty;
				List<int> indexes = new List<int>();
				List<double> rates = new List<double>();
				List<string> lines = new List<string>();
				double[] ratesArray = { 1, 0.95, 0.85 };
				using (StreamReader file = new StreamReader(srcFile))
				{
					string line;
					int index = 0;
					double rate = 0D;
					while ((line = file.ReadLine()) != null)
					{
						lines.Add(line);
						indexes.Add(index++);
						Random random = new Random(Guid.NewGuid().GetHashCode());
						int i = random.Next(0, ratesArray.Length);
						rate = ratesArray[i];
						rates.Add(rate);
					}
				}
				WordTextHighLighter whl = new WordTextHighLighter();
				string message = whl.HighlightTextInWord(targetPath, indexes.ToArray(), rates.ToArray(), lines.ToArray());
				if (!string.IsNullOrEmpty(message)) this.WriteMatchLine(message, docFile, targetDir);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
			}
			//File.Delete(targetPath);
		}

		private void WriteMatchLine(string message, string docFile, string targetDir)
		{
			string filename = Path.Combine(targetDir, Path.GetFileName(docFile) + ".txt");
			File.Delete(filename);
			rwl.AcquireWriterLock(Timeout.Infinite);
			// ファイルオープン
			try
			{
				using (FileStream fs = File.Open(filename, FileMode.Append))
				using (StreamWriter writer = new StreamWriter(fs, Encoding.UTF8))
				{
					writer.WriteLine(message);
				}
			}
			finally
			{
				// ロック解除は finally の中で行う
				rwl.ReleaseWriterLock();
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			this.Close();
		}
	}
}
