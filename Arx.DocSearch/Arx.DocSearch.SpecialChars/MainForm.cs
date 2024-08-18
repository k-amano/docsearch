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
			string textDir = Path.Combine(srcPath, ".adsidx");
			if (Directory.Exists(textDir)) { }
			List<string> srcFiles = this.FindDocuments(textDir);
			foreach (string srcFile in srcFiles)
			{
				string fileName = Path.GetFileName(srcFile);
				// 拡張子を除いたファイル名を取得
				string docFile = Path.Combine(srcPath, Path.GetFileNameWithoutExtension(fileName));
				this.label1.Text = docFile + "を処理中。";
				Debug.WriteLine(docFile);
				this.EditWord(srcFile, docFile, targetPath);
			}
			this.label1.Text = "終了しました。";

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
				WordTextExtractor wte = new WordTextExtractor(docFile);
				string docText = wte.Text;
				string text = string.Empty;
				List<string> lines = new List<string>();
				using (StreamReader file = new StreamReader(srcFile))
				{
					string line;
					while ((line = file.ReadLine()) != null)
					{
						lines.Add(line);
					}
				}
				bool found = true;
				for (int i = 0; i < lines.Count; i++)
				{
					string line = lines[i];
					if (!this.FindMatchLine(i, line, docFile, docText, targetDir)) found = false;
				}
				if (!found) this.WriteMatchLine(docText, docFile, targetDir);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
			}
			File.Delete(targetPath);
		}

		private bool FindMatchLine(int index, string line, string docFile, string docText, string targetDir)
		{
			bool found = true;
			line = line.Trim();
			if (0 == line.Length) return found;
			try
			{
				// 検索テキストを正規表現パターンに変換
				// Replace smart quotes with regular quotes
				string normalizedText = Regex.Replace(line, @"([.:;)])(?!\s)", "$1 "); //「.:;)」の後に空白を入れる
				normalizedText = Regex.Replace(normalizedText, @"(?<!\s)([.:;)])", " $1"); //「.:;)」の前に空白を入れる
				/*normalizedText = Regex.Replace(normalizedText, @"\uF06D", " ");//ミクロン記号μ
				normalizedText = Regex.Replace(normalizedText, @"eq\\o\([^,]+,¯\s\)", " ");//EQフィールド(数式)*/
				string pattern = CreateSearchPattern(normalizedText);
				// 正規表現を使用して検索
				Match match = Regex.Match(docText, pattern, RegexOptions.IgnoreCase);
				if (!match.Success)
				{
					string message = string.Format("Not found: index:{0} \nline:\n{1}\npattern:\n{2}\n", index + 1, normalizedText, pattern);
					this.WriteMatchLine(message, docFile, targetDir);
					found = false;
				} else {

					Debug.WriteLine(string.Format("index={0} pattern={1}", match.Index, pattern)) ;
				}
			}
			catch (Exception e)
			{
				Debug.WriteLine("FindMatchLine:" + e.Message);
			}
			return found;
		}

		private string CreateSearchPattern(string searchText)
		{
			// Split the text into words
			string[] words = searchText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

			// Process each word
			string[] processedWords = words.Select(word =>
			{
				// Escape special regex characters except [ and ]
				string escaped = Regex.Replace(word, @"[.^$*+?()[\]\\|{}]", @"\$&");

				// Handle apostrophes specially
				escaped = Regex.Replace(escaped, @"'", @"[‘’']");
				escaped = Regex.Replace(escaped, @"""", @"[“”®™–—""]");

				return escaped;
			}).ToArray();
			// Join the words with flexible whitespace
			return string.Join(@"\s*", processedWords);
		}

		private void WriteMatchLine(string message, string docFile, string targetDir)
		{
			string filename = Path.Combine(targetDir, Path.GetFileName(docFile) + ".txt");
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
	}
}
