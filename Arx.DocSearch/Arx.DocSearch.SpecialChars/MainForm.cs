using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Threading;
using System.Text;


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
			try
			{
				this.word = new Application();
				this.word.Visible = false;
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.StackTrace);
			}
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (null != this.word)
			{
				((_Application)this.word).Quit();
				Marshal.ReleaseComObject(this.word);  // オブジェクト参照を解放
				this.word = null;
			}

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
			// OpenXML を使用して文書を処理
			this.ReplaceSymbolChars(targetPath);
			Document doc = null;
			object miss = System.Reflection.Missing.Value;
			object path = targetPath;
			object readOnly = false;
			try
			{
				doc = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
				// 特殊空白文字を置換
				this.ReplaceSpecialChar(doc, "^s", " ", true);
				this.ReplaceSpecialChar(doc, "^-", "", true);
				doc.Fields.ToggleShowCodes();
				doc.Fields.Unlink();
				string docText = doc.Content.Text;
				docText = Regex.Replace(docText, @"\uF06D", " ");//ミクロン記号μ
				docText = Regex.Replace(docText, @"[\x00-\x1F]", " ");//非表示キャラクタ
				docText = this.ConvertMathematicalString(docText);
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
			finally
			{
				if (null != doc)
				{
					((_Document)doc).Close();
					Marshal.ReleaseComObject(doc);  // オブジェクト参照を解放
					doc = null;
					File.Delete(targetPath);
				}
			}
		}

		private void ReplaceSpecialChar(Document doc, string text, string replacement, bool matchWildcards)
		{
			// 特殊空白文字を置換
			Find find = doc.Content.Find;
			find.ClearFormatting();
			find.Replacement.ClearFormatting();
			find.Text = text;
			find.Replacement.Text = replacement;
			find.Forward = true;
			find.Wrap = WdFindWrap.wdFindContinue;
			find.Format = false;
			find.MatchCase = false;
			find.MatchWholeWord = false;
			find.MatchPhrase = false;
			find.MatchSoundsLike = false;
			find.MatchAllWordForms = false;
			find.MatchFuzzy = false;
			find.MatchWildcards = matchWildcards;  // ワイルドカード検索を有効化
			find.Execute(Replace: WdReplace.wdReplaceAll);
		}

		private void ReplaceSymbolChars(string filePath)
		{
			using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
			{
				var body = wordDoc.MainDocumentPart.Document.Body;
				if (body != null)
				{
					this.ReplaceSymbolCharsByElement(body);
				}
			}
		}

		private void ReplaceSymbolCharsByElement(OpenXmlElement element)
		{
			// すべての子要素に対して再帰的に処理を行う
			foreach (var childElement in element.Elements().ToList())
			{
				if (childElement is Run run)
				{
					var symbolChars = run.Elements<SymbolChar>().Where(s => s.Font == "Symbol" && s.Char == "F06D").ToList();
					foreach (var symbolChar in symbolChars)
					{
						var newText = new Text(" ") { Space = SpaceProcessingModeValues.Preserve };
						run.ReplaceChild(newText, symbolChar);
					}
				}
				else
				{
					ReplaceSymbolCharsByElement(childElement);
				}
			}
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
				normalizedText = Regex.Replace(normalizedText, @"\uF06D", " ");//ミクロン記号μ
				normalizedText = Regex.Replace(normalizedText, @"eq\\o\([^,]+,¯\s\)", " ");//EQフィールド(数式)
				normalizedText = Regex.Replace(normalizedText, @"apparatusincludes", @"apparatus includes");//F1-16A21C9U.docxでタブが抜ける対応
				string pattern = CreateSearchPattern(normalizedText);
				// 正規表現を使用して検索
				Match match = Regex.Match(docText, pattern, RegexOptions.IgnoreCase);
				if (!match.Success)
				{
					string message = string.Format("Not found: index:{0} \nline:\n{1}\npattern:\n{2}\n", index + 1, normalizedText, pattern);
					this.WriteMatchLine(message, docFile, targetDir);
					found = false;
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
		private string ConvertMathematicalString(string input)
		{
			StringBuilder result = new StringBuilder();

			for (int i = 0; i < input.Length; i++)
			{
				int codePoint = char.ConvertToUtf32(input, i);
				if (char.IsSurrogatePair(input, i))
				{
					i++; // サロゲートペアなので2文字分進める
				}

				string convertedChar = ConvertCodePoint(codePoint);
				result.Append(convertedChar);
			}

			return result.ToString();
		}

		private string ConvertCodePoint(int codePoint)
		{
			// 数学的イタリック小文字（𝑎 to 𝑧）を通常の小文字（a to z）に変換
			if (codePoint >= 0x1D44E && codePoint <= 0x1D467)
			{
				return ((char)('a' + (codePoint - 0x1D44E))).ToString();
			}
			// 数学的イタリック大文字（𝐴 to 𝑍）を通常の大文字（A to Z）に変換
			else if (codePoint >= 0x1D434 && codePoint <= 0x1D44D)
			{
				return ((char)('A' + (codePoint - 0x1D434))).ToString();
			}
			// 数学的ボールド小文字ギリシャ文字（𝛂 to 𝛚）を通常の小文字ギリシャ文字（α to ω）に変換
			else if (codePoint >= 0x1D6C2 && codePoint <= 0x1D6DA)
			{
				return ((char)('α' + (codePoint - 0x1D6C2))).ToString();
			}
			// ギリシャ文字とその他の特殊文字はそのまま
			else if ((codePoint >= 0x0391 && codePoint <= 0x03C9) || // ギリシャ文字範囲
					 codePoint == 0x0394 || codePoint == 0x03B4 || codePoint == 0x03BB) // Δ, δ, λ
			{
				return char.ConvertFromUtf32(codePoint);
			}
			// その他の文字はそのまま
			else
			{
				return char.ConvertFromUtf32(codePoint);
			}
		}

	}
}
