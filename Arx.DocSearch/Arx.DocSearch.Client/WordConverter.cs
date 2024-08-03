using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;
using Color = System.Drawing.Color;
using Arx.DocSearch.Util;
//using NPOI.SS.Formula.Functions;
using System.Threading;
using NPOI.SS.Formula;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Arx.DocSearch.Client
{
	public class WordConverter
	{
				#region コンストラクタ
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public WordConverter(string srcFile, List<string> lsSrc, string targetFile, List<string> lsTarget, Dictionary<int, MatchLine> matchLines, string seletedPath, MainForm mainForm)
		{
			this.srcFile = srcFile;
			this.targetFile = targetFile;
			this.lsSrc = lsSrc;
			this.lsTarget = lsTarget;
			this.matchLines = matchLines;
			this.seletedPath = seletedPath;
			this.mainForm = mainForm;
		}
		#endregion
		#region フィールド
		private string srcFile;
		private string targetFile;
		List<string> lsSrc;
		List<string> lsTarget;
		private Dictionary<int, MatchLine> matchLines;
		string seletedPath;
		List<List<int>> srcParagraphs;
		List<List<int>> targetParagraphs;
        private MainForm mainForm;
		// ロック用のインスタンス
		private static ReaderWriterLock rwl = new ReaderWriterLock();
		#endregion

		#region メソッド
		public void Run()
		{
			this.GetSrcParagraphs();
			this.GetTargetParagraphs();
			//this.DumpParagraphs(targetParagraphs, true);
			//this.DumpMatchLines();
			Process[] wordProcesses = Process.GetProcessesByName("WINWORD");
			if (0 < wordProcesses.Length)
			{
				//メッセージボックスを表示する
				DialogResult result = MessageBox.Show("起動中の Word を終了してよろしいですか？",
						"確認",
						MessageBoxButtons.YesNoCancel,
						MessageBoxIcon.Exclamation,
						MessageBoxDefaultButton.Button2);

				//何が選択されたか調べる
				if (result == DialogResult.Yes)
				{
					//「はい」が選択された時
					foreach (Process p in wordProcesses)
					{
						try
						{
							p.Kill();
							p.WaitForExit(); // possibly with a timeout
						}
						catch (Exception e)
						{
							Debug.WriteLine(e.StackTrace);
						}
					}
				} else return;
			}
			Application word = null;
			try
			{
				word = new Application();
				word.Visible = false;
				this.EditWord(word, srcFile, false);
				this.EditWord(word, targetFile, true);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
			}
			finally
			{
				if (null != word)
				{
					((_Application)word).Quit();
					Marshal.ReleaseComObject(word);  // オブジェクト参照を解放
					word = null;
				}
			}

		}

		private void EditWord(Application word, string docFile, bool isTarget)
		{
			string targetPath = Path.Combine(this.seletedPath, Path.GetFileName(docFile));
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
				string text = string.Empty;
				int pos = 0;
				foreach (KeyValuePair<int, MatchLine> ml in this.matchLines)
				{
					MatchLine m = ml.Value;
					int index = isTarget ? m.TargetLine: ml.Key;
					double rate =  m.Rate;
					this.FindMatchLine(index, rate, isTarget, doc, docFile, docText, ref pos);
					//if (index > 500) break;
				}
				doc.SaveAs(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
                this.mainForm.WriteLog("WordConverter.EditWord:" + e.Message + "\n"+  e.StackTrace);
            }
			finally
			{
				if (null != doc)
				{
					((_Document)doc).Close();
					Marshal.ReleaseComObject(doc);  // オブジェクト参照を解放
					doc = null;
				}
			}
		}

		private void FindMatchLine(int index, double rate, bool isTarget, Document doc, string docFile, string docText, ref int pos)
		{
			string line = isTarget ? this.lsTarget[index].Trim() : this.lsSrc[index].Trim();
			if (0 == line.Length) return;
			try
			{
				// 検索テキストを正規表現パターンに変換
				// Replace smart quotes with regular quotes
				string normalizedText = Regex.Replace(line, @"(\.|:|;)(?!\s)", "$1 "); //「.:;」の後に空白を入れる
				normalizedText = Regex.Replace(normalizedText, @"\uF06D", " ");//ミクロン記号μ
				string pattern = CreateSearchPattern(normalizedText);
				// 正規表現を使用して検索
				Match match = Regex.Match(docText, pattern, RegexOptions.IgnoreCase);
				if (match.Success)
				{
					// 見つかったテキストに黄色の背景色をつける
					int start = match.Index - pos;
					Range range = doc.Range(start, start + match.Length);
					int offset = 0;
					for (int i = 0; i < 10; i++)
					{
						offset = this.CompareResult(normalizedText, range);
						if (0 == offset) break;
						start -= offset;
						range = doc.Range(start, start + match.Length);
						//string message = string.Format("Wrong position2: i:{0} index:{1} offset:{2} pos:{3}\nline:\n{4}\nrange.Text:\n{5}\n", i, match.Index, offset, pos, normalizedText, range.Text);
						//this.WriteMatchLine(message, docFile);
						pos += offset;
					}
					if (0 != offset)
					{
						string message = string.Format("Wrong position: index:{0} length:{1}\nline:\n{2}\nrange.Text:\n{3}\n", match.Index, match.Length, normalizedText, range.Text);
						this.WriteMatchLine(message, docFile);
					}
					this.DoChangeColor(range, rate, index);
				}
				else
				{
					string message = string.Format("Not found: index:{0} rate:{1:0.00}\nline:\n{2}\npattern:\n{3}\n", index, rate, line, pattern);
					this.WriteMatchLine(message, docFile);
				}
			}
			catch (Exception e)
			{
				this.mainForm.WriteLog("FindMatchLine:" + e.Message);
			}
		}

		private void DoChangeColor(Range range, double rate,int index)
		{
			Color color = Color.White;
			if (1D == rate) color = Color.LightPink;
			else if (0.9 <= rate) color = Color.LightCyan;
			else if (0D < rate) color = Color.LightGreen;
			range.Font.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(color);
			range.Select();
		}

		private void GetSrcParagraphs() {
			this.srcParagraphs = new List<List<int>>();
			List<int> paragraph = new List<int>();
			for (int i = 0; i < this.lsSrc.Count;i++ )
			{
				string line = this.lsSrc[i];
				paragraph.Add(i);
				if (!line.EndsWith("  ")) {
					this.srcParagraphs.Add(paragraph);
					paragraph = new List<int>();
				}
			}
			if (0 < paragraph.Count) this.srcParagraphs.Add(paragraph);
		}

		private void GetTargetParagraphs()
		{
			this.targetParagraphs = new List<List<int>>();
			List<int> paragraph = new List<int>();
			for (int i = 0; i < this.lsTarget.Count; i++)
			{
				string line = this.lsTarget[i];
				paragraph.Add(i);
				if (!line.EndsWith("  "))
				{
					this.targetParagraphs.Add(paragraph);
					paragraph = new List<int>();
				}
			}
			if (0 < paragraph.Count) this.targetParagraphs.Add(paragraph);
		}

		/*private void DumpParagraphs(List<List<int>> paragraphs, bool isTarget)
		{
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < paragraphs.Count; i++)
			{
				string text = this.GetParagraphText(paragraphs[i], isTarget);
				sb.Append(string.Format("{0}: {1}\n", i, text));
			}
			this.mainForm.WriteLog(sb.ToString());
		}

		private void DumpMatchLines()
		{
			StringBuilder sb = new StringBuilder();
			foreach (KeyValuePair<int, MatchLine> ml in this.matchLines)
			{

				MatchLine m = ml.Value;
				sb.Append(string.Format("index={0} TargetLine={1} Rate={2}, MatchWords={3}, TotalWords{4}\n", ml.Key, m.TargetLine, m.Rate, m.MatchWords, m.TotalWords));
			}
			this.mainForm.WriteLog(sb.ToString());

		}*/

		/*private void SearchMatchLine(string line, double rate, int index, Document doc, string docFile, string docText, ref int pos)
		{
			// 検索テキストを正規表現パターンに変換
			// Replace smart quotes with regular quotes
			string normalizedText = Regex.Replace(line, @"(\.|:|;)(?!\s)", "$1 "); //「.:;」の後に空白を入れる
			normalizedText = Regex.Replace(normalizedText, @"\uF06D", " ");//ミクロン記号μ
			string pattern = CreateSearchPattern(normalizedText);
			// 正規表現を使用して検索
			Match match = Regex.Match(docText, pattern, RegexOptions.IgnoreCase);
			if (match.Success)
			{
				// 見つかったテキストに黄色の背景色をつける
				int start = match.Index - pos;
				Range range = doc.Range(start, start + match.Length);
				int offset = 0;
				for (int i = 0; i < 10; i++)
				{
					offset = this.CompareResult(normalizedText, range);
					if (0 == offset) break;
					start -= offset;
					range = doc.Range(start, start + match.Length);
					//string message = string.Format("Wrong position2: i:{0} index:{1} offset:{2} pos:{3}\nline:\n{4}\nrange.Text:\n{5}\n", i, match.Index, offset, pos, normalizedText, range.Text);
					//this.WriteMatchLine(message, docFile);
					pos += offset;
				}
				if (0 != offset) {
					string message = string.Format("Wrong position: index:{0} length:{1}\nline:\n{2}\nrange.Text:\n{3}\n", match.Index, match.Length, normalizedText, range.Text);
					this.WriteMatchLine(message, docFile);
				}
				this.DoChangeColor(range, rate, index);
			}
			else{
				string message = string.Format("Not found: index:{0} rate:{1:0.00}\nline:\n{2}\npattern:\n{3}\n", index, rate, line, pattern);
				this.WriteMatchLine(message, docFile);
			}
		}*/

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
				escaped = Regex.Replace(escaped, @"'", @"[’']");
				escaped = Regex.Replace(escaped, @"""", @"[“”®™–—""]");

				return escaped;
			}).ToArray();

			// Join the words with flexible whitespace
			return string.Join(@"\s*", processedWords);
		}

		private void WriteMatchLine(string message, string docFile)
		{
			string filename = Path.Combine(this.seletedPath, Path.GetFileName(docFile) + ".txt");
			rwl.AcquireWriterLock(Timeout.Infinite);
			// ファイルオープン
			try
			{
				using (FileStream fs = File.Open(filename, FileMode.Append))
				using (StreamWriter writer = new StreamWriter(fs))
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

		private int CompareResult(string line, Range range)
		{
			string text = range.Text;
			string head1 = line.Length < 10 ? line : line.Substring(0, 10);
			string head2 = text.Length < 10 ? text : text.Substring(0, 10);
			if (head1.Equals(head2)) return 0;
			int pos = line.IndexOf(head2);
			if (pos == -1) return 1;
			else return pos;
		}
	}
	#endregion
}

