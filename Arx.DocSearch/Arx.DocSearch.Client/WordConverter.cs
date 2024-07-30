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
using Arx.DocSearch.Util;
//using NPOI.SS.Formula.Functions;
using System.Threading;
using NPOI.SS.Formula;

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
			Document doc = null;
			object miss = System.Reflection.Missing.Value;
			object path = docFile;
			object path2 = Path.Combine(this.seletedPath, Path.GetFileName(docFile));
			object readOnly = false;
			try
			{
				doc = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
				// 特殊空白文字を置換
				this.ReplaceSpecialChar(doc, "^s", " ", true);
				this.ReplaceSpecialChar(doc, "^-", "", true);
				doc.Fields.ToggleShowCodes();
				doc.Fields.Unlink();
				string text = string.Empty;
				foreach (KeyValuePair<int, MatchLine> ml in this.matchLines)
				{
					MatchLine m = ml.Value;
					int index = isTarget ? m.TargetLine: ml.Key;
					double rate =  m.Rate;
					this.FindMatchLine(index, rate, isTarget, doc, docFile);
					//if (index > 500) break;
				}
				doc.SaveAs(ref path2, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

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

		private void FindMatchLine(int index, double rate, bool isTarget, Document doc, string docFile)
		{
			Range range = doc.Range();
			string line = isTarget ? this.lsTarget[index].Trim() : this.lsSrc[index].Trim();
			if (0 == line.Length) return;
			if (index < 1000) this.mainForm.WriteLog(string.Format("FindMatchLine:{0}:{1}:{2:0.00}:{3}", isTarget, index, rate, line));
			this.ChangeColorOfDocument(line, rate, index, range, doc, docFile);
		}

		private void ChangeColorOfDocument(string line, double rate, int index, Range range, Document doc, string docFile)
		{
			string searchPattern = "";
			try
			{
				line = line.Trim();
				/*if (180 < line.Length)
				{
					// 先頭100文字と末尾100文字を使用し、特殊文字をエスケープ
					string start = EscapeSpecialChars(line.Substring(0, 90));
					string end = EscapeSpecialChars(line.Substring(line.Length - 90));
					searchPattern = start + "*" + end;
				}
				else
				{
					searchPattern = EscapeSpecialChars(line);
				}
				Regex re = new Regex(@"[""']");
				searchPattern = re.Replace(searchPattern, "##f3qgSJhXgamY##", 7, 0); //次の行で変換されないように一旦別文字列とする
				searchPattern = Regex.Replace(searchPattern, @"\s*[""']\s*", "*");
				searchPattern = Regex.Replace(searchPattern, @"##f3qgSJhXgamY##", @"[“”’""'®]");//本来置き換えたい文字列に変換
				searchPattern = Regex.Replace(searchPattern, @"\s*\*\s*", "*");
				searchPattern = Regex.Replace(searchPattern, @"(:|;|and) +", "$1*");
				searchPattern = Regex.Replace(searchPattern, @"\. *(?!$)", "[. ]@");
				searchPattern = Regex.Replace(searchPattern, @" +(?!\])", " @");
				searchPattern = Regex.Replace(searchPattern, @"\.$", @"\.");
				searchPattern = searchPattern.Trim('*');
				if (index < 1000) this.mainForm.WriteLog(string.Format("searchPattern={0}", searchPattern));
				if (range.Find.Execute(searchPattern, MatchWildcards: true))
				{
					DoChangeColor(range, rate, index);
				} else {
					this.ResearchMatchLine(line, rate, index, doc, docFile);
					//this.WriteMatchLine(line, searchPattern, rate, index, docFile);
				}*/
				this.ResearchMatchLine(line, rate, index, doc, docFile);
			} catch(Exception e) {
				this.mainForm.WriteLog("searchPattern:" + searchPattern + "\n" +e.Message);
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
			if (index < 1000) this.mainForm.WriteLog(string.Format("DoChangeColor:{0:0.00}:{1}:{2}", rate, color, range.Text));
		}

		private string EscapeSpecialChars(string input)
		{
			string[] specialChars = new[] { "\\", "(", ")", "[", "]", "{", "}", "^", "$", "|", "?", "*", "+", "<", ">" };
			foreach (var specialChar in specialChars)
			{
				input = input.Replace(specialChar, "\\" + specialChar);
			}
			return input;
		}

		private string EscapeRegexSpecialChars(string input)
		{
			string[] specialChars = { "\\", ".", "+", "*", "?", "[", "^", "]", "$", "(", ")", "{", "}", "=", "!", "<", ">", "|", ":", "-" };
			string escapedInput = input;

			foreach (string specialChar in specialChars)
			{
				escapedInput = escapedInput.Replace(specialChar, "\\" + specialChar);
			}

			return escapedInput;
		}
		private void EditWord_bak(Application word, string docFile, bool isTarget)
		{
			Document doc = null;
			object miss = System.Reflection.Missing.Value;
			object path = docFile;
			object path2 = Path.Combine(this.seletedPath, Path.GetFileName(docFile));
			object readOnly = false;
			int index = -1;
			int nextIndex = -1;
			int paragraphPos = -1;
			int nextPos = 0;
			int charPos = -1;
			this.mainForm.WriteLog(string.Format("WordConverter.EditWord: path={0} path2={1}", path, path2));
			try
			{
				//throw new Exception("An exception occurs.");
				doc = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
				string text = string.Empty;
				//this.mainForm.WriteLog(string.Format("doc.Paragraphs.Count={0}", doc.Paragraphs.Count));
				//this.mainForm.WriteLog(string.Format("doc.Tables.Count={0}", doc.Tables.Count));
				for (int i = 0; i < doc.Tables.Count; i++)
				{
					Table table = doc.Tables[i + 1];
					table.Delete();
				}
				//this.mainForm.WriteLog(string.Format("doc.Paragraphs.Count={0}", doc.Paragraphs.Count));
				for (int i = 0; i < doc.Paragraphs.Count; i++)
				{
					if (i < 10) this.mainForm.WriteLog(string.Format("Paragraphs[]={0}", i + 1));
					Paragraph para = doc.Paragraphs[i + 1];
					Range r = para.Range;
					try
					{
						if (0 == (r?.Text?.Trim().Length ?? 0)) continue;
					}
					catch (Exception e)
					{
						this.mainForm.WriteLog(string.Format("path={0} i={1}:{2}\n{3}", path, i, e.Message, e.StackTrace));
						throw new Exception("An exception occurs.");
					}
					bool isDebug = false;
					/*if (i < 10)
					{
						string str = r.Text;
						if (80 < str.Length) str = str.Substring(0, 80);
						this.mainForm.WriteLog(string.Format("#### EditWord i={0} paragraphPos={1} r.Text.Length={2}:{3}", i, paragraphPos, r.Text.Trim().Length, str));
						isDebug = true;
					}*/
					if (i < 10) isDebug = true;
					else isDebug = false;
					//if (isDebug) this.mainForm.WriteLog(string.Format("i={0} r.Text={1}#", i, r.Text));
					this.FindParagraph(r.Text, ref paragraphPos, ref nextPos, isTarget, isDebug);
					//if (65 < i && i < 70) Debug.WriteLine(string.Format("paragraphPos={0} nextPos={1}#", paragraphPos, nextPos));
					List<string> sentences = new List<string>();
					StringBuilder sb = new StringBuilder();
					Console.WriteLine(string.Format("Sentences.Count={0}", r.Sentences.Count));
					for (int j = 0; j < r.Sentences.Count; j++)
					{
						string str = r.Sentences[j + 1].Text;
						//if (isDebug) Debug.WriteLine(string.Format("count={0} text={1}#", j, str));
						Console.WriteLine(string.Format("Sentences[{0}].Text:{1}", j + 1, str));
						if (0 < j)
						{
							if (this.StartsWithCapital(str))
							{
								sentences.Add(sb.ToString());
								sb = new StringBuilder(); ;
							}
							else
							{
								sentences.Add(string.Empty);
							}
						}
						sb.Append(str);
					}
					sentences.Add(sb.ToString());
					Color[] colors = new Color[sentences.Count];
					Color color;
					for (int j = 0; j < sentences.Count; j++)
					{
						color = Color.White;
						if (!string.IsNullOrEmpty(sentences[j]))
						{
							string[] lines = this.ReplaceLine(sentences[j]);
							text = lines[lines.Length - 1].Trim();
							if (isTarget) color = this.GetTargetColor(text, paragraphPos, ref index, ref charPos, ref nextIndex);
							else color = this.GetSrcColor(text, paragraphPos, ref index, ref charPos, ref nextIndex);
							//if (isDebug) Debug.WriteLine(string.Format("color={0} text={1}#", color, text));
						}
						colors[j] = color;
					}
					//センテンスの長さが0の場合は後ろの色とするため、終わりから色を振り直す。
					int id = 0;
					color = Color.White;
					for (int j = 0; j < sentences.Count; j++)
					{
						id = sentences.Count - 1 - j;
						if (!string.IsNullOrEmpty(sentences[id]))
						{
							color = colors[id];
						}
						colors[id] = color;
						//if (isDebug) Debug.WriteLine(string.Format("id={0} color={1}#", id, colors[id]));
					}
					for (int j = 0; j < r.Sentences.Count; j++)
					{
						Range r2 = r.Sentences[j + 1];
						r2.Font.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(colors[j]);
						r2.Select();
					}
				}
				doc.SaveAs(ref path2, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
				this.mainForm.WriteLog("WordConverter.EditWord:" + e.Message + "\n" + e.StackTrace);
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

		private Color GetSrcColor(string text, int paragraphPos, ref int index, ref int charPos, ref int nextIndex)
		{
			Color ret = Color.White;
			text = text.Trim();
			if (paragraphPos < 0 || string.IsNullOrEmpty(text)) return ret;
			List<int> paragraph = this.srcParagraphs[paragraphPos];
			this.getLineIndex(this.lsSrc, text, paragraph, ref index, ref charPos, ref nextIndex);
			if (this.lsSrc.Count <= index) return ret;
			if (this.matchLines.ContainsKey(index))
			{
				if (1D == this.matchLines[index].Rate) ret = Color.LightPink;
				else if (0.9 <= this.matchLines[index].Rate) ret = Color.Yellow;
				else if (0D < this.matchLines[index].Rate) ret = Color.LightGreen;
			}
			return ret;
		}

		private Color GetTargetColor(string text, int paragraphPos, ref int index, ref int charPos, ref int nextIndex)
		{
			Color ret = Color.White;
			text = text.Trim();
			if (paragraphPos < 0 || string.IsNullOrEmpty(text)) return ret;
			List<int> paragraph = this.targetParagraphs[paragraphPos];
			this.getLineIndex(this.lsTarget, text, paragraph, ref index, ref charPos, ref nextIndex);
			if (this.lsTarget.Count <= index) return ret;
			double rate = getTargetLineRate(index);
			if (1D == rate) ret = Color.LightPink;
			else if (0.9 <= rate) ret = Color.Yellow;
			else if (0D < rate) ret = Color.LightGreen;
			return ret;
		}

		private void getLineIndex(List<string> ls, string text, List<int> paragraph, ref int index, ref int charPos, ref int nextIndex)
		{
			string replaced = text.Trim().Trim('.');
			if (120 < replaced.Length) replaced = replaced.Substring(0, 120);
			for (int i = 0; i < paragraph.Count; i++)
			{
				if (paragraph[i] < nextIndex) continue;
				string line = ls[paragraph[i]];
				line = line.Trim().Trim('.');
				if (0 < replaced.Length && 0 < line.Length)
				{
					int start = charPos < 0 ? 0 : line.Length <= charPos ? line.Length - 1 : charPos;
					int newPos = line.IndexOf(replaced, start);
					if (0 <= newPos)
					{
						index = paragraph[i];
						charPos = newPos + text.Length;
						if (line.Length <= charPos)
						{
							nextIndex = index + 1;
							charPos = -1;
						}
						else
						{
							nextIndex = index;
						}
						return;
					}
				}
				index++;
				charPos = -1;
				nextIndex = index;
			}
		}

		private double getTargetLineRate(int targetLine)
		{
			foreach (KeyValuePair<int, MatchLine> pair in this.matchLines)
			{
				MatchLine ml = pair.Value;
				if (ml.TargetLine == targetLine) return ml.Rate;
			}
			return 0D;
		}

		private string[] ReplaceLine(string line)
		{
			line = Regex.Replace(line ?? "", @"[\x00-\x1F\x7F]", "");
			line = Regex.Replace(line ?? "", @"[\u00a0\uc2a0]", " "); //文字コードC2A0（UTF-8の半角空白）
			line = Regex.Replace(line ?? "", @"[\u0091\u0092\u2018\u2019]", "'"); //UTF-8のシングルクォーテーション
			line = Regex.Replace(line ?? "", @"[\u0093\u0094\u00AB\u201C\u201D]", "\""); //UTF-8のダブルクォーテーション
			line = Regex.Replace(line ?? "", @"[\u0097\u2013\u2014]", "\""); //UTF-8のハイフン
			line = Regex.Replace(line ?? "", @"[\u00A9\u00AE\u2022\u2122]", "\""); //UTF-8のスラッシュ
			//スペースに挟まれた「Fig.」で次が大文字でない場合は、文末と混同しないようにドットの後ろのスペースを削除する。
			line = Regex.Replace(line ?? "", @"(^fig| fig)\. +([^A-Z])", "$1.$2", RegexOptions.IgnoreCase);
			//2個以上連続するスペースは1個の半角スペースにする。
			line = Regex.Replace(line ?? "", @"\s+", " ");
			line = TextConverter.ZenToHan(line);
			line = TextConverter.HankToZen(line);
			line = Regex.Replace(line ?? "", @"^((([\(\[<（＜〔【≪《])([^0-9]*[0-9]*)([\)\]>）＞〕】≫》])(\s*))+)", "\n$1  \n"); //【数字】
			line = Regex.Replace(line ?? "", @"^([0-9]+)(\.?)", "\n$1$2  \n"); //数字
			string[] lines = line.Split('\n');
			return lines;
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

		private void FindParagraph(string para, ref int paragraphPos, ref int nextPos, bool isTarget, bool isDebug)
		{
			List<List<int>> paragraphs = isTarget ? this.targetParagraphs : this.srcParagraphs;
			string[] lines = this.ReplaceLine(para);
			foreach (string line in lines)
			{
				string replaced = line.Trim();
				if (0 == replaced.Length) continue;
				if (120 < replaced.Length) replaced = replaced.Substring(0, 120);
				for (int i = nextPos; i < paragraphs.Count; i++)
				{
					string text = this.GetParagraphText(paragraphs[i], isTarget);
					if (isDebug) this.mainForm.WriteLog(string.Format("FindParagraph i={0}:text={1} paragraphPos={2}, nextPos={3}#", i, text, paragraphPos, nextPos));
					if (0 < text.Trim().Length && 0 <= text.IndexOf(replaced))
					{
						paragraphPos = i;
						if (para.Trim().EndsWith(".")) nextPos = paragraphPos + 1;
						if (isDebug) this.mainForm.WriteLog(string.Format("found: text.Index={0} replaced={1} paragraphPos={2}, nextPos={3}#", text.IndexOf(replaced), replaced, paragraphPos, nextPos));
						return;
					}
					//if (isDebug) Debug.WriteLine(string.Format("pos={0} replaced={1}#", text.IndexOf(replaced), replaced));
				}
			}
		}

		private string GetParagraphText(List<int> paragraph, bool isTarget)
		{
			StringBuilder sb = new StringBuilder();
			List<string> ls = isTarget ? this.lsTarget : this.lsSrc;
			foreach (int i in paragraph)
			{
				string line = ls[i];
				sb.Append(line);
			}
			//2個以上連続するスペースは1個の半角スペースにする。
			return Regex.Replace(sb.ToString(), @"\s+", " ");
		}

		private void DumpParagraphs(List<List<int>> paragraphs, bool isTarget)
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

		}

		private bool StartsWithCapital(string line)
		{
			if (Regex.IsMatch(line ?? "", @"^(\s*[A-Z])")) return true;
			else return false;
		}

		private void ResearchMatchLine(string line, double rate, int index, Document doc, string docFile)
		{
			// 文書の全テキストを取得（改行を含む）
			string docText = doc.Content.Text;
			// 検索テキストを正規表現パターンに変換
			string pattern = CreateSearchPattern(line);
			// 正規表現を使用して検索
			Match match = Regex.Match(docText, pattern, RegexOptions.IgnoreCase);
			if (match.Success)
			{
				// 見つかったテキストに黄色の背景色をつける
				Range range = doc.Range(match.Index, match.Index + match.Length);
				DoChangeColor(range, rate, index);
			}else{
				WriteMatchLine(line, pattern, rate, index, docFile);
			}
		}

		/*private string CreateSearchPattern(string searchText)
		{
			// 引用符を正規表現パターンに変換
			string quotedText = Regex.Replace(searchText, @"[“”""]", @"[“”®–""]");
			quotedText = Regex.Replace(quotedText, @"[’']", @"[’']");

			// 単語ごとに分割し、エスケープして結合
			string[] words = quotedText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
			return string.Join(@"\s*", words.Select(word => Regex.Escape(word)));
		}*/

		private string CreateSearchPattern(string searchText)
		{
			// Replace smart quotes with regular quotes
			string normalizedText = Regex.Replace(searchText, @"[’']", "'");

			// Split the text into words
			string[] words = normalizedText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

			// Process each word
			string[] processedWords = words.Select(word =>
			{
				// Escape special regex characters except [ and ]
				string escaped = Regex.Replace(word, @"[.^$*+?()[\]\\|{}]", @"\$&");

				// Handle apostrophes specially
				escaped = Regex.Replace(escaped, @"'", @"[’']");
				escaped = Regex.Replace(escaped, @"""", @"[“”®–""]");

				return escaped;
			}).ToArray();

			// Join the words with flexible whitespace
			return string.Join(@"\s*", processedWords);
		}

		private void WriteMatchLine(string line, string pattern, double rate, int index, string docFile)
		{
			string filename = Path.Combine(this.seletedPath, Path.GetFileName(docFile) + ".txt");
			string message = string.Format("index:{0} rate:{1:0.00}\nline:\n{2}\npattern:\n{3}\n", index, rate, line, pattern);
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
			#endregion
		}
}

