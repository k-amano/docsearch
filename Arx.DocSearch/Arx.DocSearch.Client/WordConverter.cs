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
//using Microsoft.Office.Interop.Word;
//using Application = Microsoft.Office.Interop.Word.Application;
//using Document = Microsoft.Office.Interop.Word.Document;
using Color = System.Drawing.Color;
using Arx.DocSearch.Util;
using System.Threading;
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
			this.EditWord(srcFile, false);
			this.EditWord(targetFile, true);

		}

		private void EditWord(string docFile, bool isTarget)
		{
			string targetPath = Path.Combine(this.seletedPath, Path.GetFileName(docFile));
			File.Copy(docFile, targetPath, true);
			string docText = WordTextExtractor.ExtractText(targetPath);
			//this.WriteMatchLine(string.Format("EditWord: targetPath:{0} Length:{1} docText:{2}", targetPath, docText.Length, docText), docFile);
			try
			{
				using (WordprocessingDocument doc = WordprocessingDocument.Open(targetPath, true))
				{
					Body body = doc.MainDocumentPart.Document.Body;
					if (body != null)
					{
						var paragraphs = body.Descendants<Paragraph>().ToList();
						CreateDocTextWithSpaces(paragraphs, out List<int> paragraphLengths);
						foreach (KeyValuePair<int, MatchLine> ml in this.matchLines)
						{
							MatchLine m = ml.Value;
							int index = isTarget ? m.TargetLine : ml.Key;
							double rate = m.Rate;
							this.FindMatchLine(index, rate, isTarget, docFile, docText, paragraphs, paragraphLengths);

						}
						doc.MainDocumentPart.Document.Save();
					}

				}
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
				this.mainForm.WriteLog("WordConverter.EditWord:" + e.Message + "\n" + e.StackTrace);
			}
		}

		private void FindMatchLine(int index, double rate, bool isTarget, string docFile, string docText, List<Paragraph> paragraphs, List<int> paragraphLengths)
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
				//string flexiblePattern = CreateFlexibleSearchPattern(pattern);
				// 正規表現を使用して検索
				Match match = Regex.Match(docText, pattern, RegexOptions.IgnoreCase);
				string message = "";
				if (match.Success)
				{
					string highlightedText = HighlightMatch(rate, paragraphs, match, docText, paragraphLengths);
					if (highlightedText != match.Value)
					{
						message = string.Format("警告: 色付け箇所と検索テキストが異なります。: index:{0} rate:{1:0.00}\n検索テキスト: {2}\n色付け箇所: {3}", index, rate, match.Value, highlightedText);
						this.WriteMatchLine(message, docFile);
					}
				}
				else
				{
					message = string.Format("指定されたテキストが見つかりませんでした。: index:{0} rate:{1:0.00}\n検索テキスト:\n{2}\nパターン:\n{3}", index, rate, line, pattern);
					this.WriteMatchLine(message, docFile);
				}
			}
			catch (Exception e)
			{
				this.mainForm.WriteLog("FindMatchLine:" + e.Message + "\n" + e.StackTrace); 
			}
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

		private string CreateDocTextWithSpaces(List<Paragraph> paragraphs, out List<int> paragraphLengths)
		{
			StringBuilder sb = new StringBuilder();
			paragraphLengths = new List<int>();
			foreach (var paragraph in paragraphs)
			{
				string paragraphText = paragraph.InnerText;
				if (!string.IsNullOrWhiteSpace(paragraphText))
				{
					sb.Append(paragraphText);
					if (sb.Length > 0 && sb[sb.Length - 1] != ' ')
					{
						sb.Append(" ");
					}
				}
				paragraphLengths.Add(paragraphText.Length);
			}
			return sb.ToString().TrimEnd();
		}

		private string CreateFlexibleSearchPattern(string searchPattern)
		{
			return string.Join(@"\s+", searchPattern.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
				.Select(Regex.Escape));
		}

		private string HighlightMatch(double rate, List<Paragraph> paragraphs, Match match, string docText, List<int> paragraphLengths)
		{
			StringBuilder highlightedText = new StringBuilder();
			int matchStart = match.Index;
			int matchEnd = match.Index + match.Length;
			int currentIndex = 0;
			int paragraphIndex = 0;

			foreach (var paragraph in paragraphs)
			{
				string paragraphText = paragraph.InnerText;
				int paragraphLength = paragraphLengths[paragraphIndex];
				int paragraphStart = currentIndex;
				int paragraphEnd = paragraphStart + paragraphLength;

				if (paragraphStart <= matchEnd && paragraphEnd > matchStart)
				{
					var runs = paragraph.Descendants<Run>().ToList();

					int startInParagraph = Math.Max(0, matchStart - paragraphStart);
					int endInParagraph = Math.Min(paragraphLength, matchEnd - paragraphStart);

					string paragraphPrefix = Regex.Match(paragraphText, @"^\[\d+\]").Value;
					if (!string.IsNullOrEmpty(paragraphPrefix))
					{
						startInParagraph = Math.Max(paragraphPrefix.Length, startInParagraph);
					}

					highlightedText.Append(ApplyBackgroundColorToRuns(rate, runs, startInParagraph, endInParagraph));
				}

				currentIndex += paragraphLength;
				if (!string.IsNullOrWhiteSpace(paragraphText))
				{
					currentIndex += 1;
				}
				paragraphIndex++;

				if (currentIndex > matchEnd) break;
			}

			return highlightedText.ToString();
		}

		private string ApplyBackgroundColorToRuns(double rate, List<Run> runs, int startIndex, int endIndex)
		{
			StringBuilder highlightedText = new StringBuilder();
			int currentIndex = 0;
			foreach (var run in runs)
			{
				int runLength = run.InnerText.Length;
				if (currentIndex + runLength > startIndex && currentIndex < endIndex)
				{
					int runStartIndex = Math.Max(0, startIndex - currentIndex);
					int runEndIndex = Math.Min(runLength, endIndex - currentIndex);

					if (runStartIndex > 0 || runEndIndex < runLength)
					{
						SplitAndApplyBackgroundColor(rate, run, runStartIndex, runEndIndex);
					}
					else
					{
						ApplyBackgroundColor(rate, run);
					}

					highlightedText.Append(run.InnerText.Substring(runStartIndex, runEndIndex - runStartIndex));
				}
				currentIndex += runLength;
				if (currentIndex >= endIndex) break;
			}
			return highlightedText.ToString();
		}

		private void SplitAndApplyBackgroundColor(double rate, Run run, int startIndex, int endIndex)
		{
			string text = run.InnerText;
			RunProperties originalProperties = run.RunProperties?.Clone() as RunProperties;

			run.RemoveAllChildren();

			if (startIndex > 0)
			{
				run.AppendChild(new Text(text.Substring(0, startIndex)));
			}

			Run coloredRun = new Run(new Text(text.Substring(startIndex, endIndex - startIndex)));
			if (originalProperties != null)
			{
				coloredRun.RunProperties = originalProperties.Clone() as RunProperties;
			}
			ApplyBackgroundColor(rate, coloredRun);
			run.InsertAfter(coloredRun, run.LastChild);

			if (endIndex < text.Length)
			{
				Run remainingRun = new Run(new Text(text.Substring(endIndex)));
				if (originalProperties != null)
				{
					remainingRun.RunProperties = originalProperties.Clone() as RunProperties;
				}
				run.InsertAfter(remainingRun, coloredRun);
			}
		}

		private void ApplyBackgroundColor(double rate, Run run)
		{
			Color color = Color.White;
			if (1D == rate) color = Color.LightPink;
			else if (0.9 <= rate) color = Color.LightCyan;
			else if (0D < rate) color = Color.LightGreen;	
			RunProperties runProperties = run.RunProperties ?? new RunProperties();
			runProperties.Shading = new Shading() { Fill = $"{color.R:X2}{color.G:X2}{color.B:X2}", Color = "auto", Val = ShadingPatternValues.Clear };
			run.RunProperties = runProperties;
		}
	}
	#endregion
}

