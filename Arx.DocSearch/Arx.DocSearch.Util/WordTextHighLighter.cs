using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;

namespace Arx.DocSearch.Util
{
	public class WordTextHighLighter
	{
		public string HighlightTextInWord(string filePath, int[] indexes, double[] rates, string[] searchPatterns, bool isDebug = false)
		{
			string docText = WordTextExtractor.ExtractText(filePath);
			StringBuilder sb = new StringBuilder();
			if (isDebug)
			{
				sb.AppendLine($"filePath: {filePath} Length: {docText.Length} docText: {docText}");
				string indexesStr = string.Join(", ", indexes);
				string searchPatternsStr = string.Join(", ", searchPatterns);
				string ratesStr = string.Join(", ", rates);
				sb.AppendLine($"indexes: [{indexesStr}]\nsearchPatterns: [{searchPatternsStr}]\nrates: [{ratesStr}]");
			}
			try
			{
				using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
				{
					Body body = doc.MainDocumentPart.Document.Body;
					if (body != null)
					{
						var paragraphs = body.Descendants<Paragraph>().ToList();
						CreateDocTextWithSpaces(paragraphs, out List<int> paragraphLengths);
						for (int i = 0; i < searchPatterns.Length && i < rates.Length; i++)
						{
							string pattern = CreateSearchPattern(searchPatterns[i]);
							Regex regexPattern = new Regex(pattern, RegexOptions.IgnoreCase);
							Match match = regexPattern.Match(docText);

							if (match.Success)
							{
								string highlightedText = HighlightMatch(rates[i], paragraphs, match, docText, paragraphLengths, sb, isDebug);
								if (!this.CompareStringsIgnoringWhitespace(highlightedText, match.Value))
								{
									sb.AppendLine("警告: 色付け箇所と検索テキストが異なります。");
									sb.AppendLine($"検索テキスト: {match.Value}");
									sb.AppendLine($"色付け箇所: {highlightedText}");
								}
								else if (isDebug)
								{
									sb.AppendLine("色付け箇所と検索テキストが一致しました。");
									sb.AppendLine($"検索テキスト: {match.Value}");
									sb.AppendLine($"色付け箇所: {highlightedText}");
								}
							}
							else
							{
								sb.AppendLine("エラー: 指定されたテキストが見つかりませんでした。");
								sb.AppendLine($"検索文: {searchPatterns[i]}\n{regexPattern}");
							}
						}
					}
					doc.MainDocumentPart.Document.Save();
				}
			}
			catch (Exception ex)
			{
				sb.AppendLine($"エラーが発生しました: {ex.Message}");
				sb.AppendLine($"スタックトレース: {ex.StackTrace}");
			}
			return sb.ToString();
		}

		private string CreateDocTextWithSpaces(List<Paragraph> paragraphs, out List<int> paragraphLengths)
		{
			StringBuilder sb = new StringBuilder();
			paragraphLengths = new List<int>();
			foreach (var paragraph in paragraphs)
			{
				string paragraphText = paragraph.InnerText.Trim();
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
				escaped = Regex.Replace(escaped, @"([,.:;])(?!$)", @"$1\s*");

				return escaped;
			}).ToArray();

			// Join the words with flexible whitespace
			return string.Join(@"\s*", processedWords);
		}

		private string HighlightMatch(double rate, List<Paragraph> paragraphs, Match match, string docText, List<int> paragraphLengths, StringBuilder sb, bool isDebug)
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
					if (isDebug)
					{
						string text = docText.Substring(paragraphStart, paragraphEnd - paragraphStart);
						sb.AppendLine($"partFromDocText:#{text}#");
						sb.AppendLine($"currentIndex: {currentIndex}\nparagraphText:#{paragraphText}#\nparagraphLength: {paragraphLength} paragraphStart: {paragraphStart} paragraphEnd: {paragraphEnd}");
						sb.AppendLine($"matchStart: {matchStart} matchEnd: {matchEnd}  match.Length: {match.Length} startInParagraph: {startInParagraph} endInParagraph: {endInParagraph}");
					}
					string matchedText = paragraph.InnerText.Substring(startInParagraph, endInParagraph - startInParagraph);
					if (!this.CompareStringsIgnoringWhitespace(matchedText, match.Value))
					{
						string head = match.Value.Substring(0, Math.Min(10, match.Value.Length));
						int pos = paragraphText.IndexOf(head);
						if (0 <= pos && startInParagraph != pos)
						{
							startInParagraph = pos;
							endInParagraph = pos + match.Value.Length;
							if (isDebug) sb.AppendLine($"pos:{pos}\nmatch.Value:#{match.Value}#\nmatchedText:#{matchedText}#\nstartInParagraph:{startInParagraph}\nendInParagraph:{endInParagraph}\nparagraph.InnerText:#{paragraph.InnerText}#");
							matchedText = paragraph.InnerText.Substring(startInParagraph, Math.Min(endInParagraph - startInParagraph, paragraph.InnerText.Length - startInParagraph));
						}
					}
					ApplyBackgroundColorToParagraph(paragraph, rate, startInParagraph, endInParagraph);
					highlightedText.Append(matchedText);
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

		private void ApplyBackgroundColor(double rate, Run run)
		{
			Color color = GetHighlightColor(rate);
			RunProperties runProperties = run.RunProperties;
			if (runProperties == null)
			{
				runProperties = new RunProperties();
				run.PrependChild(runProperties);
			}

			Shading shading = runProperties.Shading;
			if (shading == null)
			{
				shading = new Shading();
				runProperties.AppendChild(shading);
			}

			shading.Fill = $"{color.R:X2}{color.G:X2}{color.B:X2}";
			shading.Color = "auto";
			shading.Val = ShadingPatternValues.Clear;
		}

		private void ApplyBackgroundColorToParagraph(Paragraph paragraph, double rate, int startIndex, int endIndex)
		{
			int currentIndex = 0;
			List<Run> newRuns = new List<Run>();

			foreach (var run in paragraph.Elements<Run>().ToList())
			{
				string runText = run.InnerText;
				int runLength = runText.Length;
				int runStart = currentIndex;
				int runEnd = runStart + runLength;

				if (runEnd > startIndex && runStart < endIndex)
				{
					int colorStart = Math.Max(0, startIndex - runStart);
					int colorEnd = Math.Min(runLength, endIndex - runStart);

					if (colorStart > 0)
					{
						Run beforeRun = new Run(new Text(runText.Substring(0, colorStart)) { Space = SpaceProcessingModeValues.Preserve });
						CopyRunProperties(run, beforeRun);
						newRuns.Add(beforeRun);
					}

					Run coloredRun = new Run(new Text(runText.Substring(colorStart, colorEnd - colorStart)) { Space = SpaceProcessingModeValues.Preserve });
					CopyRunProperties(run, coloredRun);
					ApplyBackgroundColor(rate, coloredRun);
					newRuns.Add(coloredRun);

					if (colorEnd < runLength)
					{
						Run afterRun = new Run(new Text(runText.Substring(colorEnd)) { Space = SpaceProcessingModeValues.Preserve });
						CopyRunProperties(run, afterRun);
						newRuns.Add(afterRun);
					}
				}
				else
				{
					newRuns.Add((Run)run.Clone());
				}

				currentIndex += runLength;
			}

			paragraph.RemoveAllChildren<Run>();
			foreach (var newRun in newRuns)
			{
				paragraph.AppendChild(newRun);
			}
		}

		private void CopyRunProperties(Run sourceRun, Run targetRun)
		{
			if (sourceRun.RunProperties != null)
			{
				targetRun.RunProperties = (RunProperties)sourceRun.RunProperties.Clone();
			}
		}

		private Color GetHighlightColor(double rate)
		{
			if (1D == rate) return Color.LightPink;
			else if (0.9 <= rate) return Color.Cyan;
			else if (0D < rate) return Color.LightGreen;
			return Color.White;
		}

		private bool CompareStringsIgnoringWhitespace(string str1, string str2)
		{
			// 正規表現を使用して全ての種類の空白を削除
			string pattern = @"\s+";
			string str1WithoutWhitespace = Regex.Replace(str1, pattern, "");
			string str2WithoutWhitespace = Regex.Replace(str2, pattern, "");

			// 空白を除去した文字列を比較
			return str1WithoutWhitespace.Equals(str2WithoutWhitespace);
		}
	}
}
