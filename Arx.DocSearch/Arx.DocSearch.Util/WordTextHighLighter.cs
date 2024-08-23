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
			StringBuilder sb = new StringBuilder();
			WordTextExtractor wte = new WordTextExtractor(filePath, true, false);
			try
			{
				using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
				{
					Body body = doc.MainDocumentPart.Document.Body;
					if (body != null)
					{
						var paragraphs = body.Descendants<Paragraph>().ToList();
						string docText = CreateDocTextWithSpaces(wte.ParagraphTexts, out List<int> paragraphLengths);
						for (int i = 0; i < searchPatterns.Length && i < rates.Length; i++)
						{
							string pattern = CreateSearchPattern(searchPatterns[i]);
							var result = MatchIgnoringWhitespace(pattern, docText, sb);
							if (result.matched)
							{
								List<int[]> matchedParagraphs = GetMatchedParagraphs(result.beginIndex, result.endIndex, paragraphLengths, docText, sb);
								string highlightedText = HighlightMatch(rates[i], paragraphs, paragraphLengths, matchedParagraphs, result.beginIndex, result.endIndex, sb, isDebug);
								string matchedText = docText.Substring(result.beginIndex, result.endIndex - result.beginIndex + 1);

								if (!CompareStringsIgnoringWhitespace(highlightedText, matchedText))
								{
									sb.AppendLine("警告: 色付け箇所と検索テキストが異なります。");
									sb.AppendLine($"検索テキスト: {matchedText}");
									sb.AppendLine($"色付け箇所: {highlightedText}");
								}
								else if (isDebug)
								{
									sb.AppendLine("色付け箇所と検索テキストが一致しました。");
									sb.AppendLine($"検索テキスト: {matchedText}");
									sb.AppendLine($"色付け箇所: {highlightedText}");
								}
							}
							else
							{
								sb.AppendLine("エラー: 指定されたテキストが見つかりませんでした。");
								sb.AppendLine($"検索文: {searchPatterns[i]}\n{pattern}");
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

		private string CreateDocTextWithSpaces(List<string> paragraphTexts, out List<int> paragraphLengths)
		{
			StringBuilder sb = new StringBuilder();
			paragraphLengths = new List<int>();
			foreach (string paragraphText in paragraphTexts)
			{
				if (!string.IsNullOrWhiteSpace(paragraphText))
				{
					sb.Append(paragraphText);
					paragraphLengths.Add(paragraphText.Length);
				}
				else paragraphLengths.Add(0);
			}
			return sb.ToString();
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

		private List<int[]> GetMatchedParagraphs(int beginIndex, int endIndex, List<int> paragraphLengths, string docText, StringBuilder sb)
		{
			int index = 0;
			List<int[]> matcheParagraphs = new List<int[]>();
			for (int i = 0; i < paragraphLengths.Count; i++)
			{
				int paragraphLength = paragraphLengths[i];
				if (index <= beginIndex && beginIndex <= index + paragraphLength
				|| beginIndex < index && index + paragraphLength < endIndex
				|| index <= endIndex && endIndex <= index + paragraphLength)
				{
					int[] info = new int[2];
					info[0] = i;
					info[1] = index;
					matcheParagraphs.Add(info);
				}
				index += paragraphLength;
				if (endIndex < index) break;
			}
			return matcheParagraphs;
		}

		private string HighlightMatch(double rate, List<Paragraph> paragraphs, List<int> paragraphLengths, List<int[]> matchedParagraphs, int beginIndex, int endIndex, StringBuilder sb, bool isDebug)
		{
			StringBuilder highlightedText = new StringBuilder();
			foreach (int[] matchedParagraph in matchedParagraphs)
			{
				int index = matchedParagraph[0];
				int pos = matchedParagraph[1];
				Paragraph paragraph = paragraphs[index];

				// SpecialCharConverterを使用
				string convertedParagraphText = SpecialCharConverter.ConvertSpecialCharactersInParagraph(paragraph);
				if (isDebug) sb.AppendLine($"index: {index} pos: {pos} beginIndex: {beginIndex} convertedParagraphText: {convertedParagraphText}");

				int startInParagraph = Math.Max(0, beginIndex - pos);
				int endInParagraph = Math.Min(convertedParagraphText.Length, endIndex - pos + 1);

				if (startInParagraph < convertedParagraphText.Length && endInParagraph > 0)
				{
					string matchedText = SafeSubstring(convertedParagraphText, startInParagraph, endInParagraph - startInParagraph);
					highlightedText.Append(matchedText);
					ApplyBackgroundColorToParagraph(paragraph, rate, startInParagraph, endInParagraph);
				}
			}

			return highlightedText.ToString();
		}

		public string SafeSubstring(string str, int startIndex, int length)
		{
			if (string.IsNullOrEmpty(str))
				return string.Empty;

			// startIndexを0以上、文字列の長さ未満に調整
			startIndex = Math.Max(0, Math.Min(str.Length - 1, startIndex));

			// lengthを0以上、残りの文字列の長さ以下に調整
			length = Math.Max(0, Math.Min(str.Length - startIndex, length));

			return str.Substring(startIndex, length);
		}

		public string SafeSubstring(string str, int startIndex)
		{
			if (string.IsNullOrEmpty(str))
				return string.Empty;

			// startIndexを0以上、文字列の長さ以下に調整
			startIndex = Math.Max(0, Math.Min(str.Length, startIndex));

			return str.Substring(startIndex);
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
				string runText = SpecialCharConverter.ConvertSpecialCharactersInRun(run);
				int runLength = runText.Length;
				int runStart = currentIndex;
				int runEnd = runStart + runLength;

				if (runEnd > startIndex && runStart < endIndex)
				{
					int colorStart = Math.Max(0, startIndex - runStart);
					int colorEnd = Math.Min(runLength, endIndex - runStart);

					if (colorStart > 0)
					{
						Run beforeRun = CreateNewRun(run, runText.Substring(0, colorStart));
						newRuns.Add(beforeRun);
					}

					Run coloredRun = CreateNewRun(run, runText.Substring(colorStart, colorEnd - colorStart));
					ApplyBackgroundColor(rate, coloredRun);
					newRuns.Add(coloredRun);

					if (colorEnd < runLength)
					{
						Run afterRun = CreateNewRun(run, runText.Substring(colorEnd));
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

		private Run CreateNewRun(Run originalRun, string text)
		{
			Run newRun = new Run();
			CopyRunProperties(originalRun, newRun);

			foreach (var child in originalRun.ChildElements)
			{
				if (child is Text)
				{
					newRun.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
				}
				else if (child.LocalName == "sym")
				{
					// シンボル要素を適切に処理
					var symElement = (OpenXmlElement)child.Clone();
					newRun.AppendChild(symElement);
				}
				else
				{
					newRun.AppendChild((OpenXmlElement)child.Clone());
				}
			}

			return newRun;
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

		public (bool matched, int beginIndex, int endIndex) MatchIgnoringWhitespace(string pattern, string text, StringBuilder sb)
		{
			try
			{
				Regex regex = new Regex(pattern, RegexOptions.Compiled | RegexOptions.Multiline);
				Match match = regex.Match(text);

				if (!match.Success)
				{
					return (false, -1, -1);
				}

				int beginIndex = match.Index;
				int endIndex = match.Index + match.Length - 1;

				// 先頭の空白をスキップ
				while (beginIndex <= endIndex && char.IsWhiteSpace(text[beginIndex]))
				{
					beginIndex++;
				}

				// 末尾の空白をスキップ
				while (endIndex >= beginIndex && char.IsWhiteSpace(text[endIndex]))
				{
					endIndex--;
				}

				return (true, beginIndex, endIndex);
			}
			catch (ArgumentException ex)
			{
				sb.AppendLine($"正規表現エラー: {ex.Message}");
				sb.AppendLine($"スタックトレース: {ex.StackTrace}");
				return (false, -1, -1);
			}
		}
	}
}
