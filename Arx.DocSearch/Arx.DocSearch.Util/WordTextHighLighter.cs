using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;

namespace Arx.DocSearch.Util
{
	public class WordTextHighLighter
	{
		public string HighlightTextInWord(string filePath, int[] indexes, double[] rates, string[] searchPatterns, bool isDebug = false)
		{
			StringBuilder sb = new StringBuilder();
			CleanDocument(filePath, filePath);
			WordTextExtractor wte = new WordTextExtractor(filePath, true, false);
			try
			{
				using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
				{
					Body body = doc.MainDocumentPart.Document.Body;
					if (body != null)
					{
						var paragraphs = body.Descendants<Paragraph>().ToList();
						/*for (int i = 0; i < paragraphs.Count && i< wte.ParagraphTexts.Count; i++) {
							string text1 = paragraphs[i].InnerText;
							string text2 = wte.ParagraphTexts[i];
							//if (text1 != text2) sb.AppendLine($"index: {i}\nparagraphs:\n{text1}\nwte.ParagraphTexts:\n{text2}");
							sb.AppendLine($"index: {i}: {text2}");
						}*/
						string docText = CreateDocTextWithSpaces(wte.ParagraphTexts, out List<int> paragraphLengths);
						//docText = TextConverter.ZenToHan(docText ?? "");
						//docText = TextConverter.HankToZen(docText ?? "");
						if (isDebug) sb.AppendLine($"filePath:\n{filePath}\ndocText:\n{wte.Text}");
						for (int i = 0; i < searchPatterns.Length && i < rates.Length; i++)
						{
							string searchPattern = Regex.Replace(searchPatterns[i], @"^[0-9]+\.?\s+", "");
							searchPattern = Regex.Replace(searchPattern, @"\s+[0-9]+\.?\s*$", "");
							//sb.AppendLine($"searchPattern: index:{i}\n{searchPattern}");
							string[] words = searchPattern.Split(' ');
							if (searchPattern.Length < 20 && words.Length < 3) continue;
							string pattern = CreateSearchPattern(searchPattern);
							var results = MatchIgnoringWhitespace(pattern, docText, sb);
							//sb.AppendLine($"results.Count: {results.Count}");
							if (results.Count > 0)
							{
								if (isDebug)
								{
									sb.AppendLine("エラー: 指定されたテキストが見つかりました。");
									sb.AppendLine($"検索文: {searchPatterns[i]}\n{pattern}");
								}
								foreach (var result in results)
								{
									List<int[]> matchedParagraphs = GetMatchedParagraphs(result.beginIndex, result.endIndex, paragraphLengths, docText, sb);
									//色付け処理の位置を確認するための予備検索。検索もれを防ぐために50文字余分に検索する。
									string matchedText = SafeSubstring(docText, result.beginIndex, result.endIndex - result.beginIndex + 51);
									matchedText = Regex.Replace(matchedText, @"F[0-9A-F]{3}|[<>]", @"");
									//色付け処理の位置を確認するための予備検索。検索もれを防ぐために50文字余分に検索する。
									string[] ret = ExecuteHighlightMatch(rates[i], paragraphs, paragraphLengths, matchedParagraphs, result.beginIndex, result.endIndex + 50, false);
									string highlightedText = ret[0];
									string paragrapghText = ret[1];
									int? offset = StringOffsetCalculator.CalculateOffset(highlightedText, matchedText);
									if (isDebug && offset != 0) sb.AppendLine($"offset: {offset}\nhighlightedText: {highlightedText}\nmatchedText: {matchedText}");
									if (offset != null)
									{
										matchedText = SafeSubstring(docText, result.beginIndex, result.endIndex - result.beginIndex + 1);
										matchedText = Regex.Replace(matchedText, @"F[0-9A-F]{3}|[<>]", @"");
										string textNoSymbol = SpecialCharConverter.ReplaceMathSymbols(matchedText);
										int lengthDiff = matchedText.Length - textNoSymbol.Length;
										ret = ExecuteHighlightMatch(rates[i], paragraphs, paragraphLengths, matchedParagraphs, result.beginIndex + offset.Value, result.endIndex + offset.Value - lengthDiff, true);
										highlightedText = ret[0];
										paragrapghText = ret[1];
									}
									bool colorMatched = true;
									if (!CompareStringsIgnoringWhitespace(highlightedText, matchedText))
									{
										sb.AppendLine("警告: 色付け箇所と検索テキストが異なります。");
										sb.AppendLine($"検索テキスト: {matchedText}");
										sb.AppendLine($"色付け箇所: {highlightedText}");
										sb.AppendLine($"paragrapghText: {paragrapghText}");
										colorMatched = false;
									}
									if (isDebug && colorMatched)
									{
										sb.AppendLine("色付け箇所と検索テキストが一致しました。");
										sb.AppendLine($"検索テキスト: {matchedText}");
										sb.AppendLine($"色付け箇所: {highlightedText}");
									}
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
				escaped = Regex.Replace(escaped, @"'", @"[‘’']");
				escaped = Regex.Replace(escaped, @"""", @"[“”®™–—""]");
				escaped = Regex.Replace(escaped, @"([,.:;()])(?!$)", @"$1\s*");

				// 修正箇所: エスケープされた文字と未エスケープの文字の両方に対応
				escaped = Regex.Replace(escaped, @"(\\[,.:;()])|([,.:;()])", m =>
				{
					if (m.Groups[1].Success) // エスケープされた文字の場合
						return @"\s*" + m.Groups[1].Value;
					else // エスケープされていない文字の場合
						return m.Groups[2].Value;
				});

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
				// 段落の末尾を含む場合、または段落の開始位置を含む場合、
				// または段落が検索範囲内にある場合に追加
				if ((index <= beginIndex && beginIndex <= index + paragraphLength) ||
					(beginIndex < index && index + paragraphLength < endIndex) ||
					(index <= endIndex && endIndex <= index + paragraphLength) ||
					// 以下の条件を追加
					(beginIndex < index && index < endIndex))
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

		private string[] ExecuteHighlightMatch(double rate, List<Paragraph> paragraphs, List<int> paragraphLengths, List<int[]> matchedParagraphs, int beginIndex, int endIndex, bool isFinal)
		{
			string[] ret = HighlightMatch(rate, paragraphs, paragraphLengths, matchedParagraphs, beginIndex, endIndex, isFinal);
			string highlightedText = ret[0];
			string paragraphText = ret[1];
			highlightedText = TextConverter.ZenToHan(highlightedText ?? "");
			highlightedText = TextConverter.HankToZen(highlightedText ?? "");
			paragraphText = TextConverter.ZenToHan(paragraphText ?? "");
			paragraphText = TextConverter.HankToZen(paragraphText ?? "");
			highlightedText = Regex.Replace(highlightedText, @"F[0-9A-F]{3}|[<>]", @"");
			string[] ret2 = new string[3];
			ret[0] = highlightedText;
			ret[1] = paragraphText;
			return ret;
		}

		private string[] HighlightMatch(double rate, List<Paragraph> paragraphs, List<int> paragraphLengths, List<int[]> matchedParagraphs, int beginIndex, int endIndex, bool isFinal)
		{
			StringBuilder highlightedText = new StringBuilder();
			StringBuilder paragraphText = new StringBuilder();
			foreach (int[] matchedParagraph in matchedParagraphs)
			{
				int index = matchedParagraph[0];
				int pos = matchedParagraph[1];
				Paragraph paragraph = paragraphs[index];
				int paragraphLength = paragraph.InnerText.Length;

				// パラグラフ内での相対位置を計算
				int relativeStart = beginIndex - pos;
				int relativeEnd = endIndex - pos;

				// パラグラフの境界をまたぐ場合の処理
				// 前のパラグラフの末尾部分
				if (relativeStart >= paragraphLength && relativeEnd > paragraphLength)
				{
					relativeStart = Math.Max(0, paragraphLength - 3);
					relativeEnd = paragraphLength;
				}
				// 次のパラグラフの先頭部分
				else if (relativeStart < 0 && relativeEnd > 0)
				{
					relativeStart = 0;
					relativeEnd = Math.Min(paragraphLength, relativeEnd);
				}

				if (relativeEnd > 0 && relativeStart < paragraphLength)
				{
					int effectiveStart = Math.Max(0, relativeStart);
					int effectiveEnd = Math.Min(paragraphLength, relativeEnd);

					if (effectiveStart < effectiveEnd)
					{
						string matchedText = SafeSubstring(paragraph.InnerText, effectiveStart, effectiveEnd - effectiveStart);
						highlightedText.Append(matchedText);
						if (isFinal) ApplyBackgroundColorToParagraph(paragraph, rate, effectiveStart, effectiveEnd);
					}
				}
			}
			string[] ret = new string[2];
			ret[0] = highlightedText.ToString();
			ret[1] = paragraphText.ToString();

			return ret;
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

			// RunPropertiesの作成を明示的に行う
			if (run.RunProperties == null)
			{
				run.RunProperties = new RunProperties();
			}

			// 既存のShadingを削除して新しく作成
			var existingShading = run.RunProperties.GetFirstChild<Shading>();
			if (existingShading != null)
			{
				existingShading.Remove();
			}

			// 新しいShadingを作成して設定
			Shading shading = new Shading()
			{
				Fill = $"{color.R:X2}{color.G:X2}{color.B:X2}",
				Color = "auto",
				Val = ShadingPatternValues.Clear
			};

			// RunPropertiesの先頭に追加
			run.RunProperties.InsertAt(shading, 0);

			//Console.WriteLine($"Debug - Applied color {color.R:X2}{color.G:X2}{color.B:X2} to text: {run.InnerText}");
		}

		private void ApplyBackgroundColorToParagraph(Paragraph paragraph, double rate, int startIndex, int endIndex)
		{
			string paragraphText = paragraph.InnerText;
			List<(int start, int end, OpenXmlElement element)> elementRanges = new List<(int, int, OpenXmlElement)>();
			int currentIndex = 0;
			// 各Runの範囲を記録
			foreach (var element in paragraph.Elements())
			{
				int elementLength = element.InnerText.Length;
				elementRanges.Add((currentIndex, currentIndex + elementLength, element));
				currentIndex += elementLength;
			}
			List<OpenXmlElement> newElements = new List<OpenXmlElement>();
			foreach (var (elemStart, elemEnd, element) in elementRanges)
			{
				if (element is Run run)
				{
					// 既存のRun処理コード
					if (DoRangesOverlap(elemStart, elemEnd - 1, startIndex, endIndex))
					{
						int colorStart = Math.Max(0, startIndex - elemStart);
						int colorEnd = Math.Min(run.InnerText.Length, endIndex - elemStart + 1);

						if (colorStart == 0 && colorEnd == run.InnerText.Length)
						{
							Run newRun = (Run)run.CloneNode(true);
							ApplyBackgroundColor(rate, newRun);
							newElements.Add(newRun);
						}
						else
						{
							SplitAndColorRun(run, colorStart, colorEnd, rate, newElements);
						}
					}
					else
					{
						newElements.Add((Run)run.CloneNode(true));
					}
				}
				else if (element is DocumentFormat.OpenXml.Math.OfficeMath officeMath)
				{
					if (DoRangesOverlap(elemStart, elemEnd - 1, startIndex, endIndex))
					{
						// OfficeMath要素全体をRunに変換して色付け
						Run newRun = new Run();
						newRun.AppendChild(new Text(officeMath.InnerText));
						ApplyBackgroundColor(rate, newRun);
						newElements.Add(newRun);
					}
					else
					{
						newElements.Add(element.CloneNode(true));
					}
				}
				else
				{
					newElements.Add(element.CloneNode(true));
				}
			}

			// 古い要素を削除して新しい要素を追加
			paragraph.RemoveAllChildren();
			foreach (var newElement in newElements)
			{
				paragraph.AppendChild(newElement);
			}
		}

		private bool DoRangesOverlap(int start1, int end1, int start2, int end2)
		{
			bool overlaps = start1 <= end2 && start2 <= end1;
			return overlaps;
		}

		private void SplitAndColorRun(Run originalRun, int colorStart, int colorEnd, double rate, List<OpenXmlElement> newElements)
		{
			// 色付け前の部分
			if (colorStart > 0)
			{
				Run beforeRun = new Run();
				CopyRunProperties(originalRun, beforeRun);
				beforeRun.AppendChild(new Text(originalRun.InnerText.Substring(0, colorStart)));
				newElements.Add(beforeRun);
			}

			// 色付け部分
			Run coloredRun = new Run();
			CopyRunProperties(originalRun, coloredRun);
			coloredRun.AppendChild(new Text(originalRun.InnerText.Substring(colorStart, colorEnd - colorStart)));
			ApplyBackgroundColor(rate, coloredRun);
			newElements.Add(coloredRun);

			// 色付け後の部分
			if (colorEnd < originalRun.InnerText.Length)
			{
				Run afterRun = new Run();
				CopyRunProperties(originalRun, afterRun);
				afterRun.AppendChild(new Text(originalRun.InnerText.Substring(colorEnd)));
				newElements.Add(afterRun);
			}
		}

		private void CopyRunProperties(Run sourceRun, Run targetRun)
		{
			if (sourceRun.RunProperties != null)
			{
				targetRun.RunProperties = (RunProperties)sourceRun.RunProperties.CloneNode(true);
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
			string str1WithoutWhitespace = SpecialCharConverter.ReplaceMathSymbols(str1);
			string str2WithoutWhitespace = SpecialCharConverter.ReplaceMathSymbols(str2);
			str1WithoutWhitespace = SpecialCharConverter.ReplaceLine(str1WithoutWhitespace ?? "");
			str2WithoutWhitespace = SpecialCharConverter.ReplaceLine(str2WithoutWhitespace ?? "");
			str1WithoutWhitespace = SpecialCharConverter.RemoveSymbols(str1WithoutWhitespace ?? "");
			str2WithoutWhitespace = SpecialCharConverter.RemoveSymbols(str2WithoutWhitespace ?? "");
			str1WithoutWhitespace = Regex.Replace(str1WithoutWhitespace, pattern, "");
			str2WithoutWhitespace = Regex.Replace(str2WithoutWhitespace, pattern, "");

			// 空白、記号を除去した文字列を比較
			if (0 <= str1WithoutWhitespace.IndexOf(str2WithoutWhitespace) || 0 <= str2WithoutWhitespace.IndexOf(str1WithoutWhitespace)) return true;
			else return false;
		}

		public List<(int beginIndex, int endIndex)> MatchIgnoringWhitespace(string pattern, string text, StringBuilder sb)
		{
			try
			{
				Regex regex = new Regex(pattern, RegexOptions.Compiled | RegexOptions.Multiline);
				MatchCollection matches = regex.Matches(text);

				List<(int beginIndex, int endIndex)> result = new List<(int beginIndex, int endIndex)>();

				foreach (Match match in matches)
				{
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

					result.Add((beginIndex, endIndex));
				}

				return result;
			}
			catch (ArgumentException ex)
			{
				sb.AppendLine($"正規表現エラー: {ex.Message}");
				sb.AppendLine($"スタックトレース: {ex.StackTrace}");
				return new List<(int beginIndex, int endIndex)>();
			}
		}

		private void ColorMathElementRecursive(OpenXmlElement element, double rate, int startIndex, int endIndex)
		{
			int currentIndex = 0;
			foreach (var child in element.Elements().ToList())
			{
				if (child is Run run)
				{
					string runText = run.InnerText;
					int runLength = runText.Length;
					if (currentIndex + runLength > startIndex && currentIndex < endIndex)
					{
						int colorStart = Math.Max(0, startIndex - currentIndex);
						int colorEnd = Math.Min(runLength, endIndex - currentIndex);

						var newRun = (Run)run.Clone();
						newRun.RemoveAllChildren();

						if (colorStart > 0)
						{
							newRun.AppendChild(new Text(runText.Substring(0, colorStart)));
						}

						var coloredRun = (Run)run.Clone();
						coloredRun.RemoveAllChildren();
						coloredRun.AppendChild(new Text(runText.Substring(colorStart, colorEnd - colorStart)));
						ApplyBackgroundColor(rate, coloredRun);
						newRun.AppendChild(coloredRun);

						if (colorEnd < runLength)
						{
							newRun.AppendChild(new Text(runText.Substring(colorEnd)));
						}

						element.ReplaceChild(newRun, run);
					}
					currentIndex += runLength;
				}
				else
				{
					ColorMathElementRecursive(child, rate, startIndex - currentIndex, endIndex - currentIndex);
					currentIndex += child.InnerText.Length;
				}
			}
		}

		public static void CleanDocument(string inputFilePath, string outputFilePath)
		{
			try
			{
				if (inputFilePath == outputFilePath)
				{
					// 同じファイルの場合、直接編集
					using (WordprocessingDocument doc = WordprocessingDocument.Open(inputFilePath, true))
					{
						CleanDocumentContent(doc);
						doc.Save();
					}
				}
				else
				{
					// 新しいファイルとして保存する場合
					File.Copy(inputFilePath, outputFilePath, true);
					using (WordprocessingDocument doc = WordprocessingDocument.Open(outputFilePath, true))
					{
						CleanDocumentContent(doc);
						doc.Save();
					}
				}
			}
			catch (Exception ex)
			{
			}
		}

		private static void CleanDocumentContent(WordprocessingDocument doc)
		{
			RemoveFieldCodesInDocument(doc);
			ClearBackgroundAndHighlight(doc);
		}

		private static void RemoveFieldCodesInDocument(WordprocessingDocument doc)
		{
			var body = doc.MainDocumentPart.Document.Body;
			if (body != null)
			{
				RemoveFieldCodesInElement(body);
			}

			// ヘッダーとフッターも処理
			var headerParts = doc.MainDocumentPart.HeaderParts;
			foreach (var headerPart in headerParts)
			{
				RemoveFieldCodesInElement(headerPart.Header);
			}

			var footerParts = doc.MainDocumentPart.FooterParts;
			foreach (var footerPart in footerParts)
			{
				RemoveFieldCodesInElement(footerPart.Footer);
			}
		}

		private static void RemoveFieldCodesInElement(OpenXmlElement element)
		{
			var runs = element.Descendants<Run>().ToList();
			foreach (var run in runs)
			{
				var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
				if (fieldChar != null && fieldChar.FieldCharType == FieldCharValues.Begin)
				{
					var fieldCode = run.NextSibling<Run>()?.GetFirstChild<FieldCode>();
					if (fieldCode != null)
					{
						string fieldCodeText = fieldCode.InnerText;
						var nextRun = run.NextSibling<Run>();
						while (nextRun != null)
						{
							var endFieldChar = nextRun.Elements<FieldChar>().FirstOrDefault(fc => fc.FieldCharType == FieldCharValues.End);
							if (endFieldChar != null)
							{
								break;
							}
							nextRun = nextRun.NextSibling<Run>();
						}
						if (nextRun != null)
						{
							// フィールドの結果を保持
							string result = GetFieldResult(run, nextRun);
							run.RemoveAllChildren();
							run.AppendChild(new Text(result));

							// フィールドコードの残りの部分を削除
							while (run.NextSibling<Run>() != nextRun)
							{
								run.NextSibling<Run>().Remove();
							}
							nextRun.Remove();
						}
					}
				}
			}
		}

		private static string GetFieldResult(Run startRun, Run endRun)
		{
			string result = "";
			var currentRun = startRun.NextSibling<Run>();
			while (currentRun != null && currentRun != endRun)
			{
				var text = currentRun.GetFirstChild<Text>();
				if (text != null)
				{
					result += text.Text;
				}
				currentRun = currentRun.NextSibling<Run>();
			}
			return result.Trim();
		}

		private static void ClearBackgroundAndHighlight(WordprocessingDocument doc)
		{
			var body = doc.MainDocumentPart.Document.Body;
			if (body != null)
			{
				ClearBackgroundAndHighlightInElement(body);
			}

			// ヘッダーとフッターも処理
			var headerParts = doc.MainDocumentPart.HeaderParts;
			foreach (var headerPart in headerParts)
			{
				ClearBackgroundAndHighlightInElement(headerPart.Header);
			}

			var footerParts = doc.MainDocumentPart.FooterParts;
			foreach (var footerPart in footerParts)
			{
				ClearBackgroundAndHighlightInElement(footerPart.Footer);
			}
		}

		private static void ClearBackgroundAndHighlightInElement(OpenXmlElement element)
		{
			var runs = element.Descendants<Run>().ToList();
			foreach (var run in runs)
			{
				var runProperties = run.RunProperties;
				if (runProperties != null)
				{
					// 背景色をクリア
					var shading = runProperties.GetFirstChild<Shading>();
					if (shading != null)
					{
						shading.Remove();
					}

					// ハイライトをクリア
					var highlight = runProperties.GetFirstChild<Highlight>();
					if (highlight != null)
					{
						highlight.Remove();
					}
				}
			}
		}
	}
}
