using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;
using Microsoft.International.Converters;

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
									string matchedText = docText.Substring(result.beginIndex, result.endIndex - result.beginIndex + 1);
									matchedText = Regex.Replace(matchedText, @"F[0-9A-F]{3}|[<>]", @"");
									string[] ret = ExecuteHighlightMatch(rates[i], paragraphs, paragraphLengths, matchedParagraphs, result.beginIndex, result.endIndex, false);
									string highlightedText = ret[0];
									string paragrapghText = ret[1];
									int? offset = StringOffsetCalculator.CalculateOffset(highlightedText, matchedText);
									if (isDebug && offset != 0) sb.AppendLine($"offset: {offset}\nhighlightedText: {highlightedText}\nmatchedText: {matchedText}");
									if (offset != null)
									{
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

				// SpecialCharConverterを使用
				; paragraphText.Append(paragraph.InnerText);

				int startInParagraph = Math.Max(0, beginIndex - pos);
				int endInParagraph = Math.Min(paragraph.InnerText.Length, endIndex - pos + 1);

				if (startInParagraph < paragraph.InnerText.Length && endInParagraph > 0)
				{
					string matchedText = SafeSubstring(paragraph.InnerText, startInParagraph, endInParagraph - startInParagraph);
					highlightedText.Append(matchedText);
					if (isFinal) ApplyBackgroundColorToParagraph(paragraph, rate, startInParagraph, endInParagraph);
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
			string paragraphText = paragraph.InnerText;
			List<(int start, int end, OpenXmlElement element)> elementRanges = new List<(int, int, OpenXmlElement)>();
			int currentIndex = 0;

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
					ProcessRunWithGlobalIndices(run, rate, elemStart, elemEnd, startIndex, endIndex, newElements);
				}
				else
				{
					newElements.Add(element.Clone() as OpenXmlElement);
				}
			}

			paragraph.RemoveAllChildren();
			foreach (var newElement in newElements)
			{
				paragraph.AppendChild(newElement);
			}
		}

		private void ProcessRunWithGlobalIndices(Run run, double rate, int runStart, int runEnd, int startIndex, int endIndex, List<OpenXmlElement> newElements)
		{
			string runText = run.InnerText;
			int runLength = runText.Length;

			if (runEnd <= startIndex || runStart >= endIndex)
			{
				newElements.Add(run.Clone() as Run);
				return;
			}

			int colorStart = Math.Max(0, startIndex - runStart);
			int colorEnd = Math.Min(runLength, endIndex - runStart);

			if (colorStart > 0)
			{
				Run beforeRun = CreateNewRun(run, runText.Substring(0, colorStart));
				newElements.Add(beforeRun);
			}

			if (colorEnd > colorStart)
			{
				Run coloredRun = CreateNewRun(run, runText.Substring(colorStart, colorEnd - colorStart));
				ApplyBackgroundColor(rate, coloredRun);
				newElements.Add(coloredRun);
			}

			if (colorEnd < runLength)
			{
				Run afterRun = CreateNewRun(run, runText.Substring(colorEnd));
				newElements.Add(afterRun);
			}
		}

		private void ProcessRun(Run run, double rate, ref int currentIndex, int startIndex, int endIndex, List<OpenXmlElement> newElements)
		{
			int runLength = CalculateRunLength(run);
			int runStart = currentIndex;
			int runEnd = runStart + runLength;

			//Console.WriteLine($"ProcessRun: run.InnerText='{run.InnerText}', calculatedLength={runLength}, currentIndex={currentIndex}, startIndex={startIndex}, endIndex={endIndex}");

			if (runEnd > startIndex && runStart < endIndex)
			{
				int colorStart = Math.Max(0, startIndex - runStart);
				int colorEnd = Math.Min(runLength, endIndex - runStart);

				//Console.WriteLine($"Color range: colorStart={colorStart}, colorEnd={colorEnd}");

				if (colorStart > 0)
				{
					Run beforeRun = CreateNewRun(run, GetSubRunText(run, 0, colorStart));
					newElements.Add(beforeRun);
				}

				Run coloredRun = CreateNewRun(run, GetSubRunText(run, colorStart, colorEnd));
				ApplyBackgroundColor(rate, coloredRun);
				newElements.Add(coloredRun);

				if (colorEnd < runLength)
				{
					Run afterRun = CreateNewRun(run, GetSubRunText(run, colorEnd, runLength));
					newElements.Add(afterRun);
				}
			}
			else
			{
				newElements.Add(CreateNewRun(run, run.InnerText));
			}

			currentIndex += runLength;
			Console.WriteLine($"ProcessRun completed: currentIndex={currentIndex}");
		}

		private int CalculateRunLength(Run run)
		{
			int length = 0;
			foreach (var child in run.ChildElements)
			{
				if (child is Text text)
				{
					length += text.Text.Length;
				}
				else if (child.LocalName == "sym")
				{
					length += 1; // シンボル要素は1文字としてカウント
				}
			}
			return length;
		}

		private string GetSubRunText(Run run, int start, int end)
		{
			StringBuilder sb = new StringBuilder();
			int currentIndex = 0;

			foreach (var child in run.ChildElements)
			{
				if (child is Text text)
				{
					int textLength = text.Text.Length;
					if (currentIndex + textLength > start && currentIndex < end)
					{
						int subStart = Math.Max(0, start - currentIndex);
						int subEnd = Math.Min(textLength, end - currentIndex);
						sb.Append(text.Text.Substring(subStart, subEnd - subStart));
					}
					currentIndex += textLength;
				}
				else if (child.LocalName == "sym")
				{
					if (currentIndex >= start && currentIndex < end)
					{
						sb.Append(GetSymbolChar(child));
					}
					currentIndex += 1;
				}

				if (currentIndex >= end) break;
			}

			return sb.ToString();
		}

		private string GetSymbolChar(OpenXmlElement symbolElement)
		{
			var charAttribute = symbolElement.GetAttributes().FirstOrDefault(a => a.LocalName == "char");
			if (charAttribute != null)
			{
				return SpecialCharConverter.ConvertSymbolChar(charAttribute.Value);
			}
			return "";
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

		private void ProcessMathElement(OpenXmlElement mathElement, double rate, ref int currentIndex, int startIndex, int endIndex, List<OpenXmlElement> newElements)
		{
			string mathText = ExtractTextFromMathElement(mathElement);
			int mathLength = mathText.Length;
			int mathStart = currentIndex;
			int mathEnd = mathStart + mathLength;

			if (mathEnd > startIndex && mathStart < endIndex)
			{
				var newMathElement = (OpenXmlElement)mathElement.Clone();
				ColorMathElementRecursive(newMathElement, rate, startIndex - mathStart, endIndex - mathStart);
				newElements.Add(newMathElement);
			}
			else
			{
				newElements.Add((OpenXmlElement)mathElement.Clone());
			}

			currentIndex += mathLength;
		}

		private string ExtractTextFromMathElement(OpenXmlElement mathElement)
		{
			return string.Join("", mathElement.Descendants<Text>().Select(t => t.Text));
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
	}
}
