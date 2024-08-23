using System.Text;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;
using System;

namespace Arx.DocSearch.Util
{
	public class WordTextExtractor
	{
		public WordTextExtractor(string filePath, bool isSingleLine = true, bool reducesBlankSpaces = true, Action<string> debugLogger = null)
		{
			IsSingleLine = isSingleLine;
			ReducesBlankSpaces = reducesBlankSpaces;
			DebugLogger = debugLogger ?? (_ => { }); // デフォルトは何もしない

			ExtractText(filePath);
		}
		private Action<string> DebugLogger { get; set; }

		private static StringBuilder extractedText = new StringBuilder();
		private bool EnableDebugOutput { get; set; }
		public bool IsSingleLine { get; set; }
		public bool ReducesBlankSpaces { get; set; }

		public List<string> ParagraphTexts { get; private set; }

		public string Text
		{
			get { return CombineText(ParagraphTexts); }
		}

		private void ExtractText(string filePath)
		{
			ParagraphTexts = new List<string>();

			using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
			{
				var body = doc.MainDocumentPart.Document.Body;
				if (body != null)
				{
					ExtractBodyElements(body.ChildElements);
				}
			}
		}

		private void ExtractBodyElements(IEnumerable<OpenXmlElement> elements)
		{
			foreach (var element in elements)
			{
				if (element is Paragraph paragraph)
				{
					string paragraphText = SpecialCharConverter.ConvertSpecialCharactersInParagraph(paragraph);
					ParagraphTexts.Add(paragraphText);
				}
				else if (element is Table table)
				{
					ExtractTableElements(table.ChildElements);
				}
			}
		}


		// テーブル要素を処理
		private void ExtractTableElements(IEnumerable<OpenXmlElement> elements)
		{
			foreach (var element in elements)
			{
				if (element is TableRow row)
				{
					foreach (var cell in row.Elements<TableCell>())
					{
						foreach (var cellChild in cell.ChildElements)
						{
							if (cellChild is Paragraph paragraph)
							{
								string cellText = SpecialCharConverter.ConvertSpecialCharactersInParagraph(paragraph);
								ParagraphTexts.Add(cellText);
							}
						}
					}
				}
			}
		}

		private string CleanupText(string text)
		{
			// 余分な空白の削除（ただし、ハイフンの前後の空白は保持）
			text = Regex.Replace(text, @"(?<!\s-)[^\S\n\r]+(?!-\s)", " ");
			// 行頭と行末の空白を削除
			text = Regex.Replace(text, @"^\s+|\s+$", "", RegexOptions.Multiline);
			// 連続する改行を1つにまとめる
			text = Regex.Replace(text, @"\n+", "\n");
			// 段落番号の後に余分な数字がある場合、それを削除（ただし先頭の0は保持）
			return text;
		}

		private string CombineText(List<string> paragraphs)
		{
			StringBuilder combinedText = new StringBuilder();

			foreach (var paragraph in paragraphs)
			{
				if (!string.IsNullOrEmpty(paragraph))
				{
					if (combinedText.Length > 0)
					{
						if (!IsSingleLine) combinedText.Append(Environment.NewLine);
					}
					combinedText.Append(paragraph);
				}
			}

			string result = combinedText.ToString();

			if (ReducesBlankSpaces)
			{
				result = CleanupText(result);
			}

			return result;
		}

		// 設定を変更するメソッド
		public void UpdateSettings(bool isSingleLine, bool reducesBlankSpaces)
		{
			IsSingleLine = isSingleLine;
			ReducesBlankSpaces = reducesBlankSpaces;
		}

		private void DebugOutput(string message, int depth)
		{
			DebugLogger($"{new string(' ', depth * 2)}{message}");
		}

		public void SetDebugLogger(Action<string> debugLogger)
		{
			DebugLogger = debugLogger ?? (_ => { });
		}

	}


}
