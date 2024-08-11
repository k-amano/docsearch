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
		static readonly Dictionary<int, string> GreekCharMap = new Dictionary<int, string>
	{
        // 大文字ギリシャ文字
        { 0x0391, "Α" }, { 0x0392, "Β" }, { 0x0393, "Γ" }, { 0x0394, "Δ" },
		{ 0x0395, "Ε" }, { 0x0396, "Ζ" }, { 0x0397, "Η" }, { 0x0398, "Θ" },
		{ 0x0399, "Ι" }, { 0x039A, "Κ" }, { 0x039B, "Λ" }, { 0x039C, "Μ" },
		{ 0x039D, "Ν" }, { 0x039E, "Ξ" }, { 0x039F, "Ο" }, { 0x03A0, "Π" },
		{ 0x03A1, "Ρ" }, { 0x03A3, "Σ" }, { 0x03A4, "Τ" }, { 0x03A5, "Υ" },
		{ 0x03A6, "Φ" }, { 0x03A7, "Χ" }, { 0x03A8, "Ψ" }, { 0x03A9, "Ω" },
        // 小文字ギリシャ文字
        { 0x03B1, "α" }, { 0x03B2, "β" }, { 0x03B3, "γ" }, { 0x03B4, "δ" },
		{ 0x03B5, "ε" }, { 0x03B6, "ζ" }, { 0x03B7, "η" }, { 0x03B8, "θ" },
		{ 0x03B9, "ι" }, { 0x03BA, "κ" }, { 0x03BB, "λ" }, { 0x03BC, "μ" },
		{ 0x03BD, "ν" }, { 0x03BE, "ξ" }, { 0x03BF, "ο" }, { 0x03C0, "π" },
		{ 0x03C1, "ρ" }, { 0x03C2, "ς" }, { 0x03C3, "σ" }, { 0x03C4, "τ" },
		{ 0x03C5, "υ" }, { 0x03C6, "φ" }, { 0x03C7, "χ" }, { 0x03C8, "ψ" },
		{ 0x03C9, "ω" },
        // 追加の数学記号
        { 0x2206, "Δ" }, // INCREMENT
        { 0x2207, "∇" }, // NABLA
        { 0x2200, "∀" }, // FOR ALL
        { 0x2203, "∃" }, // THERE EXISTS
        { 0x2205, "∅" }, // EMPTY SET
        { 0x2208, "∈" }, // ELEMENT OF
        { 0x2209, "∉" }, // NOT AN ELEMENT OF
        { 0x220B, "∋" }, // CONTAINS AS MEMBER
        { 0x220F, "∏" }, // N-ARY PRODUCT
        { 0x2211, "∑" }, // N-ARY SUMMATION
        { 0x221A, "√" }, // SQUARE ROOT
        { 0x221D, "∝" }, // PROPORTIONAL TO
        { 0x221E, "∞" }, // INFINITY
        { 0x2229, "∩" }, // INTERSECTION
        { 0x222A, "∪" }, // UNION
        { 0x2248, "≈" }, // ALMOST EQUAL TO
        { 0x2260, "≠" }, // NOT EQUAL TO
        { 0x2264, "≤" }, // LESS-THAN OR EQUAL TO
        { 0x2265, "≥" }, // GREATER-THAN OR EQUAL TO
    };

		static readonly Dictionary<byte, string> SymbolToUnicode = new Dictionary<byte, string>
	{
		{0x41, "Α"}, {0x42, "Β"}, {0x47, "Γ"}, {0x44, "Δ"},
		{0x45, "Ε"}, {0x5A, "Ζ"}, {0x48, "Η"}, {0x51, "Θ"},
		{0x49, "Ι"}, {0x4B, "Κ"}, {0x4C, "Λ"}, {0x4D, "Μ"},
		{0x4E, "Ν"}, {0x58, "Ξ"}, {0x4F, "Ο"}, {0x50, "Π"},
		{0x52, "Ρ"}, {0x53, "Σ"}, {0x54, "Τ"}, {0x55, "Υ"},
		{0x46, "Φ"}, {0x43, "Χ"}, {0x59, "Ψ"}, {0x57, "Ω"},
		{0x61, "α"}, {0x62, "β"}, {0x67, "γ"}, {0x64, "δ"},
		{0x65, "ε"}, {0x7A, "ζ"}, {0x68, "η"}, {0x71, "θ"},
		{0x69, "ι"}, {0x6B, "κ"}, {0x6C, "λ"}, {0x6D, "μ"},
		{0x6E, "ν"}, {0x78, "ξ"}, {0x6F, "ο"}, {0x70, "π"},
		{0x72, "ρ"}, {0x56, "ς"}, {0x73, "σ"}, {0x74, "τ"},
		{0x75, "υ"}, {0x6A, "φ"}, {0x63, "χ"}, {0x79, "ψ"},
		{0x77, "ω"}
	};

		private static bool EnableDebugOutput = false;

		private static StringBuilder extractedText = new StringBuilder();

		public static string ExtractText(string filePath, bool isSingleLine = true, bool enableDebug = false)
		{
			EnableDebugOutput = enableDebug;
			extractedText.Clear();

			using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
			{
				var body = doc.MainDocumentPart.Document.Body;
				if (body != null)
				{
					ExtractTextRecursive(body, 0);
				}
			}

			string ret = CleanupText(extractedText.ToString());
			if (isSingleLine)
			{
				ret = Regex.Replace(ret, @"\r\n|\r|\n", " "); //改行は空白に置き換える
			}
			return ret;
		}

		static void ExtractTextRecursive(OpenXmlElement element, int depth)
		{
			if (element is Paragraph paragraph)
			{
				ProcessParagraph(paragraph, depth);
			}
			else if (element is Run run)
			{
				ProcessRun(run, depth);
			}
			else if (element is OpenXmlUnknownElement unknownElement)
			{
				ProcessUnknownElement(unknownElement, depth);
			}
			else
			{
				foreach (var child in element.Elements())
				{
					ExtractTextRecursive(child, depth + 1);
				}
			}
		}

		static void ProcessParagraph(Paragraph paragraph, int depth)
		{
			foreach (var child in paragraph.Elements())
			{
				ExtractTextRecursive(child, depth + 1);
			}
			extractedText.AppendLine();
		}

		static void ProcessRun(Run run, int depth)
		{
			bool isSymbolFont = IsSymbolFont(run);
			bool hasShading = HasShading(run);

			DebugOutput($"Processing Run: IsSymbolFont={isSymbolFont}, HasShading={hasShading}", depth);

			foreach (var runChild in run.ChildElements)
			{
				if (runChild is Text textElement)
				{
					ProcessText(textElement.Text, isSymbolFont, hasShading, depth + 1);
				}
				else if (runChild.LocalName == "sym")
				{
					ProcessSymbol(runChild, depth + 1);
				}
				else if (runChild is OpenXmlUnknownElement unknownElement)
				{
					ProcessUnknownElement(unknownElement, depth + 1);
				}
				else
				{
					DebugOutput($"Unexpected element in Run: {runChild.GetType().Name}", depth + 1);
				}
			}
		}

		static void ProcessText(string textContent, bool isSymbolFont, bool hasShading, int depth)
		{
			DebugOutput($"Processing Text: Length={textContent.Length}, IsSymbolFont={isSymbolFont}, HasShading={hasShading}", depth);
			DebugOutput($"Text content: {textContent}", depth);

			if (hasShading)
			{
				DebugOutput($"Text with shading: {textContent}", depth);
			}

			extractedText.Append(textContent);
		}

		static void ProcessSymbol(OpenXmlElement symbolElement, int depth)
		{
			var charAttribute = symbolElement.GetAttributes().FirstOrDefault(a => a.LocalName == "char");
			if (charAttribute != null)
			{
				string symbolChar = charAttribute.Value;
				string converted = ConvertSymbolChar(symbolChar);
				extractedText.Append(converted);
			}
		}

		// IsSymbolFont, HasShading, ProcessSymbol, ConvertChar, ConvertSymbolChar メソッドは変更なし

		static void ProcessUnknownElement(OpenXmlUnknownElement element, int depth)
		{
			DebugOutput($"Processing Unknown Element: {element.LocalName}", depth);
			if (element.HasAttributes)
			{
				foreach (var attr in element.GetAttributes())
				{
					DebugOutput($"  Attribute: {attr.LocalName} = {attr.Value}", depth);
				}
			}

			if (element.HasChildren)
			{
				foreach (var child in element.ChildElements)
				{
					ExtractTextRecursive(child, depth + 1);
				}
			}
			else if (!string.IsNullOrWhiteSpace(element.InnerText))
			{
				DebugOutput($"  Inner Text: {element.InnerText}", depth);
				extractedText.Append(element.InnerText);
			}
		}

		static bool IsSymbolFont(Run run)
		{
			var rPr = run.RunProperties;
			if (rPr != null && rPr.RunFonts != null)
			{
				return rPr.RunFonts.Ascii?.Value == "Symbol" ||
					   rPr.RunFonts.HighAnsi?.Value == "Symbol" ||
					   rPr.RunFonts.ComplexScript?.Value == "Symbol";
			}
			return false;
		}

		static bool HasShading(Run run)
		{
			var rPr = run.RunProperties;
			if (rPr != null)
			{
				if (rPr.Shading != null)
				{
					DebugOutput($"Shading found: Fill={rPr.Shading.Fill?.Value}, Color={rPr.Shading.Color?.Value}", 0);
					return true;
				}
				if (rPr.Color != null)
				{
					DebugOutput($"Color found: {rPr.Color.Val}", 0);
					return true;
				}
			}
			return false;
		}

		static string ConvertChar(char c, bool isSymbolFont)
		{
			if (isSymbolFont)
			{
				return ConvertSymbolChar(((byte)c).ToString("X2"));
			}
			else if (GreekCharMap.TryGetValue(c, out string greekChar))
			{
				return greekChar;
			}
			return c.ToString();
		}

		static string ConvertSymbolChar(string charValue)
		{
			if (byte.TryParse(charValue, System.Globalization.NumberStyles.HexNumber, null, out byte symbolByte))
			{
				if (SymbolToUnicode.TryGetValue(symbolByte, out string unicodeChar))
				{
					return unicodeChar;
				}
			}
			return charValue;
		}

		static string ExtractFromMathElement(OpenXmlElement mathElement, int depth)
		{
			StringBuilder mathText = new StringBuilder();

			switch (mathElement.LocalName)
			{
				case "oMathPara":
				case "oMath":
					foreach (var child in mathElement.ChildElements)
					{
						mathText.Append(ExtractFromMathElement(child, depth + 1));
					}
					break;
				case "r":
					string innerText = mathElement.InnerText;
					string convertedText = ConvertMathText(innerText);
					mathText.Append(convertedText);
					break;
				default:
					foreach (var child in mathElement.ChildElements)
					{
						mathText.Append(ExtractFromMathElement(child, depth + 1));
					}
					break;
			}

			return mathText.ToString();
		}

		static string ConvertMathText(string text)
		{
			StringBuilder converted = new StringBuilder();
			for (int i = 0; i < text.Length; i++)
			{
				int charCode = char.ConvertToUtf32(text, i);

				if (char.IsSurrogate(text[i]))
				{
					i++;
				}

				if (GreekCharMap.TryGetValue(charCode, out string specialChar))
				{
					converted.Append(specialChar);
				}
				else
				{
					converted.Append(char.ConvertFromUtf32(charCode));
				}
			}
			return converted.ToString();
		}

		static string CleanupText(string text)
		{
			// 余分な空白の削除（ただし、ハイフンの前後の空白は保持）
			text = Regex.Replace(text, @"(?<!\s-)[^\S\n\r]+(?!-\s)", " ");
			// 行頭と行末の空白を削除
			text = Regex.Replace(text, @"^\s+|\s+$", "", RegexOptions.Multiline);
			// 連続する改行を1つにまとめる
			text = Regex.Replace(text, @"\n+", "\n");
			// 段落番号の後に余分な数字がある場合、それを削除（ただし先頭の0は保持）
			text = Regex.Replace(text, @"(\[0*\d+\])(?:\d+)", "$1");
			return text.Trim();
		}

		static void DebugOutput(string message, int depth)
		{
			if (EnableDebugOutput)
			{
				Console.WriteLine($"{new string(' ', depth * 2)}{message}");
			}
		}
	}

}
