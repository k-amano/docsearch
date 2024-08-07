using System.Text;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

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

		public static string ExtractText(string filePath)
		{
			StringBuilder text = new StringBuilder();
			using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
			{
				var body = doc.MainDocumentPart.Document.Body;
				if (body != null)
				{
					ExtractTextAndEquations(body, text, 0);
				}
			}
			return text.ToString();
		}

		static void ExtractTextAndEquations(OpenXmlElement element, StringBuilder text, int depth)
		{
			foreach (var child in element.ChildElements)
			{
				if (child is Run run)
				{
					ProcessRun(run, text, depth + 1);
				}
				else if (child is Paragraph para)
				{
					ExtractTextAndEquations(para, text, depth + 1);
					text.AppendLine();
				}
				else if (child.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/math")
				{
					text.Append(ExtractFromMathElement(child, depth + 1));
				}
				else
				{
					ExtractTextAndEquations(child, text, depth + 1);
				}
			}
		}

		static void ProcessRun(Run run, StringBuilder text, int depth)
		{
			bool isSymbolFont = IsSymbolFont(run);
			foreach (var runChild in run.ChildElements)
			{
				if (runChild is Text textElement)
				{
					ProcessText(textElement.Text, isSymbolFont, text, depth + 1);
				}
				else if (runChild.LocalName == "sym")
				{
					ProcessSymbol(runChild, text, depth + 1);
				}
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

		static string GetRunProperties(Run run)
		{
			var rPr = run.RunProperties;
			if (rPr != null)
			{
				var props = new List<string>();
				if (rPr.RunFonts != null)
				{
					props.Add($"Ascii: {rPr.RunFonts.Ascii?.Value}");
					props.Add($"HighAnsi: {rPr.RunFonts.HighAnsi?.Value}");
					props.Add($"ComplexScript: {rPr.RunFonts.ComplexScript?.Value}");
				}
				if (rPr.FontSize != null) props.Add($"FontSize: {rPr.FontSize.Val}");
				if (rPr.Bold != null) props.Add("Bold");
				if (rPr.Italic != null) props.Add("Italic");
				return string.Join(", ", props);
			}
			return "No properties";
		}

		static void ProcessText(string textContent, bool isSymbolFont, StringBuilder text, int depth)
		{
			foreach (char c in textContent)
			{
				string converted = ConvertChar(c, isSymbolFont);
				text.Append(converted);
			}
		}

		static void ProcessSymbol(OpenXmlElement symbolElement, StringBuilder text, int depth)
		{
			var charAttribute = symbolElement.GetAttributes().FirstOrDefault(a => a.LocalName == "char");
			if (charAttribute != null)
			{
				string symbolChar = charAttribute.Value;
				string converted = ConvertSymbolChar(symbolChar);
				text.Append(converted);
			}
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
				// ... (その他のケースは変更なし)
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
					i++; // サロゲートペアの場合、次の文字をスキップ
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

		static string GetUnicodeCharacterName(int codePoint)
		{
			// この関数は簡易的なものです。実際の Unicode 文字名データベースを使用するとよりよいでしょう。
			if (GreekCharMap.ContainsKey(codePoint))
			{
				return $"GREEK CHARACTER ({GreekCharMap[codePoint]})";
			}
			switch (codePoint)
			{
				case 0x2206: return "INCREMENT";
				case 0x2207: return "NABLA";
				case 0x2200: return "FOR ALL";
				case 0x2203: return "THERE EXISTS";
				case 0x2205: return "EMPTY SET";
				case 0x2208: return "ELEMENT OF";
				case 0x2209: return "NOT AN ELEMENT OF";
				case 0x220B: return "CONTAINS AS MEMBER";
				case 0x220F: return "N-ARY PRODUCT";
				case 0x2211: return "N-ARY SUMMATION";
				case 0x221A: return "SQUARE ROOT";
				case 0x221D: return "PROPORTIONAL TO";
				case 0x221E: return "INFINITY";
				case 0x2229: return "INTERSECTION";
				case 0x222A: return "UNION";
				case 0x2248: return "ALMOST EQUAL TO";
				case 0x2260: return "NOT EQUAL TO";
				case 0x2264: return "LESS-THAN OR EQUAL TO";
				case 0x2265: return "GREATER-THAN OR EQUAL TO";
				default: return "UNKNOWN";
			}
		}

		static string EscapeNonPrintable(string text)
		{
			StringBuilder sb = new StringBuilder();
			foreach (char c in text)
			{
				if (char.IsControl(c) || char.IsWhiteSpace(c) || c > 127)
				{
					sb.Append($"\\u{(int)c:X4}");
				}
				else
				{
					sb.Append(c);
				}
			}
			return sb.ToString();
		}
	}

}
