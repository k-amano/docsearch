using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Arx.DocSearch.Util
{
	public static class SpecialCharConverter
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
		static readonly Dictionary<int, string> SymbolUnicodeMap = new Dictionary<int, string>
	{
		{ 0xF020, "\u0020" }, // SPACE
        { 0xF021, "\u0021" }, // EXCLAMATION MARK
        { 0xF022, "\u2200" }, // FOR ALL
        { 0xF023, "\u0023" }, // NUMBER SIGN
        { 0xF024, "\u2203" }, // THERE EXISTS
        { 0xF025, "\u0025" }, // PERCENT SIGN
        { 0xF026, "\u0026" }, // AMPERSAND
        { 0xF027, "\u220B" }, // CONTAINS AS MEMBER
        { 0xF028, "\u0028" }, // LEFT PARENTHESIS
        { 0xF029, "\u0029" }, // RIGHT PARENTHESIS
        { 0xF02A, "\u2217" }, // ASTERISK OPERATOR
        { 0xF02B, "\u002B" }, // PLUS SIGN
        { 0xF02C, "\u002C" }, // COMMA
        { 0xF02D, "\u2212" }, // MINUS SIGN
        { 0xF02E, "\u002E" }, // FULL STOP
        { 0xF02F, "\u002F" }, // SOLIDUS
        // 数字 (0-9)
        { 0xF030, "\u0030" }, { 0xF031, "\u0031" }, { 0xF032, "\u0032" },
		{ 0xF033, "\u0033" }, { 0xF034, "\u0034" }, { 0xF035, "\u0035" },
		{ 0xF036, "\u0036" }, { 0xF037, "\u0037" }, { 0xF038, "\u0038" },
		{ 0xF039, "\u0039" },
        // その他の記号
        { 0xF03A, "\u003A" }, // COLON
        { 0xF03B, "\u003B" }, // SEMICOLON
        { 0xF03C, "\u003C" }, // LESS-THAN SIGN
        { 0xF03D, "\u003D" }, // EQUALS SIGN
        { 0xF03E, "\u003E" }, // GREATER-THAN SIGN
        { 0xF03F, "\u003F" }, // QUESTION MARK
        { 0xF040, "\u2245" }, // APPROXIMATELY EQUAL TO
        // ギリシャ文字 (大文字)
        { 0xF041, "\u0391" }, { 0xF042, "\u0392" }, { 0xF043, "\u03A7" },
		{ 0xF044, "\u0394" }, { 0xF045, "\u0395" }, { 0xF046, "\u03A6" },
		{ 0xF047, "\u0393" }, { 0xF048, "\u0397" }, { 0xF049, "\u0399" },
		{ 0xF04A, "\u03D1" }, { 0xF04B, "\u039A" }, { 0xF04C, "\u039B" },
		{ 0xF04D, "\u039C" }, { 0xF04E, "\u039D" }, { 0xF04F, "\u039F" },
		{ 0xF050, "\u03A0" }, { 0xF051, "\u0398" }, { 0xF052, "\u03A1" },
		{ 0xF053, "\u03A3" }, { 0xF054, "\u03A4" }, { 0xF055, "\u03A5" },
		{ 0xF056, "\u03C2" }, { 0xF057, "\u03A9" }, { 0xF058, "\u039E" },
		{ 0xF059, "\u03A8" }, { 0xF05A, "\u0396" },
        // その他の記号
        { 0xF05B, "\u005B" }, // LEFT SQUARE BRACKET
        { 0xF05C, "\u2234" }, // THEREFORE
        { 0xF05D, "\u005D" }, // RIGHT SQUARE BRACKET
        { 0xF05E, "\u22A5" }, // UP TACK
        { 0xF05F, "\u005F" }, // LOW LINE
        { 0xF060, "\uF8E5" }, // RADICAL EXTENDER
        // ギリシャ文字 (小文字)
        { 0xF061, "\u03B1" }, { 0xF062, "\u03B2" }, { 0xF063, "\u03C7" },
		{ 0xF064, "\u03B4" }, { 0xF065, "\u03B5" }, { 0xF066, "\u03C6" },
		{ 0xF067, "\u03B3" }, { 0xF068, "\u03B7" }, { 0xF069, "\u03B9" },
		{ 0xF06A, "\u03D5" }, { 0xF06B, "\u03BA" }, { 0xF06C, "\u03BB" },
		{ 0xF06D, "\u03BC" }, { 0xF06E, "\u03BD" }, { 0xF06F, "\u03BF" },
		{ 0xF070, "\u03C0" }, { 0xF071, "\u03B8" }, { 0xF072, "\u03C1" },
		{ 0xF073, "\u03C3" }, { 0xF074, "\u03C4" }, { 0xF075, "\u03C5" },
		{ 0xF076, "\u03D6" }, { 0xF077, "\u03C9" }, { 0xF078, "\u03BE" },
		{ 0xF079, "\u03C8" }, { 0xF07A, "\u03B6" },
        // その他の記号
        { 0xF07B, "\u007B" }, // LEFT CURLY BRACKET
        { 0xF07C, "\u007C" }, // VERTICAL LINE
        { 0xF07D, "\u007D" }, // RIGHT CURLY BRACKET
        { 0xF07E, "\u223C" }, // TILDE OPERATOR
    };

		public static string ConvertSpecialCharactersInParagraph(Paragraph paragraph)
		{
			StringBuilder extractedText = new StringBuilder();
			ExtractTextRecursive(paragraph, 0, extractedText);
			return extractedText.ToString();
		}

		private static void ExtractTextRecursive(OpenXmlElement element, int depth, StringBuilder extractedText)
		{
			if (element is Run run)
			{
				ProcessRun(run, depth, extractedText);
			}
			else if (element is OpenXmlUnknownElement unknownElement)
			{
				ProcessUnknownElement(unknownElement, depth, extractedText);
			}
			else if (element.LocalName == "oMath" || element.LocalName == "oMathPara")
			{
				extractedText.Append(ExtractFromMathElement(element, depth));
			}
			else
			{
				foreach (var child in element.Elements())
				{
					ExtractTextRecursive(child, depth + 1, extractedText);
				}
			}
		}

		private static void ProcessRun(Run run, int depth, StringBuilder extractedText)
		{
			string convertedText = ConvertSpecialCharactersInRun(run);
			extractedText.Append(convertedText);
			/*
			bool isSymbolFont = IsSymbolFont(run);

			foreach (var runChild in run.ChildElements)
			{
				if (runChild is Text textElement)
				{
					ProcessText(textElement.Text, isSymbolFont, depth + 1, extractedText);
				}
				else if (runChild.LocalName == "sym")
				{
					ProcessSymbol(runChild, depth + 1, extractedText);
				}
				else if (runChild is OpenXmlUnknownElement unknownElement)
				{
					ProcessUnknownElement(unknownElement, depth + 1, extractedText);
				}
			}*/
		}

		private static void ProcessText(string textContent, bool isSymbolFont, int depth, StringBuilder extractedText)
		{
			for (int i = 0; i < textContent.Length; i++)
			{
				char c = textContent[i];
				string converted = ConvertChar(c, isSymbolFont);
				extractedText.Append(converted);
			}
		}

		private static void ProcessSymbol(OpenXmlElement symbolElement, int depth, StringBuilder extractedText)
		{
			var charAttribute = symbolElement.GetAttributes().FirstOrDefault(a => a.LocalName == "char");
			if (charAttribute != null)
			{
				string symbolChar = charAttribute.Value;
				string converted = ConvertSymbolChar(symbolChar);
				extractedText.Append(converted);
			}
		}

		private static void ProcessUnknownElement(OpenXmlUnknownElement element, int depth, StringBuilder extractedText)
		{
			if (element.HasChildren)
			{
				foreach (var child in element.ChildElements)
				{
					ExtractTextRecursive(child, depth + 1, extractedText);
				}
			}
			else if (!string.IsNullOrWhiteSpace(element.InnerText))
			{
				extractedText.Append(element.InnerText);
			}
		}

		public static string ConvertChar(char c, bool isSymbolFont)
		{
			if (char.IsControl(c) || char.IsWhiteSpace(c))
			{
				return c.ToString();
			}

			if (isSymbolFont)
			{
				if (SymbolToUnicode.TryGetValue((byte)c, out string unicodeChar))
				{
					return unicodeChar;
				}
			}

			if (GreekCharMap.TryGetValue(c, out string greekChar))
			{
				return greekChar;
			}

			// 変換できない外字の場合、空白を返す
			return ((int)c >= 0xF000 && (int)c <= 0xF0FF) ? " " : c.ToString();
		}

		public static string ConvertSymbolChar(string charValue)
		{
			if (int.TryParse(charValue, System.Globalization.NumberStyles.HexNumber, null, out int symbolCode))
			{
				if (SymbolUnicodeMap.TryGetValue(symbolCode, out string unicodeChar))
				{
					return unicodeChar;
				}
			}
			// 変換できない外字の場合、空白を返す
			if (symbolCode >= 0xF000 && symbolCode <= 0xF0FF)
			{
				return " ";
			}
			return charValue;
		}

		public static bool IsSymbolFont(Run run)
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

		public static string ExtractFromMathElement(OpenXmlElement mathElement, int depth)
		{
			StringBuilder mathText = new StringBuilder();

			switch (mathElement.LocalName)
			{
				case "oMathPara":
				case "oMath":
				case "sSup": // 上付き文字
				case "sSubSup": // 下付きおよび上付き文字
				case "sSub": // 下付き文字
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
				case "sup": // 上付き文字の内容
					mathText.Append("^(");
					foreach (var child in mathElement.ChildElements)
					{
						mathText.Append(ExtractFromMathElement(child, depth + 1));
					}
					mathText.Append(")");
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

		public static string ConvertMathText(string text)
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

		public static string ConvertSpecialCharactersInRun(Run run)
		{
			StringBuilder convertedText = new StringBuilder();
			bool isSymbolFont = IsSymbolFont(run);

			foreach (var runChild in run.ChildElements)
			{
				if (runChild is Text textElement)
				{
					foreach (char c in textElement.Text)
					{
						string converted = ConvertChar(c, isSymbolFont);
						convertedText.Append(converted);
					}
				}
				else if (runChild.LocalName == "sym")
				{
					string converted = ConvertSymbolElement(runChild);
					convertedText.Append(converted);
				}
			}

			return convertedText.ToString();
		}

		private static string ConvertSymbolElement(OpenXmlElement symbolElement)
		{
			var charAttribute = symbolElement.GetAttributes().FirstOrDefault(a => a.LocalName == "char");
			if (charAttribute != null)
			{
				string symbolChar = charAttribute.Value;
				return ConvertSymbolChar(symbolChar);
			}
			return " "; // 変換できない場合は空白を返す
		}

		public static string ReplaceLine(string line)
		{
			line = Regex.Replace(line ?? "", @"[\u00a0\uc2a0\u200e]", " "); //文字コードC2A0（UTF-8の半角空白）
			line = Regex.Replace(line ?? "", @"[\u0091\u0092\u2018\u2019]", "'"); //UTF-8のシングルクォーテーション
			line = Regex.Replace(line ?? "", @"[\u0093\u0094\u00AB\u201C\u201D]", "\""); //UTF-8のダブルクォーテーション
			line = Regex.Replace(line ?? "", @"[\u0097\u2013\u2014]", "\""); //UTF-8のハイフン
			line = Regex.Replace(line ?? "", @"[\u00A9\u00AE\u2022\u2122]", "\""); //UTF-8のスラッシュ}
			line = TextConverter.ZenToHan(line ?? "");
			line = TextConverter.HankToZen(line ?? "");
			return line;
		}

		public static string ReplaceMathSymbols(string input)
		{
			// ギリシャ文字と数学記号のUnicode範囲
			string pattern = @"[\u0370-\u03FF\u1F00-\u1FFF" +  // ギリシャ文字
							  @"\u2100-\u214F" +               // 文字様記号
							  @"\u2190-\u21FF" +               // 矢印
							  @"\u2200-\u22FF" +               // 数学記号
							  @"\u2300-\u23FF" +               // その他の技術記号
							  @"\u25A0-\u25FF" +               // 幾何学模様
							  @"\u2600-\u26FF" +               // その他の記号
							  @"\u2700-\u27BF" +               // 装飾記号
							  @"\u27C0-\u27EF" +               // その他の数学記号-A
							  @"\u2980-\u29FF" +               // その他の数学記号-B
							  @"\u2A00-\u2AFF" +               // 補助数学演算子
							  @"\u2B00-\u2BFF]";               // その他の記号と矢印

			return Regex.Replace(input, pattern, "");
		}

		static public string RemoveSymbols(string input)
		{
			// 半角記号を削除するための正規表現パターン
			string pattern = @"[!""#$%&'()*+,-./:;<=>?@[\\\]^_`{|}~]";

			// 正規表現を使用して半角記号を空文字に置換
			string result = Regex.Replace(input, pattern, "");

			return result;
		}

	}
}
