using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Arx.DocSearch.Util
{
	public class StringOffsetCalculator
	{
		public static int? CalculateOffset(string firstString, string secondString)
		{
			int matchingLength = GetMatchingLength(firstString, secondString);

			//Console.WriteLine($"最長共通部分の長さ: {matchingLength}");
			//Console.WriteLine($"firstString: {firstString}");
			//Console.WriteLine($"secondString: {secondString}");

			// secondStringがfirstStringの後ろにある場合を探す
			for (int i = 0; i < secondString.Length; i++)
			{
				string secondPrefix = secondString.Substring(0, secondString.Length - i);
				if (secondString.Length - i < matchingLength) break;
				int index = IndexOfIgnoreWhitespace(firstString, secondPrefix);
				if (0 <= index)
				{
					return index;  // firstStringの開始位置がsecondStringより前
				}
			}

			// secondStringがfirstStringの前にある場合を探す
			for (int i = 0; i < firstString.Length; i++)
			{
				string firstPrefix = firstString.Substring(0, firstString.Length - i);
				if (firstString.Length - i < matchingLength) break;
				int index = IndexOfIgnoreWhitespace(secondString, firstPrefix);
				if (0 <= index)
				{
					return -index;  // secondStringの開始位置がfirstStringより前
				}
			}

			// 一致する部分が見つからない場合
			return null;
		}

		public static int GetMatchingLength(string str1, string str2)
		{
			int maxLength = 0;
			int currentLength = 0;

			string str1NoWhitespace = new string(str1.Where(c => !char.IsWhiteSpace(c)).ToArray());
			string str2NoWhitespace = new string(str2.Where(c => !char.IsWhiteSpace(c)).ToArray());

			for (int i = 0; i < str1NoWhitespace.Length; i++)
			{
				for (int j = 0; j < str2NoWhitespace.Length; j++)
				{
					currentLength = 0;
					while (i + currentLength < str1NoWhitespace.Length &&
						j + currentLength < str2NoWhitespace.Length &&
						str1NoWhitespace[i + currentLength] == str2NoWhitespace[j + currentLength])
					{
						currentLength++;
					}
					if (currentLength > maxLength)
					{
						maxLength = currentLength;
					}
				}
			}
			return maxLength;
		}

		private static int IndexOfIgnoreWhitespace(string source, string target)
		{
			string sourceNoSymbol = SpecialCharConverter.ReplaceMathSymbols(source);
			string targetNoSymbol = SpecialCharConverter.ReplaceMathSymbols(target);
			string sourceNoWhitespace = new string(sourceNoSymbol.Where(c => !char.IsWhiteSpace(c)).ToArray());
			string targetNoWhitespace = new string(targetNoSymbol.Where(c => !char.IsWhiteSpace(c)).ToArray());

			int index = sourceNoWhitespace.IndexOf(targetNoWhitespace);
			if (index == -1) return -1;

			int originalIndex = 0;
			int noWhitespaceIndex = 0;

			while (noWhitespaceIndex <= index) // '<' を '<=' に変更
			{
				if (!char.IsWhiteSpace(source[originalIndex]))
				{
					noWhitespaceIndex++;
				}
				originalIndex++;
			}

			return originalIndex - 1; // '-1' を追加
		}
	}
}
