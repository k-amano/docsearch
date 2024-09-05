using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.International.Converters;

namespace Arx.DocSearch.Util
{
	public class TextConverter
	{
		/// <summary>
		/// 全角文字を半角に変換します。
		/// </summary>
		/// <param name="source">元の文字列。</param>
		/// <returns>変換された文字列。</returns>
		public static string ZenToHan(string source)
		{
			return Regex.Replace(source, @"[　！-～]", delegate (Match m)
			{
				//全角スペース
				if ("　".Equals(m.Value)) return " ";
				//英数記号
				char ch = (char)('!' + (m.Value[0] - '！'));
				return ch.ToString();
			});
		}

		/// <summary>
		/// 半角かな文字を全角に変換します。
		/// </summary>
		/// <param name="source">元の文字列。</param>
		/// <returns>変換された文字列。</returns>
		public static string HankToZen(string source)
		{
			return Regex.Replace(source, @"[｡-ﾟ]", delegate (Match m)
			{
				return KanaConverter.HalfwidthKatakanaToKatakana(m.Value);
			});
		}

		/// <summary>
		/// 指定した Unicode 文字が、ひらがなかどうかを示します。
		/// </summary>
		/// <param name="c">評価する Unicode 文字。</param>
		/// <returns>c がひらがなである場合は true。それ以外の場合は false。</returns>
		public static bool IsHiragana(char c)
		{
			//「ぁ」～「ゟ(より)」までをひらがなとする
			return '\u3041' <= c && c <= '\u309F';
		}

		/// <summary>
		/// 指定した Unicode 文字が、ひらがな・カタカナ共通文字かどうかを示します。
		/// </summary>
		/// <param name="c">評価する Unicode 文字。</param>
		/// <returns>c がひらがなである場合は true。それ以外の場合は false。</returns>
		public static bool IsKanaCommon(char c)
		{
			//「゠(\u30A0)」「・(\u30FB)」「ー(\u30FC)」をひらがな・カタカナ共通文字とする
			return c == '\u30A0' || c == '\u30FB' || c == '\u30FC';
		}

		/// <summary>
		/// 指定した Unicode 文字が、全角カタカナかどうかを示します。
		/// </summary>
		/// <param name="c">評価する Unicode 文字。</param>
		/// <returns>c が全角カタカナである場合は true。それ以外の場合は false。</returns>
		public static bool IsFullwidthKatakana(char c)
		{
			//「゠(\u30A0)」から「ヿ(\u30FF)コト」までと、カタカナフリガナ拡張「ㇰ」(\u31F0)～「ㇿ」(\u31FF)と、
			//濁点と半濁点「゙」(\u3099)～「゜」(\u309C)を全角カタカナとする
			return ('\u30A0' <= c && c <= '\u30FF')
					|| ('\u31F0' <= c && c <= '\u31FF')
					|| ('\u3099' <= c && c <= '\u309C');
		}

		/// <summary>
		/// 指定した Unicode 文字が、全角句読点・括弧記号かどうかを示します。
		/// </summary>
		/// <param name="c">評価する Unicode 文字。</param>
		/// <returns>c が全角句読点・括弧記号である場合は true。それ以外の場合は false。</returns>
		public static bool IsFullwidthPunctuation(char c)
		{
			//「」(\u3000)～「〿」(\u303F)を全角句読点・括弧記号とする
			return '\u3000' <= c && c <= '\u303F';
		}

		/// <summary>
		/// 指定した Unicode 文字が ASCII 文字(0x00-0x7f)かどうかを示します。
		/// </summary>
		/// <param name="c">評価する Unicode 文字。</param>
		/// <returns>ASCII 文字であれば true、その他が含まれていれば false を返します。</returns>
		public static bool IsAscii(char c)
		{
			return (c <= (char)0x7f);
		}

		/// <summary>
		/// 指定した Unicode 文字が半角英数かどうかを示します。
		/// </summary>
		/// <param name="c">評価する Unicode 文字。</param>
		/// <returns>ASCII 文字であれば true、その他が含まれていれば false を返します。</returns>
		public static bool IsAlphaNumeric(char c)
		{
			//半角英数と、
			//'@_+-'中点も含む
			return ('0' <= c && c <= '9')
					|| ('A' <= c && c <= 'Z')
					|| ('a' <= c && c <= 'z')
					|| ('@' == c) || ('_' == c) || ('+' == c) || ('-' == c);
		}

		/// <summary>
		/// 指定した Unicode 文字が、漢字かどうかを示します。
		/// </summary>
		/// <param name="c">評価する Unicode 文字。</param>
		/// <returns>c が漢字である場合は true。それ以外の場合は false。</returns>
		public static bool IsKanji(char c)
		{
			//CJK統合漢字、CJK互換漢字、CJK統合漢字拡張Aの範囲にあるか調べる
			return ('\u4E00' <= c && c <= '\u9FCF')
					|| ('\uF900' <= c && c <= '\uFAFF')
					|| ('\u3400' <= c && c <= '\u4DBF');
		}

		/// <summary>
		/// 文字種類が前後で変更されているかどうかを判定します。
		/// </summary>
		/// <param name="c1">最初の文字。</param>
		/// <param name="c2">次の文字。</param>
		/// <returns>最初の文字と次の文字で文字種類が異なれば true、同じであれば false を返します。</returns>
		public static bool IsCharTypeChanged(char c1, char c2)
		{
			if (IsAscii(c1) && IsAscii(c2))
			{
				if (IsAlphaNumeric(c1) && IsAlphaNumeric(c2)) return false;
				else return true;
			}
			else if (IsAscii(c1) || IsAscii(c2)) return true;
			else if (IsFullwidthPunctuation(c1)) return true;
			else if (IsHiragana(c1) && IsHiragana(c2)) return false;
			else if (IsHiragana(c1) && !IsKanaCommon(c2)) return true;
			else if (!IsKanaCommon(c1) && IsHiragana(c2)) return true;
			else if (IsFullwidthKatakana(c1) && IsFullwidthKatakana(c2)) return false;
			else if (IsFullwidthKatakana(c1) || IsFullwidthKatakana(c2)) return true;
			return false;
		}

		/// <summary>
		/// 文字列を文字種類の変更された箇所で分かち書きします。
		/// </summary>
		/// <param name="src">元の文字列。</param>
		/// <returns>分かち書きされた文字列。</returns>
		public static string SplitWords(string src)
		{
			List<char> ls = new List<char>();
			for (int i = 0; i < src.Length; i++)
			{
				char c1 = src[i];
				if ((!IsAscii(c1) || IsAlphaNumeric(c1)) && !IsFullwidthPunctuation(c1)) ls.Add(c1);
				if (i == src.Length - 1) break;
				char c2 = src[i + 1];
				char lastChar = ' ';
				if (0 < ls.Count) lastChar = ls[ls.Count - 1];
				if (IsCharTypeChanged(c1, c2) && lastChar != ' ') ls.Add(' ');
			}
			return new String(ls.ToArray());
		}

		public static bool IsJp(string src)
		{
			foreach (char c in src)
			{
				if (IsHiragana(c)) return true;
				if (IsFullwidthKatakana(c)) return true;
				if (IsKanji(c)) return true;
			}
			return false;
		}
	}
}
