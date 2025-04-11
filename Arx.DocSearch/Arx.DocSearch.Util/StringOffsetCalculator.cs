using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            // secondStringがfirstStringの後ろにある場合を探す
            for (int i = 0; i < secondString.Length; i++)
            {
                string secondPrefix = secondString.Substring(0, secondString.Length - i);
                if (secondString.Length - i < matchingLength) break;
                int? index = IndexOfIgnoreWhitespace(firstString, secondPrefix);
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
                int? index = IndexOfIgnoreWhitespace(secondString, firstPrefix);
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

        private static int? IndexOfIgnoreWhitespace(string source, string target)
        {
            if (string.IsNullOrEmpty(source) && string.IsNullOrEmpty(target))
            {
                return 0; // 両方とも空の場合は一致しているので0を返す
            }
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target))
            {
                return null; // どちらか一方が空の場合は一致しないのでnullを返す
            }
            string sourceNoSymbol = SpecialCharConverter.RemoveSymbolsAll(source);
            string targetNoSymbol = SpecialCharConverter.RemoveSymbolsAll(target);
            if (string.IsNullOrEmpty(sourceNoSymbol) || string.IsNullOrEmpty(targetNoSymbol))
            {
                return null; // 変換後の文字列が空になった場合は一致しないのでnullを返す
            }

            string sourceNoWhitespace = new string(sourceNoSymbol.Where(c => !char.IsWhiteSpace(c)).ToArray());
            string targetNoWhitespace = new string(targetNoSymbol.Where(c => !char.IsWhiteSpace(c)).ToArray());

            int index = sourceNoWhitespace.IndexOf(targetNoWhitespace);
            if (index == -1) return null; // 一致が見つからない場合はnullを返す

            int originalIndex = 0;
            int noSymbolIndex = 0;

            while (noSymbolIndex < index && originalIndex < source.Length)
            {
                if (!char.IsWhiteSpace(source[originalIndex]) &&
                    sourceNoSymbol.Contains(source[originalIndex]))
                {
                    noSymbolIndex++;
                }
                originalIndex++;
            }

            // 元の文字列で空白や無視される文字をスキップ
            while (originalIndex < source.Length &&
                   (char.IsWhiteSpace(source[originalIndex]) ||
                    !sourceNoSymbol.Contains(source[originalIndex])))
            {
                originalIndex++;
            }

            // ターゲット文字列の先頭にある空白や無視される文字の数を計算
            int targetSkipCount = 0;
            while (targetSkipCount < target.Length &&
                   (char.IsWhiteSpace(target[targetSkipCount]) ||
                    !targetNoSymbol.Contains(target[targetSkipCount])))
            {
                targetSkipCount++;
            }

            // ソース文字列のオフセットからターゲット文字列のスキップ数を引く
            return originalIndex - targetSkipCount;
        }
    }
}
