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
        private int currentTestCaseIndex = -1; // デバッグ用：現在のテストケースインデックス

        public void SetTestCaseIndex(int index)
        {
            currentTestCaseIndex = index;
        }

        public string HighlightTextInWord(string filePath, int[] indexes, double[] rates, string[] searchPatterns, bool isDebug = false)
        {
            StringBuilder sb = new StringBuilder();
            CleanDocument(filePath, filePath);
            WordTextExtractor wte = new WordTextExtractor(filePath, true, false);
            //Arx.DocSearch.Util.WordTextExtractor wte = new Arx.DocSearch.Util.WordTextExtractor(filePath, true, false);
            string pattern;

            //Console.WriteLine("=== HighlightTextInWord start ===");
            //Console.WriteLine($"File: {filePath}");
            //Console.WriteLine($"Search patterns count: {searchPatterns.Length}");

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    Body body = doc.MainDocumentPart.Document.Body;
                    if (body != null)
                    {
                        var paragraphs = body.Descendants<Paragraph>().ToList();
                        //Console.WriteLine($"Found {paragraphs.Count} paragraphs");

                        if (isDebug) sb.AppendLine($"filePath:\n{filePath}\ndocText:\n{wte.Text}");

                        // XMLから直接パラグラフ位置を計算
                        Dictionary<Paragraph, int> paragraphStartPositions = BuildParagraphPositionMap(body);


                        // 全パラグラフのテキストを連結
                        StringBuilder combinedDocText = new StringBuilder();
                        List<(int start, int end, Paragraph paragraph)> paragraphRanges = new List<(int start, int end, Paragraph paragraph)>();

                        // paragraphRangesを新しい位置マッピングから構築
                        foreach (var paragraph in paragraphs)
                        {
                            if (paragraphStartPositions.ContainsKey(paragraph))
                            {
                                int startPos = paragraphStartPositions[paragraph];
                                int length = CalculateParagraphTextLength(paragraph);

                                if (length > 0)
                                {
                                    paragraphRanges.Add((startPos, startPos + length - 1, paragraph));
                                }
                            }

                            // combinedDocTextは検索用に引き続き必要
                            var paragraphText = CreateCombinedText(paragraph);
                            if (!string.IsNullOrWhiteSpace(paragraphText))
                            {
                                combinedDocText.Append(paragraphText);
                            }
                        }

                        string fullDocText = combinedDocText.ToString();
                        //Console.WriteLine("fullDocText:\n" + fullDocText);

                        for (int i = 0; i < searchPatterns.Length && i < rates.Length; i++)
                        {
                            pattern = PrepareSearchPattern(searchPatterns[i], out string[] words);
                            if (searchPatterns[i].Length < 20 && words.Length < 3) continue;

                            bool foundInDocument = false;
                            var results = MatchIgnoringWhitespace(pattern, fullDocText, sb);

                            if (results.Count > 0)
                            {
                                foundInDocument = true;
                                LogFoundPattern(isDebug, sb, i, searchPatterns, pattern);

                                foreach (var result in results)
                                {
                                    // 位置マッピングを使用してマッチ位置から該当するパラグラフを特定
                                    Paragraph targetParagraph = FindParagraphByPosition(result.beginIndex, paragraphStartPositions);

                                    if (targetParagraph != null)
                                    {
                                        // パラグラフの開始位置を取得
                                        int paragraphStartPos = paragraphStartPositions[targetParagraph];


                                        // 既存の処理も並行して実行（比較のため）
                                        var affectedParagraphs = FindAffectedParagraphs(paragraphRanges, result);
                                        // Process affected paragraphs
                                        ProcessAffectedParagraphsAll(affectedParagraphs, result, rates[i], fullDocText, sb, isDebug);
                                    }
                                }
                            }

                            if (!foundInDocument)
                            {
                                LogPatternNotFound(sb, i, searchPatterns, pattern);
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

        /*
        /// <summary>
        /// エラー情報をログに記録します
        /// </summary>
        private void LogError(StringBuilder sb, Exception ex)
        {
            sb.AppendLine($"エラーが発生しました: {ex.Message}");
            sb.AppendLine($"スタックトレース: {ex.StackTrace}");
        }
        */

        /// <summary>
        /// 検索パターンを準備します
        /// </summary>
        private string PrepareSearchPattern(string originalPattern, out string[] words)
        {
            string searchPattern = Regex.Replace(originalPattern, @"^[0-9]+\.?\s+", "");
            //searchPattern = Regex.Replace(searchPattern, @"\s+[0-9]+\.?\s*$", "");
            words = searchPattern.Split(' ');
            string pattern = CreateSearchPattern(searchPattern);


            return pattern;
        }

        /// <summary>
        /// 見つかった場合のログ出力
        /// </summary>
        private void LogFoundPattern(bool isDebug, StringBuilder sb, int index, string[] searchPatterns, string pattern)
        {
            if (isDebug)
            {
                sb.AppendLine("エラー: 指定されたテキストが見つかりました。");
                sb.AppendLine($"検索文:searchPatterns[{index}] {searchPatterns[index]}\n{pattern}");
            }
        }

        /// <summary>
        /// 見つからなかった場合のログ出力
        /// </summary>
        private void LogPatternNotFound(StringBuilder sb, int index, string[] searchPatterns, string pattern)
        {
            sb.AppendLine("エラー: 指定されたテキストが見つかりませんでした。");
            sb.AppendLine($"検索文: searchPatterns[{index}]{searchPatterns[index]}\n{pattern}");
        }

        /// <summary>
        /// 検索結果に影響を受けるパラグラフを特定します
        /// </summary>
        private List<(int start, int end, Paragraph paragraph)> FindAffectedParagraphs(
            List<(int start, int end, Paragraph paragraph)> paragraphRanges,
            (int beginIndex, int endIndex) result)
        {
            return paragraphRanges
                .Where(p => DoRangesOverlap(result.beginIndex, result.endIndex, p.start, p.end))
                .ToList();
        }

        /// <summary>
        /// パラグラフをテキスト結合用に処理します
        /// </summary>
        private void ProcessParagraphForCombinedText(
            Paragraph paragraph,
            StringBuilder combinedDocText,
            List<(int start, int end, Paragraph paragraph)> paragraphRanges,
            ref int currentPosition)
        {
            var paragraphText = CreateCombinedText(paragraph);
            if (!string.IsNullOrWhiteSpace(paragraphText))
            {
                paragraphRanges.Add((currentPosition, currentPosition + paragraphText.Length - 1, paragraph));
                combinedDocText.Append(paragraphText);
                currentPosition += paragraphText.Length;
            }
        }

        /*
        /// <summary>
        /// 影響を受けるパラグラフを処理します
        /// </summary>
        private void ProcessAffectedParagraph(
            (int start, int end, Paragraph paragraph) paragraphRange,
            (int beginIndex, int endIndex) result,
            double rate,
            string fullDocText,
            StringBuilder sb,
            bool isDebug)
        {
            var paragraph = paragraphRange.paragraph;
            var combinedText = CreateCombinedText(paragraph);
            var elementRanges = GetElementRanges(paragraph);
            string displayText = GetDisplayText(elementRanges);

            // パラグラフ内での相対位置を計算
            int relativeStart = Math.Max(0, result.beginIndex - paragraphRange.start);
            int relativeEnd = Math.Min(combinedText.Length - 1, result.endIndex - paragraphRange.start);
            //Console.WriteLine($"Searching in range: {relativeStart}-{relativeEnd}");

            Console.WriteLine($"[ProcessAffectedParagraph] AdjustRelativePositionsを呼び出します");
            AdjustRelativePositions(combinedText, displayText, ref relativeStart, ref relativeEnd);

            var matchedElements = elementRanges
                .Where(r =>
                {
                    bool overlaps = DoRangesOverlap(relativeStart, relativeEnd, r.start, r.end);
                    //Console.WriteLine($"Search:{relativeStart}-{relativeEnd} Element:{r.start}-{r.end} Type:{r.element.LocalName} Text:'{r.element.InnerText}' Overlaps:{overlaps}");
                    return overlaps;
                })
                .ToList();

            StringBuilder highlightedText = new StringBuilder();

            //Console.WriteLine($"=== About to process {matchedElements.Count} elements ===");
            //Console.WriteLine($"relativeStart: {relativeStart}, relativeEnd: {relativeEnd}");

            foreach (var elem in matchedElements)
            {
                ProcessMatchedElement(elem, relativeStart, relativeEnd, rate, highlightedText);
            }

            VerifyHighlightResult(fullDocText, result, highlightedText.ToString(), sb, isDebug);
        }
        */

        /// <summary>
        /// 相対位置を調整します
        /// </summary>
        /*private void AdjustRelativePositions(
            string combinedText,
            string displayText,
            ref int relativeStart,
            ref int relativeEnd)
        {
            string searchText = SafeSubstring(combinedText, relativeStart, relativeEnd);
            //^(任意の文字列) と _(任意の文字列) パターンを除去する
            //ssup: 上付き"^()", ssub: 下付き"_()" が位置計算の誤差になることを考慮する
            searchText = Regex.Replace(Regex.Replace(searchText, @"\^\(([^()]*)\)", "$1"), @"_\(([^()]*)\)", "$1");
            string highlightText = SafeSubstring(displayText, relativeStart, relativeEnd);
            int? offset = StringOffsetCalculator.CalculateOffset(highlightText, searchText);
            if (offset.HasValue)
            {
                relativeStart += offset.Value;
                relativeEnd += offset.Value;
                Console.WriteLine($"offset:{offset}:{offset.Value}:relativeStart:{relativeStart}\nsearchText:{searchText}\nhighlightText:{highlightText}\ndisplayText:{displayText}");
            }
        }*/

        private void AdjustRelativePositions(
            string combinedText,
            string displayText,
            ref int relativeStart,
            ref int relativeEnd)
        {
            Console.WriteLine($"[AdjustRelativePositions] 呼び出されました！");
            Console.WriteLine($"[AdjustRelativePositions] 元の位置: relativeStart={relativeStart}, relativeEnd={relativeEnd}");

            displayText = SpecialCharConverter.ReplaceLine(displayText);

            // デバッグ情報の詳細出力
            StringBuilder debugInfo = new StringBuilder();
            //Console.WriteLine($"=== AdjustRelativePositions デバッグ情報 ===");
            //Console.WriteLine($"Original positions: relativeStart:{relativeStart}, relativeEnd:{relativeEnd}");

            // combinedTextとdisplayTextの概要を出力
            //Console.WriteLine($"combinedText長さ: {combinedText.Length}, displayText長さ: {displayText.Length}");
            //Console.WriteLine($"combinedText先頭50文字: {SafeSubstring(combinedText, 0, 50)}");
            //Console.WriteLine($"displayText先頭50文字: {SafeSubstring(displayText, 0, 50)}");

            // 検索開始位置前後のテキスト確認
            //Console.WriteLine($"relativeStart前後の文字: {SafeSubstring(combinedText, Math.Max(0, relativeStart - 20), 40)}");

            // 検索開始位置までの余分な文字数をカウント（デバッグ情報付き）
            int extraCharsBefore = CountExtraCharsBeforeWithDebug(combinedText, displayText, relativeStart, debugInfo);

            // 検索範囲内の余分な文字数をカウント（デバッグ情報付き）
            int extraCharsInRange = CountExtraCharsInRangeWithDebug(combinedText, displayText, relativeStart, relativeEnd, debugInfo);

            // 位置を補正
            int originalStart = relativeStart;
            int originalEnd = relativeEnd;
            relativeStart -= extraCharsBefore;
            relativeEnd -= (extraCharsBefore + extraCharsInRange);

            // 範囲が有効になるよう調整
            relativeStart = Math.Max(0, relativeStart);
            relativeEnd = Math.Max(relativeStart, relativeEnd);

            // 補正結果の確認
            //Console.WriteLine($"位置補正: {originalStart}-{originalEnd} -> {relativeStart}-{relativeEnd}");
            //Console.WriteLine($"Extra chars before start: {extraCharsBefore}, within range: {extraCharsInRange}");
            //Console.WriteLine($"補正後の範囲のテキスト: {SafeSubstring(displayText, relativeStart, relativeEnd - relativeStart + 1)}");

            // デバッグ情報をファイルに出力（必要に応じて）
            //string debugLogPath = Path.Combine(Path.GetDirectoryName(Path.GetTempPath()), "position_adjust_debug.log");
            //File.AppendAllText(debugLogPath, debugInfo.ToString());

            // コンソールにも出力
            //Console.WriteLine(debugInfo.ToString());
            //Console.WriteLine($"Original positions: relativeStart:{originalStart}, relativeEnd:{originalEnd}");
            //Console.WriteLine($"Extra chars before start: {extraCharsBefore}, within range: {extraCharsInRange}");
            //Console.WriteLine($"Adjusted positions: relativeStart:{relativeStart}, relativeEnd:{relativeEnd}");

            Console.WriteLine($"[AdjustRelativePositions] 調整後の位置: relativeStart={relativeStart}, relativeEnd={relativeEnd}");
            Console.WriteLine($"[AdjustRelativePositions] 調整量: extraCharsBefore={extraCharsBefore}, extraCharsInRange={extraCharsInRange}");
        }

        // 検索開始位置までの余分な文字数をカウント（デバッグ情報付き）
        private int CountExtraCharsBeforeWithDebug(string combinedText, string displayText, int position, StringBuilder debugInfo)
        {
            List<(int position, char character)> extraCharList = new List<(int, char)>();
            int extraChars = 0;        // 余分な文字数
            int combinedPos = 0;       // combinedTextの位置
            int displayPos = 0;        // displayTextの位置

            //Console.WriteLine($"\n== 検索開始位置({position})までの余分な文字カウント ==");

            // combinedTextの指定位置まで1文字ずつ比較
            while (combinedPos < position && displayPos < displayText.Length)
            {
                // 現在比較している文字をデバッグ出力
                if (combinedPos % 50 == 0)
                {
                    //Console.WriteLine($"位置 {combinedPos}/{position}: combinedText='{SafeSubstring(combinedText, combinedPos, 10)}...' displayText='{SafeSubstring(displayText, displayPos, 10)}...'");
                }

                // 現在の文字が一致する場合
                if (combinedText[combinedPos] == displayText[displayPos])
                {
                    // 両方のインデックスを進める
                    combinedPos++;
                    displayPos++;
                }
                else
                {
                    // 余分な文字の詳細をデバッグ出力
                    //Console.WriteLine($"余分な文字検出: 位置={combinedPos}, 文字='{combinedText[combinedPos]}', コード={((int)combinedText[combinedPos]).ToString("X4")}");

                    // リストに余分な文字を追加
                    extraCharList.Add((combinedPos, combinedText[combinedPos]));

                    // combinedTextの文字がdisplayTextに存在しない → 余分な文字
                    extraChars++;
                    combinedPos++;
                }
            }

            // 検出された余分な文字の一覧をデバッグ出力
            /*Console.WriteLine($"\n検出された余分な文字一覧（計: {extraChars}個）:");
            for (int i = 0; i < extraCharList.Count; i++)
            {
                Console.WriteLine($"{i + 1}. 位置: {extraCharList[i].position}, 文字: '{extraCharList[i].character}', コード: {((int)extraCharList[i].character).ToString("X4")}");
            }

            // combinedTextの位置がpositionに達していない場合（displayTextが先に終わった場合）
            if (combinedPos < position)
            {
                int remainingChars = position - combinedPos;
                Console.WriteLine($"displayTextが終了したため、残り {remainingChars} 文字も余分な文字としてカウント");
                extraChars += remainingChars;
            }

            Console.WriteLine($"検索開始位置までの余分な文字数: {extraChars}");*/
            return extraChars;
        }

        // 検索範囲内の余分な文字数をカウント（デバッグ情報付き）
        private int CountExtraCharsInRangeWithDebug(string combinedText, string displayText, int startPos, int endPos, StringBuilder debugInfo)
        {
            List<(int position, char character)> extraCharList = new List<(int, char)>();

            // startPosに対応するdisplayTextの位置を特定
            int displayStartPos = 0;
            int combinedPos = 0;

            //Console.WriteLine($"\n== startPos({startPos})に対応するdisplayTextの位置を特定 ==");

            // combinedTextのstartPosまでの対応するdisplayTextの位置を特定
            while (combinedPos < startPos && displayStartPos < displayText.Length)
            {
                /*if (combinedPos % 50 == 0)
                {
                    Console.WriteLine($"位置 {combinedPos}/{startPos}: combinedText='{SafeSubstring(combinedText, combinedPos, 10)}...' displayText='{SafeSubstring(displayText, displayStartPos, 10)}...'");
                }*/

                if (combinedText[combinedPos] == displayText[displayStartPos])
                {
                    combinedPos++;
                    displayStartPos++;
                }
                else
                {
                    // combinedTextの文字がdisplayTextにない → 余分な文字
                    combinedPos++;
                }
            }

            //Console.WriteLine($"combinedTextの位置 {startPos} は displayTextの位置 {displayStartPos} に対応");

            // 範囲内の余分な文字をカウント
            int extraChars = 0;
            combinedPos = startPos;
            int displayPos = displayStartPos;

            //Console.WriteLine($"\n== 検索範囲内({startPos}-{endPos})の余分な文字カウント ==");

            while (combinedPos <= endPos && displayPos < displayText.Length)
            {
                /*if ((combinedPos - startPos) % 50 == 0)
                {
                    Console.WriteLine($"位置 {combinedPos}/{endPos}: combinedText='{SafeSubstring(combinedText, combinedPos, 10)}...' displayText='{SafeSubstring(displayText, displayPos, 10)}...'");
                }*/

                if (combinedText[combinedPos] == displayText[displayPos])
                {
                    combinedPos++;
                    displayPos++;
                }
                else
                {
                    // 余分な文字の詳細をデバッグ出力
                    //Console.WriteLine($"余分な文字検出: 位置={combinedPos}, 文字='{combinedText[combinedPos]}', コード={((int)combinedText[combinedPos]).ToString("X4")}");

                    // リストに余分な文字を追加
                    extraCharList.Add((combinedPos, combinedText[combinedPos]));

                    // combinedTextの文字がdisplayTextにない → 余分な文字
                    extraChars++;
                    combinedPos++;
                }
            }

            // 検出された余分な文字の一覧をデバッグ出力
            /*Console.WriteLine($"\n検索範囲内の余分な文字一覧（計: {extraChars}個）:");
            for (int i = 0; i < extraCharList.Count; i++)
            {
                Console.WriteLine($"{i + 1}. 位置: {extraCharList[i].position}, 文字: '{extraCharList[i].character}', コード: {((int)extraCharList[i].character).ToString("X4")}");
            }

            // 範囲の最後までカウントできなかった場合（displayTextが先に終わった場合）
            if (combinedPos <= endPos)
            {
                int remainingChars = endPos - combinedPos + 1;
                Console.WriteLine($"displayTextが終了したため、残り {remainingChars} 文字も余分な文字としてカウント");
                extraChars += remainingChars;
            }

            Console.WriteLine($"検索範囲内の余分な文字数: {extraChars}");*/
            return extraChars;
        }

        // combinedText と displayText の内容を詳細表示するデバッグメソッド
        /*private void DebugTextContents(string combinedText, string displayText)
        {
            Console.WriteLine("===== combinedText と displayText の詳細比較 =====");

            // 両方のテキストの長さを表示
            Console.WriteLine($"combinedText長さ: {combinedText.Length}");
            Console.WriteLine($"displayText長さ: {displayText.Length}");

            // 内容を50文字ずつ表示
            int maxLen = Math.Max(combinedText.Length, displayText.Length);
            for (int i = 0; i < maxLen; i += 50)
            {
                string combinedSegment = SafeSubstring(combinedText, i, 50);
                string displaySegment = SafeSubstring(displayText, i, 50);

                Console.WriteLine($"\n=== 位置 {i} から {i + 49} ===");
                Console.WriteLine($"combinedText: \"{combinedSegment}\"");
                Console.WriteLine($"displayText:  \"{displaySegment}\"");

                // 文字ごとの比較結果を表示
                Console.WriteLine("文字ごとの比較:");
                for (int j = 0; j < 50 && i + j < maxLen; j++)
                {
                    char combinedChar = i + j < combinedText.Length ? combinedText[i + j] : ' ';
                    char displayChar = i + j < displayText.Length ? displayText[i + j] : ' ';
                    bool isSame = combinedChar == displayChar;

                    Console.WriteLine($"位置 {i + j}: combinedText[{(int)combinedChar:X4}]='{combinedChar}' " +
                                     $"displayText[{(int)displayChar:X4}]='{displayChar}' " +
                                     $"一致: {isSame}");
                }
            }
        }*/


        /// <summary>
        /// 一致した要素を処理します
        /// </summary>
        private void ProcessMatchedElement(
            (int start, int end, string displayText, OpenXmlElement element) elem,
            int relativeStart,
            int relativeEnd,
            double rate,
            StringBuilder highlightedText)
        {
            //Console.WriteLine($"\nTrying to color - Type:{elem.element.LocalName} Range:{elem.start}-{elem.end} Text:'{elem.element.InnerText}' Parent:{elem.element.Parent?.LocalName}");

            if (elem.element.LocalName == "t")
            {
                ProcessTextElement(elem, relativeStart, relativeEnd, rate, highlightedText);
            }
            else if (elem.element is DocumentFormat.OpenXml.Math.OfficeMath math)
            {
                //Console.WriteLine($"### Applying Math Color to: {math.InnerText}");
                ApplyBackgroundColor(rate, math, highlightedText);
            }
            else
            {
                //Console.WriteLine($"Other element: Type:{elem.element.LocalName} Range:{elem.start}-{elem.end}");
            }
        }

        /// <summary>
        /// テキスト要素を処理します
        /// </summary>
        private void ProcessTextElement(
            (int start, int end, string displayText, OpenXmlElement element) elem,
            int relativeStart,
            int relativeEnd,
            double rate,
            StringBuilder highlightedText)
        {
            var run = GetParentRun(elem.element);
            if (run == null)
            {
                //Console.WriteLine($"No parent Run found for text: '{elem.element.InnerText}'");
                return;
            }

            //Console.WriteLine($"Processing run: '{run.InnerText}'");
            int runStart = elem.start;
            int runEnd = elem.end;

            //Console.WriteLine($"Processing run: start={runStart}, end={runEnd}");
            if (relativeStart <= runStart && relativeEnd >= runEnd)
            {
                //Console.WriteLine("Applying full color");
                ApplyBackgroundColor(rate, run, null, null, highlightedText);
            }
            else
            {
                int start = Math.Max(relativeStart, runStart) - runStart;
                // 修正：確実に要素の終端までカバーするよう計算
                int end;
                if (relativeEnd >= runEnd)  // 検索範囲の終端が要素の終端以上の場合
                {
                    // 要素の全長を使用
                    end = runEnd - runStart + 1;
                }
                else
                {
                    // 通常のケース：一文字余計に色付けしないための調整
                    end = Math.Min(relativeEnd + 1, runEnd) - runStart;
                }
                ApplyBackgroundColor(rate, run, start, end, highlightedText);
            }
        }

        /// <summary>
        /// ハイライト結果を検証します
        /// </summary>
        private void VerifyHighlightResult(
            string fullDocText,
            (int beginIndex, int endIndex) result,
            string highlightedText,
            StringBuilder sb,
            bool isDebug)
        {
            string matchedText = SafeSubstring(fullDocText, result.beginIndex, result.endIndex - result.beginIndex + 1);
            bool colorMatched = false;
            if (isDebug) colorMatched = CompareStringsIgnoringWhitespace(highlightedText, matchedText);
            else colorMatched = RoughCompare(highlightedText, matchedText);

            if (!colorMatched)
            {
                sb.AppendLine("警告: 色付け箇所と検索テキストが異なります。");
                sb.AppendLine($"検索テキスト: {matchedText}");
                sb.AppendLine($"色付け箇所: {highlightedText}");
            }
            else if (isDebug && colorMatched)
            {
                sb.AppendLine("色付け箇所と検索テキストが一致しました。");
                sb.AppendLine($"検索テキスト: {matchedText}");
                sb.AppendLine($"色付け箇所: {highlightedText}");
            }
        }

        // 新しく追加するメソッド
        private void ProcessAffectedParagraphsAll(
            List<(int start, int end, Paragraph paragraph)> paragraphRanges,
            (int beginIndex, int endIndex) result,
            double rate,
            string fullDocText,
            StringBuilder sb,
            bool isDebug)
        {
            // Process paragraph

            // 色付けされたテキストを記録するStringBuilderを維持
            StringBuilder highlightedText = new StringBuilder();

            // 最初にすべてのパラグラフが長い検索パターンかチェック
            bool isLongSearchPattern = true;
            foreach (var paragraphRange in paragraphRanges)
            {
                // 検索長がパラグラフ長より長いかチェック
                bool exceedsParagraph = (result.endIndex - result.beginIndex + 1) > (paragraphRange.end - paragraphRange.start + 1);

                if (!exceedsParagraph)
                {
                    isLongSearchPattern = false;
                    break;
                }
            }

            foreach (var paragraphRange in paragraphRanges)
            {
                // 変換マップを構築
                var transformMap = new ParagraphTransformationMap();
                transformMap.SetTestCaseIndex(currentTestCaseIndex);
                string transformedText = transformMap.BuildTransformedText(paragraphRange.paragraph);

                // CreateCombinedTextも生成（位置変換のため）
                string createCombinedText = CreateCombinedText(paragraphRange.paragraph);


                // パラグラフ内での相対位置を計算
                int relativeStart = Math.Max(0, result.beginIndex - paragraphRange.start);


                // 条件判定
                int searchLength = result.endIndex - result.beginIndex + 1;

                int relativeEnd;

                // 長い検索パターンでない場合で、前のパラグラフから続く検索
                if (!isLongSearchPattern &&
                    result.beginIndex < paragraphRange.start &&
                    result.endIndex < paragraphRange.end)
                {
                    // 前のパラグラフから続く短い検索
                    // このパラグラフ内での実際の終了位置を計算
                    relativeEnd = Math.Min(result.endIndex - paragraphRange.start, createCombinedText.Length - 1);
                }
                else
                {
                    // 通常の検索処理
                    // 検索範囲の長さを使用
                    relativeEnd = relativeStart + searchLength - 1;
                    relativeEnd = Math.Min(relativeEnd, createCombinedText.Length - 1);
                }


                // 検索結果がこのパラグラフと重複していない場合はスキップ
                if (relativeStart > relativeEnd || relativeEnd < 0)
                {
                    continue;
                }


                // Process search result

                // CreateCombinedTextの位置をBuildTransformedTextの位置に変換
                var (buildStart, buildEnd) = transformMap.ConvertPositionFromCreateCombined(
                    createCombinedText, relativeStart, relativeEnd);

                // Get original positions


                // 位置変換が失敗した場合（パラグラフ範囲外）はスキップ
                if (buildStart == -1 || buildEnd == -1)
                {
                    continue;
                }

                // 変換後の位置から元の要素を特定
                var originalPositions = transformMap.GetOriginalPositions(buildStart, buildEnd);

                // Process each Run

                // 特定された要素に色付け
                foreach (var element in originalPositions)
                {
                    // Apply color to element

                    if (element.IsMathElement)
                    {
                        // 数式要素の場合
                        if (element.MathElement != null && element.MathElement.Parent != null)
                        {
                            ApplyBackgroundColor(rate, element.MathElement, highlightedText);
                        }
                    }
                    else
                    {
                        // Run要素の場合
                        if (element.Run != null && element.Run.Parent != null)
                        {
                            // ProcessMatchedElementと同様の処理を行うが、Runに対して直接処理
                            if (element.StartOffset == 0 && element.EndOffset == element.Run.InnerText.Length - 1)
                            {
                                // Run全体を色付け
                                ApplyBackgroundColor(rate, element.Run, null, null, highlightedText);
                            }
                            else
                            {
                                // 部分的に色付け
                                ApplyBackgroundColor(rate, element.Run, element.StartOffset, element.EndOffset + 1, highlightedText);
                            }
                        }
                        else if (element.Element != null && element.Element.LocalName == "r")
                        {
                            // smartTag内のRun要素の場合 - GetParentRunで変換
                            var convertedRun = GetParentRun(element.Element);
                            if (convertedRun != null)
                            {
                                if (element.StartOffset == 0 && element.EndOffset == element.Element.InnerText.Length - 1)
                                {
                                    ApplyBackgroundColor(rate, convertedRun, null, null, highlightedText);
                                }
                                else
                                {
                                    ApplyBackgroundColor(rate, convertedRun, element.StartOffset, element.EndOffset + 1, highlightedText);
                                }
                            }
                            else
                            {
                            }
                        }
                    }
                }
            }

            // 既存のVerifyHighlightResult呼び出しを維持
            VerifyHighlightResult(fullDocText, result, highlightedText.ToString(), sb, isDebug);
        }

        // 複数パラグラフのテキストを連結して返す新しいメソッド
        private string CreateCombinedTextAll(List<(int start, int end, Paragraph paragraph)> paragraphRanges)
        {
            StringBuilder combinedTextBuilder = new StringBuilder();
            foreach (var range in paragraphRanges)
            {
                string paragraphText = CreateCombinedText(range.paragraph);
                combinedTextBuilder.Append(paragraphText);
            }
            return combinedTextBuilder.ToString();
        }

        // 複数パラグラフの要素範囲をまとめて取得する新しいメソッド
        private List<(int start, int end, string displayText, OpenXmlElement element)> GetElementRangesAll(List<(int start, int end, Paragraph paragraph)> paragraphRanges)
        {
            List<(int start, int end, string displayText, OpenXmlElement element)> allElementRanges = new List<(int start, int end, string displayText, OpenXmlElement element)>();
            int currentPosition = 0;

            foreach (var range in paragraphRanges)
            {
                var elementRanges = GetElementRanges(range.paragraph);
                foreach (var elem in elementRanges)
                {
                    allElementRanges.Add((
                        elem.start + currentPosition,
                        elem.end + currentPosition,
                        elem.displayText,
                        elem.element
                    ));
                }
                currentPosition += CalculateParagraphTextLength(range.paragraph);
            }

            return allElementRanges;
        }

        private string GetDisplayText(List<(int start, int end, string displayText, OpenXmlElement element)> elementRanges)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var range in elementRanges)
            {
                sb.Append(range.displayText);
            }
            return sb.ToString();
        }

        private List<(int start, int end, string displayText, OpenXmlElement element)> GetElementRanges(Paragraph paragraph)
        {
            var elementRanges = new List<(int start, int end, string displayText, OpenXmlElement element)>();
            int currentPosition = 0;

            // 深さ優先で要素を探索
            ExploreElements(paragraph, elementRanges, ref currentPosition);

            return elementRanges;
        }

        private void ExploreElements(OpenXmlElement element, List<(int start, int end, string displayText, OpenXmlElement element)> ranges, ref int currentPosition)
        {
            foreach (var child in element.Elements())
            {
                // スキップする要素タイプ
                if (child.LocalName == "fldSimple" || child.LocalName == "instrText" ||
                    child.LocalName == "fldChar" || child.LocalName == "proofErr" ||
                    child.LocalName == "bookmarkStart" || child.LocalName == "bookmarkEnd" ||
                    child.LocalName == "commentRangeStart" || child.LocalName == "commentRangeEnd" ||
                    child.LocalName == "pPr" || child.LocalName == "rPr" ||
                    child.LocalName == "tblPr" || child.LocalName == "tblGrid" ||
                    child.LocalName == "tcPr" || child.LocalName == "sectPr" ||
                    child.LocalName == "pos" || child.LocalName == "posOffset" ||
                    child.LocalName == "align" || child.LocalName == "sz" ||
                    child.LocalName == "szCs" || child.LocalName == "widowControl" ||
                    child.LocalName == "numPr")
                {
                    continue;
                }

                // 子要素がない場合（テキストなど）
                if (!child.HasChildren || child is Run)
                {
                    string originalText = child.InnerText;
                    int originalLength = originalText.Length;

                    if (originalLength > 0)
                    {
                        // displayTextは未変換のテキストを保持
                        string text = originalText;

                        if (child is Run run)
                        {
                            // Runの場合は特殊文字変換を適用（ただし、displayTextには未変換を使用）
                            string convertedText = SpecialCharConverter.ConvertSpecialCharactersInRun(run);
                            // displayTextには元のテキストを保持
                            text = originalText;
                        }
                        else if (originalText.Contains("ω") || originalText.Contains("Ω") ||
                                 originalText.Contains("θ") || originalText.Contains("π") ||
                                 originalText.Contains("^") || originalText.Contains("_"))
                        {
                            // 数式要素の場合も未変換のテキストを保持
                            text = originalText;
                        }

                        ranges.Add((currentPosition, currentPosition + originalLength - 1, text, child));
                        currentPosition += originalLength;
                    }
                }
                // 数式要素の場合
                else if (child.LocalName == "oMath")
                {
                    // 数式を変換すると文字列の長さの誤差が出るのでInnerTextをそのまま使用する
                    string mathText = child.InnerText;
                    int mathLength = mathText.Length;

                    if (mathLength > 0)
                    {
                        ranges.Add((currentPosition, currentPosition + mathLength - 1, mathText, child));
                        currentPosition += mathLength;
                    }
                }
                // それ以外の複合要素
                else
                {
                    // 再帰的に子要素を処理
                    ExploreElements(child, ranges, ref currentPosition);
                }
            }
        }

        private void ApplyBackgroundColor(double rate, DocumentFormat.OpenXml.Math.OfficeMath mathElement, StringBuilder highlightedText = null)
        {
            Color color = GetHighlightColor(rate);

            var shading = new Shading()
            {
                Fill = $"{color.R:X2}{color.G:X2}{color.B:X2}",
                Val = ShadingPatternValues.Clear
            };

            var runProperties = mathElement.GetFirstChild<RunProperties>();
            if (runProperties == null)
            {
                runProperties = new RunProperties();
                mathElement.InsertBefore(runProperties, mathElement.FirstChild);
            }

            // ShadingをRunPropertiesに追加
            runProperties.Append(shading);

            // 数式のテキストを取得して追加
            string mathText = SpecialCharConverter.ExtractFromMathElement(mathElement, 0);
            highlightedText?.Append(mathText);
        }

        private string CreateCombinedText(Paragraph paragraph)
        {
            string result = SpecialCharConverter.ConvertSpecialCharactersInParagraph(paragraph);
            result = SpecialCharConverter.ReplaceLine(result);
            return result;
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
                escaped = Regex.Replace(escaped, @"'", @"[‘''']");
                escaped = Regex.Replace(escaped, @"""", @"[""®™–—""]");
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

            string pattern = string.Join(@"\s*", processedWords);
            //Console.WriteLine($"Original search text: {searchText}");
            //Console.WriteLine($"Created pattern: {pattern}");
            return pattern;
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

        private void ApplyBackgroundColor(double rate, Run run, int? startOffset = null, int? endOffset = null, StringBuilder highlightedText = null)
        {
            Color color = GetHighlightColor(rate);
            string originalText = run.InnerText;

            if (startOffset.HasValue || endOffset.HasValue)
            {
                // 部分的な色付けが必要な場合
                int start = startOffset ?? 0;
                int end = endOffset ?? originalText.Length;
                // 元のRunを3つの部分に分割
                if (start > 0)
                {
                    // 前半部分（色付けなし）
                    Run beforeRun = (Run)run.CloneNode(true);
                    beforeRun.RemoveAllChildren();
                    beforeRun.AppendChild(new Text(originalText.Substring(0, start)));
                    run.InsertBeforeSelf(beforeRun);
                }

                // 色付け部分
                Run coloredRun = (Run)run.CloneNode(true);
                coloredRun.RemoveAllChildren();
                string coloredText = SafeSubstring(originalText, start, end - start);
                coloredRun.AppendChild(new Text(coloredText));

                // RunPropertiesの作成と色付け
                if (coloredRun.RunProperties == null)
                {
                    coloredRun.RunProperties = new RunProperties();
                }

                var shading = new Shading()
                {
                    Fill = $"{color.R:X2}{color.G:X2}{color.B:X2}",
                    Color = "auto",
                    Val = ShadingPatternValues.Clear
                };

                coloredRun.RunProperties.InsertAt(shading, 0);
                run.InsertBeforeSelf(coloredRun);

                // 色付けしたテキストを追加
                highlightedText?.Append(coloredText);

                if (end < originalText.Length)
                {
                    // 後半部分（色付けなし）
                    Run afterRun = (Run)run.CloneNode(true);
                    afterRun.RemoveAllChildren();
                    afterRun.AppendChild(new Text(SafeSubstring(originalText, end)));
                    run.InsertBeforeSelf(afterRun);
                }

                // 元のRunを削除
                run.Remove();
            }
            else
            {
                // 全体を色付けする場合
                if (run.RunProperties == null)
                {
                    run.RunProperties = new RunProperties();
                    //Console.WriteLine("Created new RunProperties");
                }

                var existingShading = run.RunProperties.GetFirstChild<Shading>();
                if (existingShading != null)
                {
                    existingShading.Remove();
                    //Console.WriteLine("Removed existing shading");
                }

                Shading shading = new Shading()
                {
                    Fill = $"{color.R:X2}{color.G:X2}{color.B:X2}",
                    Color = "auto",
                    Val = ShadingPatternValues.Clear
                };

                run.RunProperties.InsertAt(shading, 0);
                // Console.WriteLine($"Inserted new shading with color: {color.R:X2}{color.G:X2}{color.B:X2}");

                // 全体のテキストをhighlightedTextに追加
                highlightedText?.Append(originalText);
            }
        }

        private bool DoRangesOverlap(int start1, int end1, int start2, int end2)
        {
            // デバッグ出力を追加
            //Console.WriteLine($"Comparing ranges: ({start1}, {end1}) vs ({start2}, {end2})");
            bool overlaps = start1 <= end2 && end1 >= start2;  // 等号を追加
                                                               //Console.WriteLine($"Overlaps: {overlaps}");

            return overlaps;
        }

        private void DebugMathStructure(OpenXmlElement element, StringBuilder debug, string indent)
        {
            debug.AppendLine($"{indent}Element: {element.LocalName}");

            if (element is Run run)
            {
                debug.AppendLine($"{indent}Run Content: '{run.InnerText}'");
                if (run.RunProperties != null)
                {
                    debug.AppendLine($"{indent}Run Properties:");
                    foreach (var prop in run.RunProperties.ChildElements)
                    {
                        debug.AppendLine($"{indent}  {prop.LocalName}: {prop.InnerText}");
                    }
                }
            }

            foreach (var child in element.Elements())
            {
                DebugMathStructure(child, debug, indent + "  ");
            }
        }

        // 数式要素の構造を出力する補助メソッド
        private void DumpMathElement(OpenXmlElement element, int depth, StringBuilder log)
        {
            string indent = new string(' ', depth * 2);
            log.AppendLine($"{indent}Element: {element.LocalName}");

            if (element is Run run)
            {
                log.AppendLine($"{indent}Run Text: {run.InnerText}");
                var length = CalculateElementLength(run);
                log.AppendLine($"{indent}Calculated Length: {length}");
            }

            foreach (var child in element.Elements())
            {
                DumpMathElement(child, depth + 1, log);
            }
        }

        // CalculateMathLengthメソッドにログ出力を追加
        private int CalculateMathLength(DocumentFormat.OpenXml.Math.OfficeMath officeMath)
        {
            StringBuilder debugLog = new StringBuilder();
            debugLog.AppendLine("\n=== CalculateMathLength Debug ===");
            int length = 0;

            foreach (var child in officeMath.Elements())
            {
                int childLength = 0;
                if (child is Run mathRun)
                {
                    childLength = mathRun.InnerText.Length;
                    debugLog.AppendLine($"Math Run Text: {mathRun.InnerText}, Length: {childLength}");
                }
                else if (child is OpenXmlCompositeElement composite)
                {
                    childLength = CalculateCompositeElementLength(composite);
                    //debugLog.AppendLine($"Composite Element: {child.LocalName}, Length: {childLength}");
                }
                length += childLength;
            }

            //debugLog.AppendLine($"Total Math Length: {length}");
            //Console.WriteLine(debugLog.ToString());
            return length;
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
            string str1WithoutWhitespace = SpecialCharConverter.RemoveSymbolsAll(str1);
            string str2WithoutWhitespace = SpecialCharConverter.RemoveSymbolsAll(str2);
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
                // 数字を含むパターンを正確にマッチさせるために、パターンを調整
                pattern = Regex.Replace(pattern, @"(\d+)", @"\s*$1\s*");

                Regex regex = new Regex(pattern, RegexOptions.Compiled | RegexOptions.Multiline);
                MatchCollection matches = regex.Matches(text);

                // テストケース26の場合、検索対象の末尾付近を表示

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
                            newRun.AppendChild(new Text(SafeSubstring(runText, 0, colorStart)));
                        }

                        var coloredRun = (Run)run.Clone();
                        coloredRun.RemoveAllChildren();
                        coloredRun.AppendChild(new Text(SafeSubstring(runText, colorStart, colorEnd - colorStart)));
                        ApplyBackgroundColor(rate, coloredRun);
                        newRun.AppendChild(coloredRun);

                        if (colorEnd < runLength)
                        {
                            newRun.AppendChild(new Text(SafeSubstring(runText, colorEnd)));
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

        private int CalculateElementLength(OpenXmlElement element)
        {
            StringBuilder debug = new StringBuilder();
            debug.AppendLine($"\nCalculateElementLength for {element.LocalName}:");
            debug.AppendLine($"Raw text: {element.InnerText}");

            if (element is DocumentFormat.OpenXml.Math.OfficeMath officeMath)
            {
                int mathLength = CalculateMathLength(officeMath);
                debug.AppendLine($"Math structure:");
                foreach (var child in officeMath.Descendants())
                {
                    //debug.AppendLine($"  - {child.LocalName}: {child.InnerText}");
                    if (child is Run run)
                    {
                        debug.AppendLine($"    Text content: {run.InnerText}");
                    }
                }
                debug.AppendLine($"Calculated math length: {mathLength}");
                //Console.WriteLine(debug.ToString());
                return mathLength;
            }

            int length = element.InnerText.Length;
            debug.AppendLine($"Standard length: {length}");
            //Console.WriteLine(debug.ToString());
            return length;
        }

        private void ApplyColorToMathElement(DocumentFormat.OpenXml.Math.OfficeMath mathElement, int elemStart, int startIndex, int endIndex, double rate)
        {
            foreach (var child in mathElement.Elements())
            {
                if (child is Run run)
                {
                    int runStart = elemStart;
                    int runEnd = runStart + run.InnerText.Length;
                    if (DoRangesOverlap(runStart, runEnd - 1, startIndex, endIndex))
                    {
                        ApplyBackgroundColor(rate, run);
                    }
                }
                else if (child.LocalName == "sPre" ||
                         child.LocalName == "sSubSup" ||
                         child.LocalName == "sSub" ||
                         child.LocalName == "sSup")
                {
                    // 数式のプロパティ要素は保持
                    continue;
                }
                else if (child is OpenXmlCompositeElement composite)
                {
                    ApplyColorToMathElement((DocumentFormat.OpenXml.Math.OfficeMath)composite, elemStart, startIndex, endIndex, rate);
                    elemStart += CalculateMathElementLength(composite);
                }
            }
        }

        private int CalculateMathElementLength(OpenXmlElement element)
        {
            if (element is Run run)
            {
                return run.InnerText.Length;
            }
            else if (element.LocalName == "sPre" ||
                     element.LocalName == "sSubSup" ||
                     element.LocalName == "sSub" ||
                     element.LocalName == "sSup")
            {
                return 0; // プロパティ要素は長さに含めない
            }
            else if (element is OpenXmlCompositeElement composite)
            {
                int length = 0;
                foreach (var child in composite.Elements())
                {
                    length += CalculateMathElementLength(child);
                }
                return length;
            }
            return 0;
        }


        private int CalculateCompositeElementLength(OpenXmlCompositeElement element)
        {
            int length = 0;

            // SuperscriptやSubscriptの代わりにOpenXmlCompositeElementとして処理
            foreach (var child in element.Elements())
            {
                if (child is Run run)
                {
                    length += run.InnerText.Length;
                }
                else if (child is OpenXmlCompositeElement composite)
                {
                    length += CalculateCompositeElementLength(composite);
                }
            }

            // 要素の種類に応じて追加の長さを計算
            switch (element.LocalName.ToLower())
            {
                case "ssup": // 上付き
                    length += 2; // ^() の分
                    break;
                case "ssub": // 下付き
                    length += 2; // _() の分
                    break;
                case "ssubsup": // 上付きと下付きの組み合わせ
                    length += 4; // _()^() の分
                    break;
                case "f": // 分数
                    length += 1; // 分数線の分
                    break;
            }

            return length;
        }

        private Run GetParentRun(OpenXmlElement element)
        {
            var current = element;
            while (current != null)
            {
                if (current is Run run)
                {
                    return run;
                }
                if (current.LocalName == "r")
                {
                    var newRun = new Run();
                    foreach (var child in current.Elements())
                    {
                        newRun.AppendChild(child.CloneNode(true));
                    }
                    var parent = current.Parent;
                    if (parent != null)
                    {
                        parent.ReplaceChild(newRun, current);
                        return newRun;
                    }
                    return null;
                }
                //Console.WriteLine($"要素の種別: {current.GetType().Name}, LocalName: {current.LocalName}");
                current = current.Parent;
            }
            return null;
        }

        public bool RoughCompare(string FirstString, string SecondString)
        {
            string str1 = SpecialCharConverter.RemoveSymbolsAll(FirstString);
            string str2 = SpecialCharConverter.RemoveSymbolsAll(SecondString);
            if (null == str1 || null == str2) return false;

            // 文字列を半角スペースで分割し、空の単語を除去
            string[] firstWords = str1.Split(' ').Where(w => !string.IsNullOrEmpty(w)).ToArray();
            string[] secondWords = str2.Split(' ').Where(w => !string.IsNullOrEmpty(w)).ToArray();

            if (firstWords.Length == 0) return secondWords.Length == 0;

            // 動的計画法による最長共通部分列(LCS)を求める
            int[,] dp = new int[firstWords.Length + 1, secondWords.Length + 1];
            bool[,] matched = new bool[firstWords.Length + 1, secondWords.Length + 1];

            for (int i = 1; i <= firstWords.Length; i++)
            {
                for (int j = 1; j <= secondWords.Length; j++)
                {
                    // 前方一致の判定
                    bool isMatch = firstWords[i - 1].StartsWith(secondWords[j - 1]) ||
                                   secondWords[j - 1].StartsWith(firstWords[i - 1]);

                    if (isMatch)
                    {
                        dp[i, j] = dp[i - 1, j - 1] + 1;
                        matched[i, j] = true;
                    }
                    else
                    {
                        if (dp[i - 1, j] > dp[i, j - 1])
                        {
                            dp[i, j] = dp[i - 1, j];
                            matched[i, j] = false;
                        }
                        else
                        {
                            dp[i, j] = dp[i, j - 1];
                            matched[i, j] = false;
                        }
                    }
                }
            }

            // 一致した単語を追跡して確認
            List<string> matchedWords = new List<string>();
            int x = firstWords.Length;
            int y = secondWords.Length;

            while (x > 0 && y > 0)
            {
                if (matched[x, y])
                {
                    matchedWords.Add(firstWords[x - 1]);
                    x--; y--;
                }
                else if (dp[x - 1, y] > dp[x, y - 1])
                {
                    x--;
                }
                else
                {
                    y--;
                }
            }

            // 一致率を計算
            double matchRate = (double)matchedWords.Count / firstWords.Length;

            // デバッグ情報（実際の使用時には削除またはログに記録）
            /* 
            Console.WriteLine($"First words: {string.Join(", ", firstWords)}");
            Console.WriteLine($"Second words: {string.Join(", ", secondWords)}");
            Console.WriteLine($"Matched words: {string.Join(", ", matchedWords)}");
            Console.WriteLine($"Match rate: {matchRate} ({matchedWords.Count}/{firstWords.Length})");
            */

            return matchRate >= 0.8;
        }

        private Dictionary<Paragraph, int> BuildParagraphPositionMap(Body body)
        {
            var positionMap = new Dictionary<Paragraph, int>();
            int currentPosition = 0;

            // Body直下の全要素を順番に処理
            foreach (var element in body.Elements())
            {
                if (element is Paragraph paragraph)
                {
                    // パラグラフの開始位置を記録
                    positionMap[paragraph] = currentPosition;

                    // パラグラフ内の実際のテキスト長を計算
                    currentPosition += CalculateParagraphTextLength(paragraph);
                }
                else if (element is Table table)
                {
                    // テーブル内のパラグラフも処理
                    ProcessTableForPositionMap(table, positionMap, ref currentPosition);
                }
            }

            return positionMap;
        }

        private int CalculateParagraphTextLength(Paragraph paragraph)
        {
            int totalLength = 0;

            foreach (var element in paragraph.Elements())
            {
                totalLength += GetElementRawTextLength(element);
            }

            return totalLength;
        }

        private int GetElementRawTextLength(OpenXmlElement element)
        {
            if (element is Run run)
            {
                int length = 0;
                foreach (var child in run.ChildElements)
                {
                    if (child is Text text)
                    {
                        // 実際のテキスト長をそのまま使用
                        length += text.Text?.Length ?? 0;
                    }
                    else if (child is TabChar)
                    {
                        length += 1; // タブは1文字として計算
                    }
                    else if (child is Break || child is CarriageReturn)
                    {
                        length += 1; // 改行も1文字として計算
                    }
                    else if (child is SymbolChar)
                    {
                        length += 1; // シンボル文字は1文字として計算
                    }
                }
                return length;
            }
            else if (element is Break)
            {
                return 1;
            }
            else if (element.LocalName == "oMath" || IsMathElement(element))
            {
                // 数式要素は内部テキストの長さを使用
                return element.InnerText.Length;
            }
            else
            {
                // その他の要素は子要素を再帰的に処理
                int totalLength = 0;
                foreach (var child in element.Elements())
                {
                    totalLength += GetElementRawTextLength(child);
                }
                return totalLength;
            }
        }

        private void ProcessTableForPositionMap(Table table, Dictionary<Paragraph, int> positionMap, ref int currentPosition)
        {
            foreach (var row in table.Elements<TableRow>())
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    foreach (var paragraph in cell.Elements<Paragraph>())
                    {
                        positionMap[paragraph] = currentPosition;
                        currentPosition += CalculateParagraphTextLength(paragraph);
                    }
                }
            }
        }

        private Paragraph FindParagraphByPosition(int absolutePosition, Dictionary<Paragraph, int> positionMap)
        {
            Paragraph result = null;
            int maxStartPos = -1;

            foreach (var kvp in positionMap)
            {
                if (kvp.Value <= absolutePosition && kvp.Value > maxStartPos)
                {
                    maxStartPos = kvp.Value;
                    result = kvp.Key;
                }
            }

            return result;
        }

        private bool IsMathElement(OpenXmlElement element)
        {
            string typeName = element.GetType().FullName;
            return typeName.StartsWith("DocumentFormat.OpenXml.Math.") &&
                   !typeName.EndsWith("Properties") &&
                   typeName != "DocumentFormat.OpenXml.Math.BeginChar" &&
                   typeName != "DocumentFormat.OpenXml.Math.EndChar";
        }
    }
}
