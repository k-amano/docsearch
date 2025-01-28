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
            string pattern;

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    Body body = doc.MainDocumentPart.Document.Body;
                    if (body != null)
                    {
                        var paragraphs = body.Descendants<Paragraph>().ToList();
                        if (isDebug) sb.AppendLine($"filePath:\n{filePath}\ndocText:\n{wte.Text}");

                        // 全パラグラフのテキストを連結
                        StringBuilder combinedDocText = new StringBuilder();
                        List<(int start, int end, Paragraph paragraph)> paragraphRanges = new List<(int start, int end, Paragraph paragraph)>();
                        int currentPosition = 0;

                        foreach (var paragraph in paragraphs)
                        {
                            var (paragraphText, _) = CreateCombinedText(paragraph);
                            if (!string.IsNullOrWhiteSpace(paragraphText))
                            {
                                paragraphRanges.Add((currentPosition, currentPosition + paragraphText.Length - 1, paragraph));
                                combinedDocText.Append(paragraphText);
                                currentPosition += paragraphText.Length;
                            }
                        }

                        string fullDocText = combinedDocText.ToString();
                        Console.WriteLine("fullDocText:\n" + fullDocText);

                        for (int i = 0; i < searchPatterns.Length && i < rates.Length; i++)
                        {
                            string searchPattern = Regex.Replace(searchPatterns[i], @"^[0-9]+\.?\s+", "");
                            searchPattern = Regex.Replace(searchPattern, @"\s+[0-9]+\.?\s*$", "");
                            string[] words = searchPattern.Split(' ');
                            if (searchPattern.Length < 20 && words.Length < 3) continue;
                            pattern = CreateSearchPattern(searchPattern);
                            bool foundInDocument = false;

                            var results = MatchIgnoringWhitespace(pattern, fullDocText, sb);

                            if (results.Count > 0)
                            {
                                foundInDocument = true;
                                if (isDebug)
                                {
                                    sb.AppendLine("エラー: 指定されたテキストが見つかりました。");
                                    sb.AppendLine($"検索文:searchPatterns[{i}] {searchPatterns[i]}\n{pattern}");
                                }

                                foreach (var result in results)
                                {
                                    // 該当する範囲に含まれるパラグラフを特定
                                    var affectedParagraphs = paragraphRanges
                                        .Where(p => DoRangesOverlap(result.beginIndex, result.endIndex, p.start, p.end))
                                        .ToList();

                                    foreach (var paragraphRange in affectedParagraphs)
                                    {
                                        var paragraph = paragraphRange.paragraph;
                                        var (combinedText, elementRanges) = CreateCombinedText(paragraph);

                                        // パラグラフ内での相対位置を計算
                                        int relativeStart = Math.Max(0, result.beginIndex - paragraphRange.start);
                                        int relativeEnd = Math.Min(combinedText.Length - 1, result.endIndex - paragraphRange.start);

                                        // パラグラフ内の該当範囲を色付け
                                        var matchedElements = elementRanges
                                            .Where(r => DoRangesOverlap(relativeStart, relativeEnd, r.start, r.end))
                                            .ToList();

                                        StringBuilder highlightedText = new StringBuilder();

                                        foreach (var elem in matchedElements)
                                        {
                                            if (elem.element is Run run)
                                            {
                                                int runStart = elem.start;
                                                int runEnd = elem.end;

                                                if (relativeStart <= runStart && relativeEnd >= runEnd)
                                                {
                                                    ApplyBackgroundColor(rates[i], run, null, null, highlightedText);
                                                }
                                                else
                                                {
                                                    int start = Math.Max(relativeStart, runStart) - runStart;
                                                    int end = Math.Min(relativeEnd + 1, runEnd) - runStart;
                                                    ApplyBackgroundColor(rates[i], run, start, end, highlightedText);
                                                }
                                            }
                                            else if (elem.element is DocumentFormat.OpenXml.Math.OfficeMath math)
                                            {
                                                ApplyBackgroundColor(rates[i], math, highlightedText);
                                            }
                                        }

                                        // 色付け結果の確認
                                        string matchedText = fullDocText.Substring(result.beginIndex, result.endIndex - result.beginIndex + 1);
                                        bool colorMatched = CompareStringsIgnoringWhitespace(highlightedText.ToString(), matchedText);
                                        if (!colorMatched)
                                        {
                                            sb.AppendLine("警告: 色付け箇所と検索テキストが異なります。");
                                            sb.AppendLine($"検索テキスト: {matchedText}");
                                            sb.AppendLine($"色付け箇所: {highlightedText.ToString()}");
                                        }
                                        else if (isDebug && colorMatched)
                                        {
                                            sb.AppendLine("色付け箇所と検索テキストが一致しました。");
                                            sb.AppendLine($"検索テキスト: {matchedText}");
                                            sb.AppendLine($"色付け箇所: {highlightedText.ToString()}");
                                        }
                                    }
                                }
                            }

                            if (!foundInDocument)
                            {
                                sb.AppendLine("エラー: 指定されたテキストが見つかりませんでした。");
                                sb.AppendLine($"検索文: searchPatterns[{i}]{searchPatterns[i]}\n{pattern}");
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
            StringBuilder mathText = new StringBuilder();
            int currentPosition = 0; // 新しい変数を追加
            ProcessElements(mathElement, new List<(int start, int end, OpenXmlElement element)>(), mathText, ref currentPosition);
            highlightedText?.Append(mathText.ToString());
        }

        private (string combinedText, List<(int start, int end, OpenXmlElement element)>)
        CreateCombinedText(Paragraph paragraph)
        {
            var elementRanges = new List<(int start, int end, OpenXmlElement element)>();
            StringBuilder combinedText = new StringBuilder();
            int currentPosition = 0;

            ProcessElements(paragraph, elementRanges, combinedText, ref currentPosition);

            string combinedTextString = SpecialCharConverter.ReplaceLine(combinedText.ToString());

            return (combinedTextString, elementRanges);
        }

        private void ProcessElements(OpenXmlElement parentElement,
           List<(int start, int end, OpenXmlElement element)> elementRanges,
           StringBuilder combinedText,
           ref int currentPosition)
        {
            foreach (var element in parentElement.Elements())
            {
                string elementText = "";
                string typeName = element.GetType().FullName;

                if (typeName == "DocumentFormat.OpenXml.Wordprocessing.Run")
                {
                    var run = (Run)element;
                    elementText = SpecialCharConverter.ConvertSpecialCharactersInRun(run);
                    if (!string.IsNullOrEmpty(elementText))
                    {
                        elementRanges.Add((currentPosition, currentPosition + elementText.Length, run));
                        combinedText.Append(elementText);
                        currentPosition += elementText.Length;
                    }
                }
                else if (typeName == "DocumentFormat.OpenXml.Wordprocessing.Text")
                {
                    var text = (DocumentFormat.OpenXml.Wordprocessing.Text)element;
                    elementText = text.Text;
                    if (!string.IsNullOrEmpty(elementText))
                    {
                        elementRanges.Add((currentPosition, currentPosition + elementText.Length, text));
                        combinedText.Append(elementText);
                        currentPosition += elementText.Length;
                    }
                }
                else if (typeName == "DocumentFormat.OpenXml.Math.Text")
                {
                    var text = (DocumentFormat.OpenXml.Math.Text)element;
                    elementText = text.Text;
                    if (!string.IsNullOrEmpty(elementText))
                    {
                        elementRanges.Add((currentPosition, currentPosition + elementText.Length, text));
                        combinedText.Append(elementText);
                        currentPosition += elementText.Length;
                    }
                }
                else if (typeName == "DocumentFormat.OpenXml.Math.OfficeMath" ||
                         typeName == "DocumentFormat.OpenXml.Math.Base" ||
                         typeName == "DocumentFormat.OpenXml.Math.SubArgument" ||
                         typeName == "DocumentFormat.OpenXml.Math.SuperArgument" ||
                         typeName == "DocumentFormat.OpenXml.Math.FunctionName" ||
                         typeName == "DocumentFormat.OpenXml.Math.Numerator" ||
                         typeName == "DocumentFormat.OpenXml.Math.Denominator" ||
                         typeName == "DocumentFormat.OpenXml.Math.Delimiter" ||
                         typeName == "DocumentFormat.OpenXml.Math.Matrix" ||
                         typeName == "DocumentFormat.OpenXml.Math.MatrixRow" ||
                         typeName == "DocumentFormat.OpenXml.Math.MathFunction" ||
                         typeName == "DocumentFormat.OpenXml.Math.Subscript" ||
                         typeName == "DocumentFormat.OpenXml.Math.Superscript" ||
                         typeName == "DocumentFormat.OpenXml.Math.Run" ||
                         typeName == "DocumentFormat.OpenXml.Math.SubSuperscript" ||
                         typeName == "DocumentFormat.OpenXml.Math.SubscriptProperties") // 追加
                {
                    string mathText = SpecialCharConverter.ExtractFromMathElement(element, 0);
                    if (!string.IsNullOrEmpty(mathText))
                    {
                        elementRanges.Add((currentPosition, currentPosition + mathText.Length, element));
                        combinedText.Append(mathText);
                        currentPosition += mathText.Length;
                    }
                    else
                    {
                        ProcessElements(element, elementRanges, combinedText, ref currentPosition);
                    }
                }
                else if (typeName.Contains("Properties") || // この条件を.Math.SubscriptPropertiesの後に移動
                         typeName.Contains("MatrixColumn") ||
                         typeName == "DocumentFormat.OpenXml.Math.BeginChar" ||
                         typeName == "DocumentFormat.OpenXml.Math.EndChar" ||
                         typeName == "DocumentFormat.OpenXml.Wordprocessing.Style" ||
                         typeName == "DocumentFormat.OpenXml.Wordprocessing.BookmarkStart")
                {
                    continue;
                }
                else
                {
                    //Console.WriteLine($"真に予期しない要素タイプ: {element.GetType().Name}");
                    ProcessElements(element, elementRanges, combinedText, ref currentPosition);
                }
            }
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
                string coloredText = originalText.Substring(start, end - start);
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
                    afterRun.AppendChild(new Text(originalText.Substring(end)));
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
                }

                var existingShading = run.RunProperties.GetFirstChild<Shading>();
                if (existingShading != null)
                {
                    existingShading.Remove();
                }

                Shading shading = new Shading()
                {
                    Fill = $"{color.R:X2}{color.G:X2}{color.B:X2}",
                    Color = "auto",
                    Val = ShadingPatternValues.Clear
                };

                run.RunProperties.InsertAt(shading, 0);

                // 色付けしたテキストを追加
                highlightedText?.Append(originalText);
            }
        }

        private bool DoRangesOverlap(int start1, int end1, int start2, int end2)
        {
            // 範囲のオーバーラップをより正確に判定
            bool overlaps = (start1 <= end2 && start2 <= end1);
            //Console.WriteLine($"  Range overlap check: ({start1},{end1}) vs ({start2},{end2}) = {overlaps}");
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
    }
}
