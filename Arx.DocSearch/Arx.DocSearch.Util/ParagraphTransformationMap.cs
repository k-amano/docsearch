using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Arx.DocSearch.Util
{
    public class ParagraphTransformationMap
    {
        private List<RunTransformation> transformations = new List<RunTransformation>();
        private int currentTestCaseIndex = -1;
        private string buildTransformedText; // 追加

        public void SetTestCaseIndex(int index)
        {
            currentTestCaseIndex = index;
        }

        private class RunTransformation
        {
            public Run Run { get; set; }
            public OpenXmlElement Element { get; set; }  // 元の要素への参照（smartTag内のRun要素など）
            public DocumentFormat.OpenXml.Math.OfficeMath MathElement { get; set; }  // 数式要素
            public bool IsMathElement { get; set; }  // 数式要素かどうか
            public bool IsFieldElement { get; set; }  // フィールド要素かどうか
            public int OriginalStart { get; set; }
            public int OriginalLength { get; set; }
            public int TransformedStart { get; set; }
            public int TransformedLength { get; set; }
            public int ElementIndex { get; set; }  // 要素インデックスを追加
        }

        public string BuildTransformedText(Paragraph paragraph)
        {
            StringBuilder result = new StringBuilder();
            int currentOriginalPos = 0;
            int currentTransformedPos = 0;

            // Process paragraph elements

            // Process all child elements, not just Runs
            int elementIndex = 0;


            foreach (var element in paragraph.ChildElements)
            {

                string original = "";
                string transformed = "";
                Run runElement = null;

                if (element is Run run)
                {
                    runElement = run;
                    original = run.InnerText;
                    transformed = SpecialCharConverter.ConvertSpecialCharactersInRun(run);

                    // Process Run element
                }
                else if (element.LocalName == "oMath" || element.LocalName == "oMathPara")
                {
                    // Handle math elements - use the entire math content
                    original = element.InnerText;
                    transformed = original; // Math content doesn't need special character conversion

                    // Process Math element
                    var mathElement = element as DocumentFormat.OpenXml.Math.OfficeMath;

                    if (mathElement != null && !string.IsNullOrEmpty(original))
                    {
                        transformations.Add(new RunTransformation
                        {
                            Run = null,
                            MathElement = mathElement,
                            IsMathElement = true,
                            OriginalStart = currentOriginalPos,
                            OriginalLength = original.Length,
                            TransformedStart = currentTransformedPos,
                            TransformedLength = transformed.Length,
                            ElementIndex = elementIndex
                        });
                    }

                    result.Append(transformed);
                    currentOriginalPos += original.Length;
                    currentTransformedPos += transformed.Length;
                }
                else if (element.LocalName == "fldSimple")
                {
                    // Handle field elements - use InnerText
                    original = element.InnerText;
                    transformed = original;

                    // フィールド要素として記録（色付け不可）
                    if (!string.IsNullOrEmpty(original))
                    {
                        transformations.Add(new RunTransformation
                        {
                            Run = null,
                            MathElement = null,
                            IsMathElement = false,
                            IsFieldElement = true,
                            OriginalStart = currentOriginalPos,
                            OriginalLength = original.Length,
                            TransformedStart = currentTransformedPos,
                            TransformedLength = transformed.Length,
                            ElementIndex = elementIndex
                        });

                        result.Append(transformed);
                        currentOriginalPos += original.Length;
                        currentTransformedPos += transformed.Length;
                    }
                }
                else if (element.LocalName == "smartTag" || element is OpenXmlUnknownElement)
                {
                    // Handle smartTag and other unknown elements - process children recursively

                    // smartTag内の子要素を処理
                    foreach (var child in element.ChildElements)
                    {
                        ProcessChildElement(child, transformations, result, ref currentOriginalPos, ref currentTransformedPos, ref elementIndex);
                    }
                    elementIndex++;
                }
                else if (ShouldSkipElement(element))
                {
                    // Skip layout-related elements that don't contain text
                    elementIndex++;
                    continue;
                }
                else if (element.HasChildren)
                {
                    // Process other elements with children recursively
                    foreach (var child in element.ChildElements)
                    {
                        ProcessChildElement(child, transformations, result, ref currentOriginalPos, ref currentTransformedPos, ref elementIndex);
                    }
                    elementIndex++;
                }
                else
                {
                    // Skip other elements
                    elementIndex++;
                    continue;
                }

                if (runElement != null && !string.IsNullOrEmpty(original))
                {
                    transformations.Add(new RunTransformation
                    {
                        Run = runElement,
                        OriginalStart = currentOriginalPos,
                        OriginalLength = original.Length,
                        TransformedStart = currentTransformedPos,
                        TransformedLength = transformed.Length,
                        ElementIndex = elementIndex  // 要素インデックスを記録
                    });

                    // Add transformation record

                    result.Append(transformed);
                    currentOriginalPos += original.Length;
                    currentTransformedPos += transformed.Length;

                    // Continue processing
                }
            }

            buildTransformedText = result.ToString(); // 保存
                                                      // Build complete

            return buildTransformedText;
        }

        // 子要素を処理するヘルパーメソッド
        private void ProcessChildElement(OpenXmlElement element,
            List<RunTransformation> transformations,
            StringBuilder result,
            ref int currentOriginalPos,
            ref int currentTransformedPos,
            ref int elementIndex)
        {

            if (element is Run childRun)
            {
                string original = childRun.InnerText;
                string transformed = SpecialCharConverter.ConvertSpecialCharactersInRun(childRun);


                if (!string.IsNullOrEmpty(original))
                {
                    transformations.Add(new RunTransformation
                    {
                        Run = childRun,
                        MathElement = null,
                        IsMathElement = false,
                        IsFieldElement = false,
                        OriginalStart = currentOriginalPos,
                        OriginalLength = original.Length,
                        TransformedStart = currentTransformedPos,
                        TransformedLength = transformed.Length,
                        ElementIndex = elementIndex
                    });

                    result.Append(transformed);
                    currentOriginalPos += original.Length;
                    currentTransformedPos += transformed.Length;
                }
            }
            else if (element.LocalName == "r") // Run要素だがキャストできない場合
            {
                string original = element.InnerText;
                if (!string.IsNullOrEmpty(original))
                {
                    // smartTag内のRun要素として処理
                    transformations.Add(new RunTransformation
                    {
                        Run = null, // Run型にキャストできないのでnull
                        Element = element, // 元の要素への参照を保持
                        MathElement = null,
                        IsMathElement = false,
                        IsFieldElement = false,
                        OriginalStart = currentOriginalPos,
                        OriginalLength = original.Length,
                        TransformedStart = currentTransformedPos,
                        TransformedLength = original.Length,
                        ElementIndex = elementIndex
                    });

                    result.Append(original);
                    currentOriginalPos += original.Length;
                    currentTransformedPos += original.Length;
                }
            }
            else if (element.LocalName == "oMath" || element.LocalName == "oMathPara")
            {
                // 数式要素の処理
                string original = element.InnerText;
                string transformed = original;

                var mathElement = element as DocumentFormat.OpenXml.Math.OfficeMath;
                if (mathElement != null && !string.IsNullOrEmpty(original))
                {
                    transformations.Add(new RunTransformation
                    {
                        Run = null,
                        MathElement = mathElement,
                        IsMathElement = true,
                        IsFieldElement = false,
                        OriginalStart = currentOriginalPos,
                        OriginalLength = original.Length,
                        TransformedStart = currentTransformedPos,
                        TransformedLength = transformed.Length,
                        ElementIndex = elementIndex
                    });

                    result.Append(transformed);
                    currentOriginalPos += original.Length;
                    currentTransformedPos += transformed.Length;
                }
            }
            else if (element.LocalName == "smartTagPr")
            {
                // smartTagのプロパティはスキップ
                return;
            }
            else if (element.HasChildren)
            {
                // 再帰的に子要素を処理
                foreach (var child in element.ChildElements)
                {
                    ProcessChildElement(child, transformations, result, ref currentOriginalPos, ref currentTransformedPos, ref elementIndex);
                }
            }
        }

        // スキップすべき要素かどうか判定
        private bool ShouldSkipElement(OpenXmlElement element)
        {
            // 旧バージョンのExploreElementsメソッドと同じスキップリスト
            string localName = element.LocalName;

            // 位置・レイアウト関連要素
            if (localName == "posOffset" || localName == "positionH" || localName == "positionV" ||
                localName == "align" || localName == "extent" || localName == "effectExtent" ||
                localName == "docPr")
            {
                return true;
            }

            // プロパティ関連
            if (localName == "pPr" || localName == "rPr" || localName == "sectPr")
            {
                return true;
            }

            // テーブルキャプション関連（Tableを含む場合）
            if ((localName == "drawing" || localName == "anchor" ||
                 localName == "txbxContent" || localName == "txbx" || localName == "wsp") &&
                element.InnerText.Contains("Table"))
            {
                return true;
            }

            return false;
        }

        // CreateCombinedTextの位置をBuildTransformedTextの位置に変換
        public (int start, int end) ConvertPositionFromCreateCombined(
            string createCombinedText,
            int createStart,
            int createEnd)
        {

            // 1. 範囲のクリッピング処理
            // パラグラフの範囲を超える場合は、パラグラフ内に収まるようにクリップする
            int clippedStart = createStart;
            int clippedEnd = createEnd;

            if (clippedStart >= createCombinedText.Length)
            {
                // 開始位置がパラグラフを超えている場合、このパラグラフには該当なし
                return (-1, -1);
            }

            if (clippedEnd >= createCombinedText.Length)
            {
                // 終了位置がパラグラフを超えている場合、パラグラフの終端までにクリップ
                clippedEnd = createCombinedText.Length - 1;
            }

            // 順次照合ロジックに基づく実装
            int createPos = 0;      // CreateCombinedText内の現在位置
            int buildPos = 0;       // BuildTransformedText内の現在位置
            int buildStart = -1;
            int buildEnd = -1;

            // 両テキストを先頭から照合
            while (createPos < createCombinedText.Length && buildPos < buildTransformedText.Length)
            {
                // 開始位置に到達したら記録
                if (createPos == clippedStart && buildStart == -1)
                {
                    buildStart = buildPos;
                }

                // CreateCombinedTextの現在文字を取得
                char createChar = createCombinedText[createPos];
                char buildChar = buildTransformedText[buildPos];

                // 両方の記号・空白を除去して比較
                string createCharNoSymbol = SpecialCharConverter.RemoveSymbolsAll(createChar.ToString());
                string buildCharNoSymbol = SpecialCharConverter.RemoveSymbolsAll(buildChar.ToString());

                // 両方が記号・空白の場合
                if (string.IsNullOrEmpty(createCharNoSymbol) && string.IsNullOrEmpty(buildCharNoSymbol))
                {
                    createPos++;
                    buildPos++;
                }
                // CreateCombined側だけが記号・空白の場合
                else if (string.IsNullOrEmpty(createCharNoSymbol))
                {
                    createPos++;
                }
                // BuildTransformed側だけが記号・空白の場合
                else if (string.IsNullOrEmpty(buildCharNoSymbol))
                {
                    buildPos++;
                }
                // 両方が記号以外の文字の場合
                else
                {
                    // 文字が一致することを確認
                    if (createCharNoSymbol == buildCharNoSymbol)
                    {
                        createPos++;
                        buildPos++;
                    }
                    else
                    {
                        // 不一致の場合はエラー
                        return (0, buildTransformedText.Length - 1);
                    }
                }

                // 終了位置に到達したら記録
                if (createPos == clippedEnd + 1 && buildEnd == -1)
                {
                    buildEnd = buildPos - 1;
                    break;
                }
            }

            // 最後まで処理してもbuildEndが設定されていない場合
            if (buildEnd == -1 && createPos == clippedEnd + 1)
            {
                buildEnd = buildPos - 1;
            }

            // 正常な範囲が取得できたか確認
            if (buildStart == -1 || buildEnd == -1 || buildStart > buildEnd)
            {
                // 範囲が取得できなかった場合、該当なしを返す
                return (-1, -1);
            }

            return (buildStart, buildEnd);
        }

        public List<PositionedElement> GetOriginalPositions(
            int transformedStart, int transformedEnd)
        {
            var results = new List<PositionedElement>();


            // Looking for overlapping transformations

            foreach (var trans in transformations)
            {
                int transEnd = trans.TransformedStart + trans.TransformedLength - 1;

                // Check for overlap

                if (transformedStart <= transEnd && transformedEnd >= trans.TransformedStart)
                {
                    // フィールド要素はスキップ（色付けできない）
                    if (trans.IsFieldElement)
                    {
                        continue;
                    }

                    // 変換後の相対位置
                    int relStart = Math.Max(0, transformedStart - trans.TransformedStart);
                    int relEnd = Math.Min(trans.TransformedLength - 1,
                                          transformedEnd - trans.TransformedStart);

                    // 元のテキストでの位置に変換
                    int origStart, origEnd;

                    if (trans.OriginalLength == trans.TransformedLength)
                    {
                        // 長さが同じなら直接マッピング
                        origStart = relStart;
                        origEnd = relEnd;
                    }
                    else if (trans.OriginalLength < trans.TransformedLength)
                    {
                        // 特殊文字変換で増えた場合
                        if (relStart == 0)
                            origStart = 0;
                        else if (relEnd == trans.TransformedLength - 1)
                            origStart = trans.OriginalLength - 1;
                        else
                            origStart = 0;

                        origEnd = Math.Min(origStart, trans.OriginalLength - 1);
                    }
                    else
                    {
                        // 通常はこのケースは発生しない
                        origStart = Math.Min(relStart, trans.OriginalLength - 1);
                        origEnd = Math.Min(relEnd, trans.OriginalLength - 1);
                    }

                    // Add to results
                    results.Add(new PositionedElement
                    {
                        Run = trans.Run,
                        Element = trans.Element,
                        MathElement = trans.MathElement,
                        IsMathElement = trans.IsMathElement,
                        IsFieldElement = trans.IsFieldElement,
                        StartOffset = origStart,
                        EndOffset = origEnd
                    });

                }
            }

            // Return results
            return results;
        }
    }
}
