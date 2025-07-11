using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Arx.DocSearch.Util
{
    public class PositionedElement
    {
        public Run Run { get; set; }
        public OpenXmlElement Element { get; set; }  // 元の要素への参照（smartTag内のRun要素など）
        public DocumentFormat.OpenXml.Math.OfficeMath MathElement { get; set; }
        public bool IsMathElement { get; set; }
        public bool IsFieldElement { get; set; }
        public int StartOffset { get; set; }
        public int EndOffset { get; set; }
    }
}
