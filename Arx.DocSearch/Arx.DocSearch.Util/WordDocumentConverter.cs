using System;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Arx.DocSearch.Util
{
	public class WordDocumentConverter
	{
		public static void ConvertDocToDocx(Application word, string inputPath, string outputPath)
		{
			if (!File.Exists(inputPath))
			{
				throw new FileNotFoundException("指定されたファイルが見つかりません。", inputPath);
			}

			Document doc = null;
			try
			{
				doc = word.Documents.Open(inputPath);
				doc.SaveAs2(outputPath, WdSaveFormat.wdFormatXMLDocument);
			}
			finally
			{
				if (doc != null)
				{
					doc.Close();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
				}
			}
		}
	}
}
