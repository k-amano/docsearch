using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Word;
//using Application = Microsoft.Office.Interop.Word.Application;
//using Document = Microsoft.Office.Interop.Word.Document;
using Color = System.Drawing.Color;
using Arx.DocSearch.Util;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Arx.DocSearch.Client
{
	public class WordConverter
	{
		#region コンストラクタ
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public WordConverter(string srcFile, List<string> lsSrc, string targetFile, List<string> lsTarget, Dictionary<int, MatchLine> matchLines, string seletedPath, MainForm mainForm)
		{
			this.srcFile = srcFile;
			this.targetFile = targetFile;
			this.lsSrc = lsSrc;
			this.lsTarget = lsTarget;
			this.matchLines = matchLines;
			this.seletedPath = seletedPath;
			this.mainForm = mainForm;
		}
		#endregion

		#region フィールド
		private string srcFile;
		private string targetFile;
		List<string> lsSrc;
		List<string> lsTarget;
		private Dictionary<int, MatchLine> matchLines;
		string seletedPath;
		List<List<int>> srcParagraphs;
		List<List<int>> targetParagraphs;
        private MainForm mainForm;
		// ロック用のインスタンス
		private static ReaderWriterLock rwl = new ReaderWriterLock();
		#endregion

		#region メソッド
		public void Run()
		{
			this.GetSrcParagraphs();
			this.GetTargetParagraphs();
			Process[] wordProcesses = Process.GetProcessesByName("WINWORD");
			if (0 < wordProcesses.Length)
			{
				//メッセージボックスを表示する
				DialogResult result = MessageBox.Show("起動中の Word を終了してよろしいですか？",
						"確認",
						MessageBoxButtons.YesNoCancel,
						MessageBoxIcon.Exclamation,
						MessageBoxDefaultButton.Button2);

				//何が選択されたか調べる
				if (result == DialogResult.Yes)
				{
					//「はい」が選択された時
					foreach (Process p in wordProcesses)
					{
						try
						{
							p.Kill();
							p.WaitForExit(); // possibly with a timeout
						}
						catch (Exception e)
						{
							Debug.WriteLine(e.StackTrace);
						}
					}
				} else return;
			}
			this.EditWord(srcFile, false);
			this.EditWord(targetFile, true);

		}

		private void EditWord(string docFile, bool isTarget)
		{
			string targetPath = Path.Combine(this.seletedPath, Path.GetFileName(docFile));
			File.Copy(docFile, targetPath, true);
			string filename = Path.Combine(this.seletedPath, Path.GetFileName(docFile) + ".txt");
			File.Delete(filename);
			WordTextExtractor wte = new WordTextExtractor(targetPath);
			string docText = wte.Text;
			List<int> indexes = new List<int>();
			List<double> rates = new List<double>();
			List<string> searchPatterns = new List<string>();
			foreach (KeyValuePair<int, MatchLine> ml in this.matchLines)
			{
				MatchLine m = ml.Value;
				int index = isTarget ? m.TargetLine : ml.Key;
				double rate = m.Rate;
				string line = isTarget ? this.lsTarget[index].Trim() : this.lsSrc[index].Trim();
				indexes.Add(index);
				searchPatterns.Add(line);
				rates.Add(rate);
			}
			WordTextHighLighter highLighter = new WordTextHighLighter();
			string message = highLighter.HighlightTextInWord(targetPath, indexes.ToArray(), rates.ToArray(), searchPatterns.ToArray());
			if (!string.IsNullOrEmpty(message)) this.WriteMatchLine(message, docFile);
		}

		private void GetSrcParagraphs() {
			this.srcParagraphs = new List<List<int>>();
			List<int> paragraph = new List<int>();
			for (int i = 0; i < this.lsSrc.Count;i++ )
			{
				string line = this.lsSrc[i];
				paragraph.Add(i);
				if (!line.EndsWith("  ")) {
					this.srcParagraphs.Add(paragraph);
					paragraph = new List<int>();
				}
			}
			if (0 < paragraph.Count) this.srcParagraphs.Add(paragraph);
		}

		private void GetTargetParagraphs()
		{
			this.targetParagraphs = new List<List<int>>();
			List<int> paragraph = new List<int>();
			for (int i = 0; i < this.lsTarget.Count; i++)
			{
				string line = this.lsTarget[i];
				paragraph.Add(i);
				if (!line.EndsWith("  "))
				{
					this.targetParagraphs.Add(paragraph);
					paragraph = new List<int>();
				}
			}
			if (0 < paragraph.Count) this.targetParagraphs.Add(paragraph);
		}

		private void WriteMatchLine(string message, string docFile)
		{
			string filename = Path.Combine(this.seletedPath, Path.GetFileName(docFile) + ".txt");
			rwl.AcquireWriterLock(Timeout.Infinite);
			// ファイルオープン
			try
			{
				using (FileStream fs = File.Open(filename, FileMode.Append))
				using (StreamWriter writer = new StreamWriter(fs))
				{
					writer.WriteLine(message);
				}
			}
			finally
			{
				// ロック解除は finally の中で行う
				rwl.ReleaseWriterLock();
			}
		}
	}
	#endregion
}

