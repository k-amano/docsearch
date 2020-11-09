using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Arx.DocSearch
{
	public class WordConverter
	{
				#region コンストラクタ
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public WordConverter(string srcFile, List<string> lsSrc, string targetFile, List<string> lsTarget, Dictionary<int, MatchLine> matchLines, string seletedPath)
		{
			this.srcFile = srcFile;
			this.targetFile = targetFile;
			this.lsSrc = lsSrc;
			this.lsTarget = lsTarget;
			this.matchLines = matchLines;
			this.seletedPath = seletedPath;
		}
		#endregion
		#region フィールド
		private string srcFile;
		private string targetFile;
		List<string> lsSrc;
		List<string> lsTarget;
		private Dictionary<int, MatchLine> matchLines;
		string seletedPath;
		#endregion

		#region メソッド
		public void Run()
		{
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
			Application word = null;
			try
			{
				word = new Application();
				word.Visible = false;
				this.EditWord(word, srcFile, false);
				//this.EditWord(word, targetFile, true);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
			}
			finally
			{
				if (null != word)
				{
					((_Application)word).Quit();
					Marshal.ReleaseComObject(word);  // オブジェクト参照を解放
					word = null;
				}
			}

		}

		private void EditWord(Application word, string docFile, bool isTarget)
		{
			Document doc = null;
			object miss = System.Reflection.Missing.Value;
			object path = docFile;
			object path2 = Path.Combine(this.seletedPath, Path.GetFileName(docFile));
			object readOnly = false;
			int index = 0;
			int pos = 0;
			int lastIndex = -1;
			try
			{
				doc = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
				int i = 0;
				string text = string.Empty;
				string prevText = string.Empty;
				//for (int i = 0; i < doc.Paragraphs.Count; i++)
				foreach (Paragraph para in doc.Paragraphs)
				{
					//if (50 < i || 50 < index) break;
					//Paragraph para = doc.Paragraphs[i + 1];
					//Range r = doc.Paragraphs[i + 1].Range;
					Range r = para.Range;
					foreach (Range r2 in r.Sentences)
					{
						prevText = text;
						if (i < 10) Debug.WriteLine(string.Format("EditWord index={0}:{1}", index, r2.Text));
						Color color;
						text = this.ReplaceLine(r2.Text);
						if (isTarget) color = this.GetTargetColor(text, prevText, ref index, ref pos, ref lastIndex);
						else color = this.GetSrcColor(text, prevText, ref index, ref pos, ref lastIndex);
						//if (61 == index) Debug.WriteLine(string.Format("color={0}", color));
						r2.Font.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(color);
						r2.Select();
					}
					i++;
				}
				//doc.SaveAs(ref path2, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
			}
			finally
			{
				if (null != doc)
				{
					((_Document)doc).Close();
					Marshal.ReleaseComObject(doc);  // オブジェクト参照を解放
					doc = null;
				}
			}
		}

		private Color GetSrcColor(string text, string prevText, ref int index, ref int pos, ref int lastIndex)
		{
			Color ret = Color.White;
			text = text.Trim();
			if (string.IsNullOrEmpty(text)) return ret;
			this.getLineIndex(this.lsSrc, text, prevText, ref index, ref pos, ref lastIndex);
			if (this.lsSrc.Count <= index) return ret;
			//string line = this.lsSrc[index];
			if (this.matchLines.ContainsKey(index))
			{
				if (1D == this.matchLines[index].Rate) ret = Color.LightPink;
				else if (0.9 <= this.matchLines[index].Rate) ret = Color.Yellow;
				else if (0D < this.matchLines[index].Rate) ret = Color.LightGreen;
			}
			return ret;
		}

		private Color GetTargetColor(string text, string prevText, ref int index, ref int pos, ref int lastIndex)
		{
			Color ret = Color.White;
			text = text.Trim();
			if (string.IsNullOrEmpty(text)) return ret;
			this.getLineIndex(this.lsTarget, text, prevText, ref index, ref pos, ref lastIndex);
			if (this.lsTarget.Count <= index) return ret;
			//string line = this.lsTarget[index];
			double rate = getTargetLineRate(index);
			if (1D == rate) ret = Color.LightPink;
			else if (0.9 <= rate) ret = Color.Yellow;
			else if (0D < rate) ret = Color.LightGreen;
			return ret;
		}

		private void getLineIndex(List<string> ls, string text, string prevText, ref int index, ref int pos, ref int lastIndex)
		{
			if (index < 65) Debug.WriteLine("### getLineIndex");
			int start = index;
			if (index <= lastIndex && !(0 < index && text.Equals(prevText) && !ls[index].Equals(ls[index - 1]))) start = lastIndex + 1;
			Debug.WriteLine(string.Format("start={0}\n text={1}\n prevText={2}", start, text, prevText));
			text = text.Trim().Trim('.');
			if (20 < text.Length) text = text.Substring(0, 20);
			//if (index <= lastIndex) index = lastIndex + 1;
			//int start = (index <= lastIndex) ? lastIndex + 1 : index;
			for (int i = start; i < ls.Count; i++)
			{
				string line = ls[i];
				line = line.Trim().Trim('.');
				if (string.IsNullOrEmpty(line))
				{
					index++;
					pos = 0;
					continue;
				}
				if (index < 65) Debug.WriteLine(string.Format("=>i loop i={0} index={1} lastIndex={2}:{3}\n{4}", i, index, lastIndex, line, text));
				string[] parts = line.Split(new Char[] { '.' });
				for (int j = pos; j < parts.Length; j++)
				{
					string part = parts[j].Trim();
					if (index < 65) Debug.WriteLine(string.Format("=>j loop j={0} pos={1} position={2}:{3}\n{4}", j, pos, part.IndexOf(text), part, text));
					if (0 <= part.IndexOf(text))
					{
						//if (index < i) pos = 0;
						//else pos++;
						index = i;
						if (j == parts.Length - 1)
						{
							pos = 0;
							lastIndex = index;
						}
						else pos = j + 1;
						//if (60 == i || 61 == i) Debug.WriteLine("Matched");
						return;
					}
				}
				//Debug.WriteLine("next");
				index++;
			}
		}

		private double getTargetLineRate(int targetLine)
		{
			foreach (KeyValuePair<int, MatchLine> pair in this.matchLines)
			{
				MatchLine ml = pair.Value;
				if (ml.TargetLine == targetLine) return ml.Rate;
			}
			return 0D;
		}

		private string ReplaceLine(string line)
		{
			line = Regex.Replace(line, @"[\x00-\x1F\x7F]", "");
			line = Regex.Replace(line, @"[\u00a0\uc2a0]", " "); //文字コードC2A0（UTF-8の半角空白）
			line = Regex.Replace(line, @"[\u0091\u0092\u2018\u2019]", "'"); //UTF-8のシングルクォーテーション
			line = Regex.Replace(line, @"[\u0093\u0094\u00AB\u201C\u201D]", "\""); //UTF-8のダブルクォーテーション
			line = Regex.Replace(line, @"[\u0097\u2013\u2014]", "\""); //UTF-8のハイフン
			line = Regex.Replace(line, @"[\u00A9\u00AE\u2022\u2122]", "\""); //UTF-8のスラッシュ
			//スペースに挟まれた「Fig.」で次が大文字でない場合は、文末と混同しないようにドットの後ろのスペースを削除する。
			line = Regex.Replace(line, @"(^fig| fig)\. +([^A-Z])", "$1.$2", RegexOptions.IgnoreCase);
			//2個以上連続するスペースは1個の半角スペースにする。
			line = Regex.Replace(line, @"\s+", " ");
			line = TextConverter.ZenToHan(line);
			line = TextConverter.HankToZen(line);
			return line;
		}
		#endregion
	}
}
