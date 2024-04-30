using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Arx.DocSearch.Util;

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
        #endregion

        #region メソッド
        public void Run()
		{
			this.GetSrcParagraphs();
			this.GetTargetParagraphs();
			//this.DumpParagraphs(targetParagraphs, true);
			//return;
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
				this.EditWord(word, targetFile, true);
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
			int index = -1;
			int nextIndex = -1;
			int paragraphPos = -1;
			int nextPos = 0;
			int charPos = -1;
			this.mainForm.WriteLog(string.Format("WordConverter.EditWord: path={0} path2={1}", path, path2));
			try
			{
                //throw new Exception("An exception occurs.");
				doc = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
				string text = string.Empty;
				this.mainForm.WriteLog(string.Format("doc.Paragraphs.Count={0}", doc.Paragraphs.Count));
				for (int i = 0; i < doc.Paragraphs.Count; i++)
				{
					Paragraph para = doc.Paragraphs[i + 1];
					Range r = para.Range;
					try
					{
						if (0 == (r?.Text?.Trim().Length ?? 0)) continue;
					}
					catch (Exception e)
					{ 
						this.mainForm.WriteLog(string.Format("path={0} i={1}:{2}\n{3}",path, i,e.Message,e.StackTrace));
						throw new Exception("An exception occurs.");
					}
					bool isDebug = false;
					/*if (i < 10)
					{
						string str = r.Text;
						if (80 < str.Length) str = str.Substring(0, 80);
						this.mainForm.WriteLog(string.Format("#### EditWord i={0} paragraphPos={1} r.Text.Length={2}:{3}", i, paragraphPos, r.Text.Trim().Length, str));
						isDebug = true;
					}*/
					/*if (119 == i) isDebug = true;
					else isDebug = false;*/
					//if (isDebug) Debug.WriteLine(string.Format("i={0} r.Text={1}#", i, r.Text));
					this.FindParagraph(r.Text, ref paragraphPos, ref nextPos, isTarget, isDebug);
					//if (65 < i && i < 70) Debug.WriteLine(string.Format("paragraphPos={0} nextPos={1}#", paragraphPos, nextPos));
					List<string> sentences = new List<string>();
					StringBuilder sb = new StringBuilder();
					for (int j = 0; j < r.Sentences.Count; j++)
					{
						string str = r.Sentences[j + 1].Text;
						//if (isDebug) Debug.WriteLine(string.Format("count={0} text={1}#", j, str));
						if (0 < j)
						{
							if (this.StartsWithCapital(str))
							{
								sentences.Add(sb.ToString());
								sb = new StringBuilder(); ;
							}
							else
							{
								sentences.Add(string.Empty);
							}
						}
						sb.Append(str);
					}
					sentences.Add(sb.ToString());
					Color[] colors = new Color[sentences.Count];
					Color color;
					for (int j = 0; j < sentences.Count; j++)
					{
						color = Color.White;
						if (!string.IsNullOrEmpty(sentences[j]))
						{
							string[] lines = this.ReplaceLine(sentences[j]);
							text = lines[lines.Length - 1].Trim();
							if (isTarget) color = this.GetTargetColor(text, paragraphPos, ref index, ref charPos, ref nextIndex);
							else color = this.GetSrcColor(text, paragraphPos, ref index, ref charPos, ref nextIndex);
							//if (isDebug) Debug.WriteLine(string.Format("color={0} text={1}#", color, text));
						}
						colors[j] = color;
					}
					//センテンスの長さが0の場合は後ろの色とするため、終わりから色を振り直す。
					int id = 0;
					color = Color.White;
					for (int j = 0; j < sentences.Count; j++)
					{
						id = sentences.Count - 1 - j;
						if (!string.IsNullOrEmpty(sentences[id]))
						{
							color = colors[id];
						}
						colors[id] = color;
						//if (isDebug) Debug.WriteLine(string.Format("id={0} color={1}#", id, colors[id]));
					}
					for (int j = 0; j < r.Sentences.Count; j++)
					{
						Range r2 = r.Sentences[j + 1];
						r2.Font.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(colors[j]);
						r2.Select();
					}
				}
				doc.SaveAs(ref path2, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
                this.mainForm.WriteLog("WordConverter.EditWord:" + e.Message + "\n"+  e.StackTrace);
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

		private Color GetSrcColor(string text, int paragraphPos, ref int index, ref int charPos, ref int nextIndex)
		{
			Color ret = Color.White;
			text = text.Trim();
			if (paragraphPos < 0 || string.IsNullOrEmpty(text)) return ret;
			List<int> paragraph = this.srcParagraphs[paragraphPos];
			this.getLineIndex(this.lsSrc, text, paragraph, ref index, ref charPos, ref nextIndex);
			if (this.lsSrc.Count <= index) return ret;
			if (this.matchLines.ContainsKey(index))
			{
				if (1D == this.matchLines[index].Rate) ret = Color.LightPink;
				else if (0.9 <= this.matchLines[index].Rate) ret = Color.Yellow;
				else if (0D < this.matchLines[index].Rate) ret = Color.LightGreen;
			}
			return ret;
		}

		private Color GetTargetColor(string text, int paragraphPos, ref int index, ref int charPos, ref int nextIndex)
		{
			Color ret = Color.White;
			text = text.Trim();
			if (paragraphPos < 0 || string.IsNullOrEmpty(text)) return ret;
			List<int> paragraph = this.targetParagraphs[paragraphPos];
			this.getLineIndex(this.lsTarget, text, paragraph, ref index, ref charPos, ref nextIndex);
			if (this.lsTarget.Count <= index) return ret;
			double rate = getTargetLineRate(index);
			if (1D == rate) ret = Color.LightPink;
			else if (0.9 <= rate) ret = Color.Yellow;
			else if (0D < rate) ret = Color.LightGreen;
			return ret;
		}

		private void getLineIndex(List<string> ls, string text, List<int> paragraph, ref int index, ref int charPos, ref int nextIndex)
		{
			string replaced = text.Trim().Trim('.');
			if (120 < replaced.Length) replaced = replaced.Substring(0, 120);
			for (int i = 0; i < paragraph.Count; i++)
			{
				if (paragraph[i] < nextIndex) continue;
				string line = ls[paragraph[i]];
				line = line.Trim().Trim('.');
				if (0 < replaced.Length && 0 < line.Length)
				{
					int start = charPos < 0 ? 0 : line.Length <= charPos ? line.Length - 1 : charPos;
					int newPos = line.IndexOf(replaced, start);
					if (0 <= newPos)
					{
						index = paragraph[i];
						charPos = newPos + text.Length;
						if (line.Length <= charPos)
						{
							nextIndex = index + 1;
							charPos = -1;
						}
						else
						{
							nextIndex = index;
						}
						return;
					}
				}
				index++;
				charPos = -1;
				nextIndex = index;
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

		private string[] ReplaceLine(string line)
		{
			line = Regex.Replace(line ?? "", @"[\x00-\x1F\x7F]", "");
			line = Regex.Replace(line ?? "", @"[\u00a0\uc2a0]", " "); //文字コードC2A0（UTF-8の半角空白）
			line = Regex.Replace(line ?? "", @"[\u0091\u0092\u2018\u2019]", "'"); //UTF-8のシングルクォーテーション
			line = Regex.Replace(line ?? "", @"[\u0093\u0094\u00AB\u201C\u201D]", "\""); //UTF-8のダブルクォーテーション
			line = Regex.Replace(line ?? "", @"[\u0097\u2013\u2014]", "\""); //UTF-8のハイフン
			line = Regex.Replace(line ?? "", @"[\u00A9\u00AE\u2022\u2122]", "\""); //UTF-8のスラッシュ
			//スペースに挟まれた「Fig.」で次が大文字でない場合は、文末と混同しないようにドットの後ろのスペースを削除する。
			line = Regex.Replace(line ?? "", @"(^fig| fig)\. +([^A-Z])", "$1.$2", RegexOptions.IgnoreCase);
			//2個以上連続するスペースは1個の半角スペースにする。
			line = Regex.Replace(line ?? "", @"\s+", " ");
			line = TextConverter.ZenToHan(line);
			line = TextConverter.HankToZen(line);
			line = Regex.Replace(line ?? "", @"^((([\(\[<（＜〔【≪《])([^0-9]*[0-9]*)([\)\]>）＞〕】≫》])(\s*))+)", "\n$1  \n"); //【数字】
			line = Regex.Replace(line ?? "", @"^([0-9]+)(\.?)", "\n$1$2  \n"); //数字
			string[] lines = line.Split('\n');
			return lines;
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

		private void FindParagraph(string para, ref int paragraphPos, ref int nextPos, bool isTarget, bool isDebug)
		{
			List<List<int>> paragraphs = isTarget ? this.targetParagraphs : this.srcParagraphs;
			string[] lines = this.ReplaceLine(para);
			foreach (string line in lines)
			{
				string replaced = line.Trim();
				if (0 == replaced.Length) continue;
				if (120 < replaced.Length) replaced = replaced.Substring(0, 120);
				for (int i = nextPos; i < paragraphs.Count; i++)
				{
					string text = this.GetParagraphText(paragraphs[i], isTarget);
					//if (isDebug) Debug.WriteLine(string.Format("FindParagraph i={0}:text={1}#", i, text));
					if (0 < text.Trim().Length && 0 <= text.IndexOf(replaced))
					{
						paragraphPos = i;
						if (para.Trim().EndsWith(".")) nextPos = paragraphPos + 1;
						//if (isDebug) Debug.WriteLine(string.Format("pos={0} replaced={1}#", text.IndexOf(replaced), replaced));
						return;
					}
					//if (isDebug) Debug.WriteLine(string.Format("pos={0} replaced={1}#", text.IndexOf(replaced), replaced));
				}
			}
		}

		private string GetParagraphText(List<int> paragraph, bool isTarget)
		{
			StringBuilder sb = new StringBuilder();
			List<string> ls = isTarget ? this.lsTarget : this.lsSrc;
			foreach (int i in paragraph)
			{
				string line = ls[i];
				sb.Append(line);
			}
			//2個以上連続するスペースは1個の半角スペースにする。
			return Regex.Replace(sb.ToString(), @"\s+", " ");
		}

		private void DumpParagraphs(List<List<int>> paragraphs, bool isTarget)
		{
			for (int i = 0; i < paragraphs.Count; i++)
			{
				string text = this.GetParagraphText(paragraphs[i], isTarget);
				Debug.WriteLine(string.Format("paragrah={0} text={1}", i, text));
			}
		}

		private bool StartsWithCapital(string line)
		{
			if (Regex.IsMatch(line ?? "", @"^(\s*[A-Z])")) return true;
			else return false;
		}
		#endregion
	}
}

