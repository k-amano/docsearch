using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Xyn.Util;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using Arx.DocSearch.Util;

namespace Arx.DocSearch.Client
{
	public partial class CompareForm : Form
	{
		public CompareForm(MainForm mainForm, ListViewItem lvi, string srcFile, bool isJp, int lineCount, int charCount, Dictionary<int, MatchLine> matchLines)
		{
			this.mainForm = mainForm;
			this.fname = lvi.SubItems[2].Text;
			this.srcFile = srcFile;
			this.isJp = isJp;
			this.matchLines = matchLines;
			InitializeComponent();
			StringBuilder sb = new StringBuilder();
			sb.Append("ファイル名：" + lvi.SubItems[2].Text + "\n");
			sb.Append("一致文数　：" + lvi.SubItems[1].Text + "　");
			sb.Append(string.Format("総文数　　：{0}　", lineCount));
			sb.Append("一致率　　：" + lvi.SubItems[0].Text + "%　");
			sb.Append(string.Format("総文字数　：{0}\n", charCount));
			this.infoLabel.Text = sb.ToString();
			this.lsSrc = new List<string>();
			this.lsTarget = new List<string>();
			this.messageLabel.Text = string.Empty;
		}
		private MainForm mainForm;
		private string fname;
		private string srcFile;
		private Dictionary<int, MatchLine> matchLines;
		List<string> lsSrc;
		List<string> lsTarget;
		bool isJp;

		private void CompareForm_Load(object sender, EventArgs e)
		{
			this.mainForm.Hide();
			this.splitContainer1.SplitterDistance = (this.splitContainer1.ClientRectangle.Width - this.splitContainer1.SplitterWidth) / 2;
			this.GetSrcText();
			this.GetTargetText();
			this.RecalculateRate();
			this.SetSrcText();
			this.SetTargetText();
			//フォルダ選択ダイアログ上部に表示する説明テキストを指定する
			this.folderBrowserDialog1.Description = "変換したWord文書の保存先フォルダを指定してください。";
		}

		private void CompareForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.mainForm.Show();
		}

		//splitContainer1の分割サイズを左右均等にする。
		private void CompareForm_SizeChanged(object sender, EventArgs e)
		{
			int distance = (this.splitContainer1.ClientRectangle.Width - this.splitContainer1.SplitterWidth) / 2;
			// 最小化したときはのdistanceがマイナスになり実行時エラーとなる。そもそも最小化時はサイズ変更不要であるので実行しない。
			if (10 < distance)
				this.splitContainer1.SplitterDistance = distance;
		}

		private void closeButton_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void macthCountButton_Click(object sender, EventArgs e)
		{
			MatchCountForm f = new MatchCountForm(this.mainForm, this.fname, this.srcFile, this.matchLines, this.lsSrc, this.lsTarget, this.isJp);
			f.Show();
		}

		private void convertButton_Click(object sender, EventArgs e)
		{
			string seletedPath = "";
			if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
			{
				seletedPath = this.folderBrowserDialog1.SelectedPath;
			}
			else return;
			this.messageLabel.Text = "Word変換中です。しばらくお待ちください。";
			var task = Task.Factory.StartNew(() =>
			{
				//this.convertButton.Enabled = false;
				WordConverter wc = new WordConverter(this.srcFile, this.lsSrc, this.fname, this.lsTarget, this.matchLines, seletedPath);
				wc.Run();
			});

			var continueTask = task.ContinueWith((t) =>
			{
				this.Invoke((MethodInvoker)delegate()
				{
					//this.convertButton.Enabled = true;
					this.messageLabel.Text = "Word変換が完了しました。";
				}
				);
			});
		}

		private void コピーCToolStripMenuItem_Click(object sender, EventArgs e)
		{
			this.srcText.Copy();
		}

		private void 検索SToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//クリップボードに文字列データがあるか確認
			if (Clipboard.ContainsText())
			{
				//文字列データがあるときはこれを取得する
				//取得できないときは空の文字列（String.Empty）を返す
				String str = Clipboard.GetText();
				//文字列を検索する
				int pos = this.srcText.Find(str);
				//指定文字列が見つかったか？
				if (-1 < pos)
				{
					//見つかった位置から、文字数分を選択
					this.srcText.Select(pos, str.Length);
					//フォーカスを当てる（フォーカスがはずれると選択が無効になってしまうため）
					this.srcText.Focus();
				}
			}
		}

		private void コピーCToolStripMenuItem1_Click(object sender, EventArgs e)
		{
			this.targetText.Copy();
		}

		private void 検索SToolStripMenuItem1_Click(object sender, EventArgs e)
		{
			//クリップボードに文字列データがあるか確認
			if (Clipboard.ContainsText())
			{
				//文字列データがあるときはこれを取得する
				//取得できないときは空の文字列（String.Empty）を返す
				String str = Clipboard.GetText();
				//文字列を検索する
				int pos = this.targetText.Find(str);

				//指定文字列が見つかったか？
				if (-1 < pos)
				{
					//見つかった位置から、文字数分を選択
					this.targetText.Select(pos, str.Length);
					//フォーカスを当てる（フォーカスがはずれると選択が無効になってしまうため）
					this.targetText.Focus();
				}
			}
		}

		private void GetTargetText()
		{
			string targetTextFile = SearchJob.GetTextFileName(this.fname);
			if (!File.Exists(targetTextFile))
			{
				MessageBox.Show(targetTextFile + "が存在しません。");
				return;
			}
			string line;
			this.lsTarget.Clear();
			using (StreamReader file = new StreamReader(targetTextFile))
			{
				for (int i = 0; (line = file.ReadLine()) != null; i++)
				{
					//line = Regex.Replace(line, @" ([.,:;]) ", "$1"); //半角句読点の前後スペースを削除。
					this.lsTarget.Add(line);
				}
			}
		}

		private void SetTargetText()
		{
			for (int i = 0; i < this.lsTarget.Count; i++)
			{
				string line = this.lsTarget[i];
				double rate = getTargetLineRate(i);
				if (1D == rate)
					this.targetText.SelectionBackColor = Color.LightPink;
				else if (0.9D <= rate)
					this.targetText.SelectionBackColor = Color.Yellow;
				else if (0D < rate)
					this.targetText.SelectionBackColor = Color.LightGreen;
				else this.targetText.SelectionBackColor = Color.White;
				this.targetText.AppendText(this.AddLineFeed(line));
			}

		}

		private void GetSrcText()
		{
			string srcTextFile = SearchJob.GetTextFileName(this.srcFile);
			if (!File.Exists(srcTextFile))
			{
				MessageBox.Show(srcTextFile + "が存在しません。");
				return;
			}
			string line;
			this.lsSrc.Clear();
			using (StreamReader file = new StreamReader(srcTextFile))
			{
				for (int i = 0; (line = file.ReadLine()) != null; i++)
				{
					//line = Regex.Replace(line, @" ([.,:;]) ", "$1"); //半角句読点の前後スペースを削除。
					this.lsSrc.Add(line);
				}
			}
		}

		private void SetSrcText()
		{
			for (int i = 0; i < this.lsSrc.Count; i++)
			{
				string line = this.lsSrc[i];
				if (this.matchLines.ContainsKey(i))
				{
					if (1D == this.matchLines[i].Rate)
						this.srcText.SelectionBackColor = Color.LightPink;
					else if (0.9 <= this.matchLines[i].Rate)
						this.srcText.SelectionBackColor = Color.Yellow;
					else if (0D < this.matchLines[i].Rate)
						this.srcText.SelectionBackColor = Color.LightGreen;
					else this.srcText.SelectionBackColor = Color.White;
				}
				else this.srcText.SelectionBackColor = Color.White;
				this.srcText.AppendText(this.AddLineFeed(line));
			}
		}

		private string AddLineFeed(string line)
		{
			int len = line.Length;
			if (3 < len)
			{
				string last3 = line.Substring(len - 3, 3);
				if (".  ".Equals(last3)) return line;
				if (Regex.IsMatch(last3, @"^[\)\]>）＞〕】≫》。]  $")) return line.TrimEnd();
			}
			return line + "\n";
		}

		private void RecalculateRate()
		{
			if (!this.isJp) return;
			foreach (int i in new List<int>(this.matchLines.Keys))
			{
				MatchLine ml = this.matchLines[i];
				string lineSrc = "";
				if (i < lsSrc.Count) lineSrc = lsSrc[i];
				if (ml.TotalWords == ml.MatchWords)
				{
					ml.TotalWords = lineSrc.Length;
					ml.MatchWords = ml.TotalWords;
				}
				else if (0 < ml.Rate)
				{
					string lineTarget = "";
					int wordCount = 0;
					int matchCount = 0;
					if (ml.TargetLine < lsTarget.Count) lineTarget = lsTarget[ml.TargetLine];
					ml.Rate = CompareForm.GetDiffRate(lineSrc, lineTarget, ref wordCount, ref matchCount);
					ml.TotalWords = wordCount;
					ml.MatchWords = matchCount;
					//Debug.WriteLine(string.Format("rate={0} {1}", ml.Rate, lineSrc));
				}
				this.matchLines[i] = ml;
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

		public static double GetDiffRate(string src, string target, ref int wordCount, ref int matchCount)
		{
			var d = new Differ();
			var inlineBuilder = new InlineDiffBuilder(d);
			var result = d.CreateCharacterDiffs(CompareForm.TrimSeparator(src), CompareForm.TrimSeparator(target), true);
			int diffCount = 0;
			wordCount = result.PiecesOld.Length;
			foreach (var block in result.DiffBlocks)
			{
				diffCount += block.DeleteCountA + block.InsertCountB;
				//diffCount += block.DeleteCountA;
			}

			matchCount = wordCount - diffCount;
			if (matchCount < 0) matchCount = 0;
			double rate = (double)matchCount / (double)wordCount;
			Debug.WriteLine(string.Format("src={0}\ntarget={1}\n matchCount={2}", src, target, matchCount));
			return rate;
		}

		// 「。」の後ろの空白は文の区切りを判定するために追加されたものなので文字数をカウントするときは削除する。
		public static string TrimSeparator(string line)
		{
			int len = line.Length;
			if (3 < len)
			{
				string last3 = line.Substring(len - 3, 3);
				if (Regex.IsMatch(last3, @"^[\)\]>）＞〕】≫》。]  $")) return line.TrimEnd();
			}
			return line;
		}
	}
}
