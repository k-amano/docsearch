using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Xyn.Util;

namespace Arx.DocSearch.MultiCore
{
	public partial class MatchCountForm : Form
	{
		public MatchCountForm(MainForm mainForm, string fname, string srcFile, Dictionary<int, MatchLine> matchLines, List<string> lsSrc, List<string> lsTarget, bool isJp)
		{
			this.mainForm = mainForm;
			this.fname = fname;
			this.srcFile = srcFile;
			this.matchLines = matchLines;
			this.lsSrc = lsSrc;
			this.lsTarget = lsTarget;
			this.isJp = isJp;
			InitializeComponent();
			StringBuilder sb = new StringBuilder();
			sb.Append("検索元ファイル名：" + srcFile + "\n");
			sb.Append("検索先ファイル名：" + fname + "\n");
			this.infoLabel.Text = sb.ToString();
		}
		private MainForm mainForm;
		private string fname;
		private string srcFile;
		private Dictionary<int, MatchLine> matchLines;
		List<string> lsSrc;
		List<string> lsTarget;
		bool isJp;

		private void MatchCountForm_Load(object sender, EventArgs e)
		{
			this.InitializeListView();
			this.GetListViewContent();
		}

		private void MatchCountForm_FormClosing(object sender, FormClosingEventArgs e)
		{

		}

		private void closeButton_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void InitializeListView()
		{
			// ListViewコントロールのプロパティを設定
			this.listView1.FullRowSelect = true;
			this.listView1.GridLines = true;
			this.listView1.Sorting = SortOrder.None;
			this.listView1.View = View.Details;
			//ColumnClickイベントハンドラの追加
			//this.listView1.ColumnClick += new ColumnClickEventHandler(ListView1_ColumnClick);

			// 列（コラム）ヘッダの作成
			string unit = isJp ? "文字" : "ワード";
			string[] captions = new string[] { "検索元No.", "検索先No.", "一致率", "一致" + unit + "数", "総" + unit + "数", "内容" };
			int[] widths = new int[] { 70, 70, 70, 80, 80, 1800 };
			ColumnHeader[] colHeaders = new ColumnHeader[6];
			for (int i = 0; i < colHeaders.Length; i++)
			{
				colHeaders[i] = new ColumnHeader();
				colHeaders[i].Text = captions[i];
				colHeaders[i].Width = widths[i];
			}
			this.listView1.Columns.AddRange(colHeaders);
		}

		private void GetListViewContent()
		{
			Debug.WriteLine(string.Format("############ isJp={0}", isJp));
			MatchLine ml100 = new MatchLine();
			MatchLine ml90 = new MatchLine();
			MatchLine mlOthers = new MatchLine();
			int i = 0;
			foreach (KeyValuePair<int, MatchLine> pair in this.matchLines)
			{
				MatchLine ml = pair.Value;
				string lineSrc = "";
				if (pair.Key < this.lsSrc.Count) lineSrc = this.lsSrc[pair.Key];
				this.listView1.Items.Add(new ListViewItem(new string[] { 
								string.Format("{0}", pair.Key + 1),
								string.Format("{0}", ml.TargetLine + 1),
								string.Format("{0:0.00}", ml.Rate * 100),
								string.Format("{0}", ml.MatchWords),
								string.Format("{0}", ml.TotalWords),
								this.GetHeadText(lineSrc) }));
				if (1D == ml.Rate)
				{
					this.listView1.Items[i].BackColor = Color.LightPink;
					ml100.MatchWords += ml.MatchWords;
					ml100.TotalWords += ml.TotalWords;
				}
				else if (0.9D <= ml.Rate)
				{
					this.listView1.Items[i].BackColor = Color.Yellow;
					ml90.MatchWords += ml.MatchWords;
					ml90.TotalWords += ml.TotalWords;
				}
				else if (0D < ml.Rate)
				{
					this.listView1.Items[i].BackColor = Color.LightGreen;
					mlOthers.MatchWords += ml.MatchWords;
					mlOthers.TotalWords += ml.TotalWords;
				}
				i++;
			}
			this.listView1.Items.Add(new ListViewItem(new string[] { 
								string.Empty,
								string.Empty,
								"100%計",
								string.Format("{0}", ml100.MatchWords),
								string.Format("{0}", ml100.TotalWords),
								string.Empty }));
			this.listView1.Items.Add(new ListViewItem(new string[] { 
								string.Empty,
								string.Empty,
								"90%以上計",
								string.Format("{0}", ml90.MatchWords),
								string.Format("{0}", ml90.TotalWords),
								string.Empty }));
			this.listView1.Items.Add(new ListViewItem(new string[] { 
								string.Empty,
								string.Empty,
								"90%未満計",
								string.Format("{0}", mlOthers.MatchWords),
								string.Format("{0}", mlOthers.TotalWords),
								string.Empty }));
		}

		private string GetHeadText(string src)
		{
			int len = src.Length;
			if (500 < len) len = 500;
			return src.Substring(0, len);
		}
	}
}

