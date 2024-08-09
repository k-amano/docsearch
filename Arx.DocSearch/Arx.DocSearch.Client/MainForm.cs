using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xyn.Util;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Arx.DocSearch.Util;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using Microsoft.Office.Interop.Word;
using Application = System.Windows.Forms.Application;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using Task = System.Threading.Tasks.Task;
using View = System.Windows.Forms.View;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;
using static Arx.DocSearch.Client.NodeManager;

namespace Arx.DocSearch.Client
{
	public partial class MainForm : Form
	{
		[DllImport("kernel32.dll")]
		private static extern bool AllocConsole();
		[DllImport("xd2txlib.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Cdecl)]

		public static extern int ExtractText(
				[MarshalAs(UnmanagedType.BStr)] String lpFileName,
				bool bProp,
				[MarshalAs(UnmanagedType.BStr)] ref String lpFileText);

		#region コンストラクタ
		public MainForm()
		{
			this.logs = new List<string>();
			InitializeComponent();
			this.matchLinesTable = new Dictionary<int, Dictionary<int, MatchLine>>();
			this.reservationList = new List<Reservation>();
			this.timer1.Interval = 5000;
			this.isProgressing = false;
			this.word = null;
			// Console表示
			//AllocConsole();
			// コンソールとstdoutの紐づけを行う。無くても初回は出力できるが、表示、非表示を繰り返すとエラーになる。
			//Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
		}
		#endregion

		#region フィールド
		private Schema config;
		private string configFile;
		//private readonly int LINE_LENGHTH = 70;
		private int minWords;
		private Dictionary<int, Dictionary<int, MatchLine>> matchLinesTable;
		private int lineCount = 0;
		private int charCount = 0;
		private bool isJp = false;
		private List<Reservation> reservationList;
		private SearchJob job;
		private string xlsdir;
		private List<string> logs;
		private bool isProgressing;
        private List<Process> agentPrograms;
		private int srcIndex;
		private WordApplication word;
		#endregion

		#region Property
		public string UserAppDataPath
		{
			get
			{
				return FileEx.GetFileSystemPath(Environment.SpecialFolder.ApplicationData);
			}
		}
		/// <summary>
		/// ログイン時にアクセスする Web ページの URL を取得または設定します。
		/// </summary>
		public string SrcFile
		{
			get
			{
				return srcCombo.Text;
			}
			set
			{
				srcCombo.Text = value;
			}
		}

		public List<string> SrcFiles
		{
			get
			{
				return this.GetSrcFiles();
			}
			set
			{
				this.SetSrcFiles(value);
			}
		}

		public int MinWords
		{
			get
			{
				return minWords;
			}
			set
			{
				minWords = value;
			}
		}

		public string WordCount
		{
			get
			{
				return wordCountText.Text;
			}
			set
			{
				wordCountText.Text = value;
			}
		}

		public string RoughLines
		{
			get
			{
				return this.roughLinesText.Text;
			}
			set
			{
				roughLinesText.Text = value;
			}
		}

		public Dictionary<int, Dictionary<int, MatchLine>> MatchLinesTable
		{
			get
			{
				return matchLinesTable;
			}
		}
		#endregion

		/// <summary>
		/// ListViewの項目の並び替えに使用するクラス
		/// </summary>
		public class ListViewItemComparer : IComparer
		{
			private int _column;

			/// <summary>
			/// ListViewItemComparerクラスのコンストラクタ
			/// </summary>
			/// <param name="col">並び替える列番号</param>
			public ListViewItemComparer(int col)
			{
				_column = col;
			}

			//xがyより小さいときはマイナスの数、大きいときはプラスの数、
			//同じときは0を返す
			public int Compare(object x, object y)
			{
				//ListViewItemの取得
				ListViewItem itemx = (ListViewItem)x;
				ListViewItem itemy = (ListViewItem)y;
				if (0 == _column)
					//xとyをDoubleとして比較する(降順)
					return ConvertEx.GetDouble(itemy.SubItems[_column].Text).CompareTo(ConvertEx.GetDouble(itemx.SubItems[_column].Text));
				else if (1 == _column)
					//xとyを整数として比較する(降順)
					return ConvertEx.GetInt(itemy.SubItems[_column].Text).CompareTo(ConvertEx.GetInt(itemx.SubItems[_column].Text));
				else  if (3 == _column)
					//xとyを整数として比較する(昇順)
					return ConvertEx.GetInt(itemx.SubItems[_column].Text).CompareTo(ConvertEx.GetInt(itemy.SubItems[_column].Text));
				//xとyを文字列として比較する
				else return string.Compare(itemx.SubItems[_column].Text,
						itemy.SubItems[_column].Text);
			}
		}

		#region メソッド
		private void MainForm_Load(object sender, EventArgs e)
		{
			//OpenFileDialog
			this.openFileDialog1 = new OpenFileDialog();
			this.openFileDialog1.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
			this.openFileDialog1.FileName = "";
			//[ファイルの種類]に表示される選択肢を指定する
			//指定しないとすべてのファイルが表示される
			this.openFileDialog1.Filter = "Word文書(*.doc;*.docx)|*.doc;*.docx|pdf文書(*.pdf)|*.pdf|テキストファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*";
			//タイトルを設定する
			this.openFileDialog2.Title = "読み込むログファイルを選択してください";
			//OpenFileDialog
			this.openFileDialog2 = new OpenFileDialog();
			this.openFileDialog2.InitialDirectory = this.UserAppDataPath;
			this.openFileDialog2.FileName = "";
			//[ファイルの種類]に表示される選択肢を指定する
			//指定しないとすべてのファイルが表示される
			this.openFileDialog2.Filter = "ログ(*.log)|*.log|すべてのファイル(*.*)|*.*";
			//タイトルを設定する
			this.openFileDialog1.Title = "検索元のテキストファイルを選択してください";
			//フォルダ選択ダイアログ上部に表示する説明テキストを指定する
			this.folderBrowserDialog1.Description = "検索先のフォルダを指定してください。";
			this.folderBrowserDialog2.Description = "検索結果を出力するフォルダを指定してください。";
			this.InitializeListView();
			this.configFile = Path.Combine(this.UserAppDataPath, "DocSearch.config");
			this.config = Schema.LoadSettings(this.configFile);
			//this.srcCombo.Text = this.config.SrcFile;
			this.SrcFiles = this.config.SrcFiles;
			this.targetText.Text = this.config.TargetFolder;
			this.rateText.Text = this.config.Rate;
			this.wordCountText.Text = this.config.WordCount;
			this.roughLinesText.Text = this.config.RoughLines;
			this.xlsdir = this.config.Xlsdir;
			this.srcIndex = ConvertEx.GetInt(this.config.SrcIndex);
			this.messageLabel.Text = string.Empty;
			this.countLabel.Text = string.Empty;
			this.GetTotalCount();
			this.timer1.Start();
			//this.job = new SearchJob(this);
			//Thread.Sleep(5000);
			//this.StartNodeManager();
			if (0 < this.srcIndex)
			{
                this.AddReservation();
                this.SearchReservationList();
            }

        }

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			this.WriteLog(string.Format("Closing MainForm. Xlsdir={0}", this.xlsdir));
			if (null != this.config)
			{
				//this.config.SrcFile = this.srcCombo.Text;
				this.config.SrcFiles = this.SrcFiles;
				this.config.TargetFolder = this.targetText.Text;
				this.config.Rate = this.rateText.Text;
				this.config.WordCount = this.wordCountText.Text;
				this.config.RoughLines = this.roughLinesText.Text;
				this.config.Xlsdir = this.xlsdir;
				this.config.SrcIndex = ConvertEx.GetString(this.srcIndex);	
				this.config.SaveSettings(this.configFile);
				this.WriteLog(string.Format("Conifg File was saved.  Xlsdir={0}", this.config.Xlsdir));
			}
			//this.job.Dispose();
			this.WriteErrorLog();
		}

		private void srcButton_Click(object sender, EventArgs e)
		{
			SelectSourceForm f = new SelectSourceForm(this);
			f.ShowDialog();
			/*if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				this.srcText.Text = this.openFileDialog1.FileName;
			}*/
			this.GetTotalCount();
		}

		private void targetButton_Click(object sender, EventArgs e)
		{
			if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
			{
				this.targetText.Text = this.folderBrowserDialog1.SelectedPath;
			}
		}

		private void searchButton_Click(object sender, EventArgs e)
		{
			this.isJp = false;
			this.SelectAction();
		}

		private void searchJpButton_Click(object sender, EventArgs e)
		{
			this.isJp = true;
			this.SelectAction();
		}

		private void compareButton_Click(object sender, EventArgs e)
		{
			if (0 == this.listView1.SelectedItems.Count)
			{
				MessageBox.Show("比較する行を選択してください。");
				return;
			}
			ListViewItem lvi = this.listView1.SelectedItems[0];
			int docId = ConvertEx.GetInt(lvi.SubItems[3].Text);
			Dictionary<int, MatchLine> matchLines = new Dictionary<int, MatchLine>();
			if (this.matchLinesTable.ContainsKey(docId)) matchLines = this.matchLinesTable[docId];
			CompareForm f = new CompareForm(this, lvi, this.srcCombo.Text, this.isJp, this.lineCount, this.charCount, matchLines);
			f.Show();
		}

		private void indexButton_Click(object sender, EventArgs e)
		{
			foreach (string srcFile in this.srcCombo.Items)
			{
				if (!File.Exists(srcFile))
				{
					MessageBox.Show("検索元が選択されていないか、存在しないファイルが含まれています。");
					return;
				}
			}
			var task = Task.Factory.StartNew(() =>
			{
				foreach (string srcFile in this.srcCombo.Items)
				{
					string textFile = SearchJob.GetTextFileName(srcFile);
					if (string.IsNullOrEmpty(textFile))
					{
						List<string> ls = new List<string>();
						ls.Add(srcFile);
						if (".pdf".Equals(Path.GetExtension(srcFile).ToLower())) this.MakeTextFileFromPdf(ls);
						else this.MakeTextFile(ls);
						textFile = SearchJob.GetTextFileName(srcFile);
					}
					this.MakeIndexFile(textFile);
				}
				this.Invoke((MethodInvoker)delegate()
				{
					this.messageLabel.Text = "インデックス作成が完了しました。";
				});
			});
		}

		private void textFileButton_Click(object sender, EventArgs e)
		{
			List<string> docs = this.FindDocuments();
			List<string> docFiles = new List<string>();
			List<string> pdfFiles = new List<string>();
			foreach (string doc in docs)
			{
				if (".pdf".Equals(Path.GetExtension(doc).ToLower())) pdfFiles.Add(doc);
				else docFiles.Add(doc);
			}
			this.messageLabel.Text = "テキスト抽出中です。";
			var task = Task.Factory.StartNew(() =>
			{
				this.MakeTextFile(docFiles);
				this.MakeTextFileFromPdf(pdfFiles);
			});

			var continueTask = task.ContinueWith((t) =>
			{
				this.Invoke((MethodInvoker)delegate()
				{
					this.messageLabel.Text = "テキスト抽出が完了しました。";
				}
				);
			}); 
		}

		private void logButton_Click(object sender, EventArgs e)
		{
			if (this.openFileDialog2.ShowDialog() == DialogResult.OK)
			{
				this.LoadLog(this.openFileDialog2.FileName);
			}

		}

		//列がクリックされた時
		private void ListView1_ColumnClick(
				object sender, ColumnClickEventArgs e)
		{
			//ListViewItemSorterを指定する
			this.listView1.ListViewItemSorter =
					new ListViewItemComparer(e.Column);
			//並び替える（ListViewItemSorterを設定するとSortが自動的に呼び出される）
			this.listView1.Sort();

		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			this.WriteErrorLog();
			if (this.isProgressing && null != this.job) this.job.showProgress();
		}

		private void InitializeListView()
		{
			// ListViewコントロールのプロパティを設定
			this.listView1.FullRowSelect = true;
			this.listView1.GridLines = true;
			this.listView1.Sorting = SortOrder.None;
			this.listView1.View = View.Details;
			//ColumnClickイベントハンドラの追加
			this.listView1.ColumnClick +=
					new ColumnClickEventHandler(ListView1_ColumnClick);
			// 列（コラム）ヘッダの作成
			string[] captions = new string[] { "一致率", "一致文数", "ファイル名", "文章No." };
			int[] widths = new int[] { 60, 60, 450, 0 };
			ColumnHeader[] colHeaders = new ColumnHeader[4];
			for (int i = 0; i < colHeaders.Length; i++)
			{
				colHeaders[i] = new ColumnHeader();
				colHeaders[i].Text = captions[i];
				colHeaders[i].Width = widths[i];
			}
			this.listView1.Columns.AddRange(colHeaders);
		}

		public void updateListView(MatchDocument[] matchArray, bool clearItems)
		{
			this.Invoke(
				(MethodInvoker)delegate()
				{
					if (clearItems && 0 < matchArray.Length) this.listView1.Items.Clear();
					foreach (MatchDocument doc in matchArray)
					{
						string strRate = string.Format("{0:0.00}", doc.Rate * 100);
						this.listView1.Items.Add(new ListViewItem(new string[] { 
							strRate,
							string.Format("{0}", doc.MatchCount), doc.Doc,
							string.Format("{0}", doc.DocId) }));
					}
				}
			);
		}

		private List<string> FindDocuments()
		{
			List<string> exts = new List<string> { ".txt", ".doc", ".docx", ".pdf" };
			List<string> docs = new List<string>();
			if (!Directory.Exists(this.targetText.Text)) return docs;
			string[] files = Directory.GetFiles(this.targetText.Text, "*", SearchOption.AllDirectories);
			foreach (string file in files)
			{
				string dir = Path.GetDirectoryName(file);
				if (dir.Contains(".adsidx")) continue; //".adsidx"ディレトクトリはパスする
				// 「.」で始まるファイルはパスする
				if (Path.GetFileName(file).StartsWith(".") || Path.GetFileName(dir).StartsWith(".")) continue;
				if (exts.Contains(Path.GetExtension(file).ToLower()))
				{
					docs.Add(file);
				}
			}
			return docs;
		}

		private string GetTargetText(string doc)
		{
			string fname = SearchJob.GetTextFileName(doc);
			string line;
			StringBuilder sb  = new StringBuilder();
			if (File.Exists(fname))
			{
				using (StreamReader file = new StreamReader(fname))
				{
					while ((line = file.ReadLine()) != null)
					{
						//if (string.IsNullOrEmpty(line)) continue;
						sb.Append(line);
						sb.Append("\n");
					}
				}
			}
			return sb.ToString();
		}

		private void MakeTextFile(List<string> docFiles)
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
				}
				else return;
			}
			this.word = null;
			StreamWriter writer = null;
			try
			{
				this.word = new WordApplication();
				this.word.Visible = false;
				foreach (string srcFile in docFiles)
				{
					string extension = Path.GetExtension(srcFile);
					string path;
					if (extension.Equals(".doc", StringComparison.OrdinalIgnoreCase))
					{
						path = Path.ChangeExtension(srcFile, ".docx");
					}
					else
					{
						path = srcFile;
					}
					string dir = Path.Combine(Path.GetDirectoryName(path), ".adsidx");
					string textFile = Path.Combine(dir, Path.GetFileName(path) + ".txt");
					if (File.Exists(textFile)) continue;
					//if (File.Exists(textFile)) File.Delete(textFile);
					this.Invoke(
						(MethodInvoker)delegate()
						{
							this.messageLabel.Text = srcFile + "をテキスト抽出中です。";
						}
					);
					if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
					writer = File.CreateText(textFile);
					writer.NewLine = "\n";
					string fileText = "";
					if (extension.Equals(".doc", StringComparison.OrdinalIgnoreCase))
					{
						WordDocumentConverter.ConvertDocToDocx(this.word, srcFile, path);
					} else {
						path = srcFile;
					}
					//int l = ExtractText(srcFile, false, ref fileText);
					fileText = WordTextExtractor.ExtractText(path);

					string[] paragraphs = fileText.Split('\n');
					int maxContiuousNumber = 0;
					bool isContinuousNumber = false;
					bool excludesTable = false;
					int pos = 0;
					for (int i = 0; i < paragraphs.Length; i++)
					{
						string line = paragraphs[i].TrimEnd();
						if (pos <= i)
						{
							int continuousNumber = this.GetContinuousNumber(i, paragraphs, ref maxContiuousNumber, ref excludesTable);
							if (0 < continuousNumber)
							{
								pos = i + continuousNumber;
								isContinuousNumber = true;
							}
						}
						bool endsContinuousNumber = (0 < i) && (i == pos - 1);
						//if ( i<460) Console.WriteLine(string.Format("i={0}, pos={1}, endsContinuousNumber={2}", i, pos, endsContinuousNumber));
						line = this.ReplaceLine(line, isContinuousNumber, excludesTable, endsContinuousNumber);
						//if (i ==457) Console.WriteLine(string.Format("i={0} line={1} isContinuousNumber={2} excludesTable={3} endsContinuousNumber={4}", i, line, isContinuousNumber, excludesTable, endsContinuousNumber));
						string[] sentences = line.Split('.');
						bool startsWithCapital = false;
						StringBuilder sb = new StringBuilder();
						for (int j = 0; j < sentences.Length; j++)
						{
							string sentence = sentences[j];
							sb.Append(sentence.Trim());
							if (j < sentences.Length - 1) {
								startsWithCapital = false;
								if (this.StartsWithCapital(sentences[j + 1])) startsWithCapital = true;
								if (startsWithCapital)
								{
									sb.Append(".\n");//次が大文字で始まっていれば改行する
								}
								else if (!string.IsNullOrEmpty(sentence.Trim()))
								{
									sb.Append(". "); //それ以外は空白を追加。
								}
							}	
						}
						writer.Write(sb.ToString());
						if (line.Trim().EndsWith(".") && !sb.ToString().Trim().EndsWith(".")) writer.Write(". ");
						startsWithCapital = false;
						if (i + 1 < paragraphs.Length && this.StartsWithCapital(paragraphs[i + 1])) startsWithCapital = true;
						if (startsWithCapital || line.TrimEnd().EndsWith(".") || line.TrimEnd().EndsWith("。"))
						{
							writer.Write("\n");//ピリオドまたは読点で終わっていれば改行する
						}
						else if (!string.IsNullOrEmpty(line.Trim()))
						{
							writer.Write(" "); //それ以外は空白を追加。
						}

					}
					writer.Close();
				}
			}
			catch (Exception e) {
				Debug.WriteLine(e.StackTrace);
			}
			finally
			{
				if (null != writer)
				{
					writer.Close();
				}
				if (null != this.word)
				{
					((_Application)this.word).Quit();
					Marshal.ReleaseComObject(this.word);  // オブジェクト参照を解放
					this.word = null;
				}
			}
		}

		private void MakeTextFileFromPdf(List<string> pdfFiles)
		{
			foreach (string srcFile in pdfFiles)
			{
				string dir = Path.Combine(Path.GetDirectoryName(srcFile), ".adsidx");
				string textFile = Path.Combine(dir, Path.GetFileName(srcFile) + ".txt");
				this.Invoke(
					(MethodInvoker)delegate()
					{
						this.messageLabel.Text = srcFile + "をテキスト抽出中です。";
					}
				);
				if (File.Exists(textFile)) continue;
				//if (File.Exists(textFile)) File.Delete(textFile);
				if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
				PdfReader pdfReader = null;
				StreamWriter writer = null;
				ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
				int maxContiuousNumber = 0;
				bool excludesTable = false;
				try
				{
					pdfReader = new PdfReader(srcFile);
					writer = File.CreateText(textFile);
					writer.NewLine = "\n";
					for (int page = 1; page <= pdfReader.NumberOfPages; page++)
					{
						string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
						string[] lines = currentText.Split('\n');
						bool isContinuousNumber = false;
						int pos = 0;
						for (int i = 0; i < lines.Length; i++)
						{
							string line = lines[i].TrimEnd();
							//if (1 == page && i<1000) Console.WriteLine(string.Format("i={0}, line={1}", i, line));
							if (pos <= i) {
								int continuousNumber = this.GetContinuousNumber(i, lines, ref maxContiuousNumber, ref excludesTable);
								if (0 < continuousNumber) {
									pos = i + continuousNumber;
									isContinuousNumber = true;
								}
							}
							bool endsContinuousNumber = (0 < i) && (pos == i);
							line = this.ReplaceLine(line, isContinuousNumber, excludesTable, endsContinuousNumber);
							writer.Write(line);
							bool startsWithCapital = false;
							if (i + 1 < lines.Length && this.StartsWithCapital(lines[i + 1])) startsWithCapital = true;
							if (startsWithCapital || line.TrimEnd().EndsWith(".") || line.TrimEnd().EndsWith("。"))
							{

								writer.Write("\n");//ピリオドまたは読点で終わっていれば改行する
							}
							else if (!string.IsNullOrEmpty(line.Trim()))
							{
								writer.Write(" "); //それ以外は空白を追加。
							}
						}
					}
				}
				finally
				{
					if (null != writer)
					{
						writer.Close();
					}
					if (null != pdfReader)
					{
						pdfReader.Close();
					}
				}
			}
		}

		private string ReplaceLine(string line, bool isContinuousNumber, bool excludesTable, bool endsContinuousNumber)
		{
			line = Regex.Replace(line ?? "", @"\s*SEQ Paragraph\s+\\#\s+""?\[\d+\]""?\s+(\\\*\s*MERGEFORMAT\s*)+", "");//フィールドコードを削除
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
			//センテンスの終わりで「半角スペース2個」+改行とする。
			line = Regex.Replace(line ?? "", @"\. +", ".  \n"); //ピリオド+半角スペース1個は改行
			line = Regex.Replace(line ?? "", @"。([^\n])", "。  \n($1)"); //読点。
			line = TextConverter.ZenToHan(line ?? "");
			line = TextConverter.HankToZen(line ?? "");
			if (isContinuousNumber && Regex.IsMatch(line ?? "".Trim(), @"^[ 0-9+-]+$"))
			{
				if (excludesTable)
				{
					if (endsContinuousNumber) return "\n";
					return "";
				}
				else
				{
					if (endsContinuousNumber) line += "\n";
					return line;
				}
			}
			//line = Regex.Replace(line, @"^([\(\[<（＜〔【≪《])([^0-9]*[0-9]*)([\)\]>）＞〕】≫》])(\s*)", "\n$1$2$3$4  \n"); //【数字】
			line = Regex.Replace(line ?? "", @"^\s*((([\(\[<（＜〔【≪《])([^0-9]*[0-9]*)([\)\]>）＞〕】≫》])(\s*))+)", "\n$1  \n"); //【数字】
			line = Regex.Replace(line ?? "", @"^\s*([0-9]+)(\.?)", "\n$1$2  \n"); //数字
			//line = Regex.Replace(line, @"([.,:;])", " $1 "); //半角句読点は前後にスペース
			return line;
		}

		private bool StartsWithCapital(string line)
		{
			if (Regex.IsMatch(line, @"^(\s*[A-Z][^.])")) return true;
			else return false;
		}

		private int GetContinuousNumber(int i, string[] paragraphs, ref int maxContiuousNumber, ref bool excludesTable)
		{
			if (0 == i) return 0;
			string previousline = paragraphs[i].Trim();
			if (Regex.IsMatch(previousline, @"^[ 0-9+-]+$"))
			{
				int count = GetNumberCount(i, paragraphs, ref maxContiuousNumber, ref excludesTable);
				return count;
			}
			else return 0;
		}

		private int GetNumberCount(int i, string[] paragraphs, ref int maxContiuousNumber, ref bool excludesTable)
		{
			int count = 0;
			for (int j = i; j < paragraphs.Length; j++)
			{
				string line = paragraphs[j].Trim();
				if (string.IsNullOrEmpty(line) || Regex.IsMatch(line, @"^[ 0-9+-]+$"))
				{
					string[] arr = line.Split(' ');
					count += arr.Length;
				}
				else break;

			}
			if (maxContiuousNumber < count)
			{
				maxContiuousNumber = count;
				if (!excludesTable && 100 < count) {
					string message = string.Format(@"{0}件の数字が連続する数表があります。
数表を除外して検索しますか？", count);
					//メッセージボックスを表示する
					DialogResult result = MessageBox.Show(message,
						"処理の選択",
							MessageBoxButtons.YesNo,
							MessageBoxIcon.Exclamation,
							MessageBoxDefaultButton.Button2);

					//何が選択されたか調べる
					if (result == DialogResult.Yes)
					{
						//「はい」が選択された時
						excludesTable = true;
					}
					else if (result == DialogResult.No)
					{
						//「いいえ」が選択された時
						excludesTable = false;
					}
				}
			}
			return count;
		}

		private void MakeIndexFile(string textFile)
		{
			string dir = Path.GetDirectoryName(textFile);
			string indexFile = Path.Combine(dir, Path.GetFileNameWithoutExtension(textFile) + ".idx");
			if (File.Exists(indexFile)) File.Delete(indexFile);
			StreamReader reader = null;
			StreamWriter writer = null;
			try
			{
				writer = File.CreateText(indexFile);
				writer.NewLine = "\n";
				reader = new StreamReader(textFile, Encoding.UTF8);
				string line = "";
				while (null != (line = reader.ReadLine()))
				{
					string line2 = TextConverter.SplitWords(line);
					writer.WriteLine(line2);
				}
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.StackTrace);
			}
			finally
			{
				if (null != writer)
				{
					writer.Close();
				}
				if (null != reader)
				{
					reader.Close();
				}
			}
		}

		private void GetTotalCount()
		{
			this.lineCount = 0;
			this.charCount = 0;
			string textFile = SearchJob.GetTextFileName(this.srcCombo.Text);
			if (File.Exists(textFile))
			{
				using (StreamReader file = new StreamReader(textFile))
				{
					string line;
					while ((line = file.ReadLine()) != null)
					{
						if (string.IsNullOrEmpty(line)) continue;
						charCount += line.Length;
						lineCount++;
					}
				}
			}
			this.totalCountLabel.Text = string.Format("総文数{0:#,0}　総文字数{1:#,0}", this.lineCount, this.charCount);
		}


		private void SaveLog(string srcFile)
		{
			string fname = Path.GetFileNameWithoutExtension(srcFile);
			var timespan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
			string logFile = Path.Combine(this.UserAppDataPath, string.Format("{0}.{1}.log", fname, (uint)timespan.TotalSeconds));
			Log log = new Log();
			log.SrcFile = srcFile;
			log.TargetFolder = this.targetText.Text;
			log.ListItems = this.ListViewToCsv();
			log.IsJp = this.isJp;
			log.LineCount = this.lineCount;
			log.CharCount = this.charCount;
			log.MatchLinesTable = this.matchLinesTable;
			log.SaveSettings(logFile);
		}

		private void LoadLog(string logFile)
		{
			Log log = null;
			if (File.Exists(logFile)) log = Log.LoadSettings(logFile);

			if (null != log)
			{
				this.srcCombo.Items.Clear();
				this.srcCombo.Text = log.SrcFile;
				this.targetText.Text = log.TargetFolder;
				this.CsvToListView(log.ListItems);
				this.isJp = log.IsJp;
				this.lineCount = log.LineCount;
				this.charCount = log.CharCount;
				this.matchLinesTable = log.MatchLinesTable;
				//Debug.WriteLine("ListItems:" + log.ListItems);
			}
		}

		private string ListViewToCsv()
		{
			StringBuilder sb = new StringBuilder();
			foreach (ListViewItem lvi in this.listView1.Items)
			{
				bool isFirst = true;
				foreach (ListViewItem.ListViewSubItem si in lvi.SubItems)
				{
					if (!isFirst) sb.Append(",");
					sb.Append(String.Format("\"{0}\"", si.Text));
					isFirst = false;
				}
				sb.Append("\n");
			}
			return sb.ToString();
		}

		private void CsvToListView(string csv)
		{
			string[] lines = csv.Split('\n');
			this.listView1.Items.Clear();
			foreach (string line in lines)
			{
				string[] items = ConvertEx.SplitCommaString(line);
				this.listView1.Items.Add(new ListViewItem(items));
			}
		}

		private void SelectAction()
		{
			foreach (string srcFile in this.srcCombo.Items)
			{
				if (!File.Exists(srcFile) || !Directory.Exists(this.targetText.Text))
				{
					MessageBox.Show("検索ファイルが指定されていないか、存在しないファイルが含まれています。");
					return;
				}
			}
			string indexFile = SearchJob.GetIndexFileName(this.SrcFile);
			if (string.IsNullOrEmpty(indexFile))
			{
				MessageBox.Show("検索元のインデックスファイルを作成してください。");
				return;
			}
			string message = @"検索を開始しますか？
　「はい」　　　… 予約したものと合わせて検索実行する
　「いいえ」　　… 検索内容を予約して、後で検索する
　「キャンセル」… 操作を取り消す";
			//メッセージボックスを表示する
			DialogResult result = MessageBox.Show(message,
				"処理の選択",
					MessageBoxButtons.YesNoCancel,
					MessageBoxIcon.Exclamation,
					MessageBoxDefaultButton.Button2);

			//何が選択されたか調べる
			if (result == DialogResult.Yes)
			{
				//「はい」が選択された時
				this.AddReservation();
				this.SearchReservationList();
			}
			else if (result == DialogResult.No)
			{
				//「いいえ」が選択された時
				this.AddReservation();
			}
			else if (result == DialogResult.Cancel)
			{
				//「キャンセル」が選択された時
				return;
			}
		}

		private void AddReservation()
		{
			foreach (string item in this.srcCombo.Items)
			{
				foreach (Reservation r in this.reservationList)
				{
					if (item.Equals(r.SrcFile) && this.targetText.Text.Equals(r.TargetFolder)) return;
				}
				Reservation newReserve = new Reservation(item, this.targetText.Text, this.isJp);
				this.reservationList.Add(newReserve);
			}
		}

		private void SearchReservationList()
		{
			//Debug.WriteLine("### SearchReservationList");
			this.progressText.Clear();
			this.messageLabel.Text = string.Empty;
			this.countLabel.Text = string.Format("検索テキスト{0}件中{1}件目検索中", this.srcCombo.Items.Count, this.srcIndex + 1);
			this.folderBrowserDialog2.SelectedPath = this.xlsdir;
			if (0 == this.srcIndex && this.folderBrowserDialog2.ShowDialog() == DialogResult.OK)
			{
				this.xlsdir = this.folderBrowserDialog2.SelectedPath;
			}
			//Debug.WriteLine(string.Format("xlsdir={0}", this.xlsdir));
			this.WriteLog(string.Format("Xlsdir was selected: {0}", this.xlsdir));
			var task = Task.Factory.StartNew(() =>
			{
				//Debug.WriteLine("Task begins ...");
				Reservation r = null;
                if (this.srcIndex < this.reservationList.Count) r = this.reservationList[this.srcIndex];
				if (null != r)
                //foreach (Reservation r in this.reservationList)
                {
                    //Debug.WriteLine(string.Format("r.SrcFile={0}", r.SrcFile));
                    bool result = false;
                    this.isProgressing = true;
					this.Invoke(
						(MethodInvoker)delegate()
						{
							this.srcCombo.Text = r.SrcFile;
							this.targetText.Text = r.TargetFolder;
							this.isJp = r.IsJp;
						}
					);
					//Debug.WriteLine(string.Format("r.SrcFile={0}, this.SrcFile={1}", r.SrcFile, this.SrcFile));
					List<string> docs = this.FindDocuments();
                    if (0 < docs.Count) {
                        this.agentPrograms = new List<Process>();
                        this.StartAgentPrograms();
                        using (SearchJob job = new SearchJob(this))
						{
							job.Docs = docs;
							job.SrcFile = r.SrcFile;
							job.MinWords = this.MinWords;
							job.RoughLines = ConvertEx.GetInt(this.RoughLines);
							job.WordCount = this.WordCount;
							job.IsJp = r.IsJp;
							job.RateLevel = ConvertEx.GetDouble(this.rateText.Text) / 100;
							result = job.StartSearch();
						}
                        foreach (Process process in this.agentPrograms) { process.Kill(); }
                    }
					if (!result) {
						this.UpdateMessageLabel(r.SrcFile + "の検索に失敗しましたので処理を中止します。");
						//break;
					}
                    string fname = Path.Combine(Application.StartupPath, "client.tmp");
                    if (this.srcIndex == this.reservationList.Count - 1 || this.reservationList.Count < 2)
					{
						this.srcIndex = 0;
                        if (File.Exists(fname)) File.Delete(fname);
					}
					else
					{
						this.srcIndex++;
						if (!File.Exists(fname))
						{
							using (StreamWriter sw = new StreamWriter(fname, true))
							{
								string str = "write test";
								sw.Write(str);
							}

						}
						this.Close();
					}
                }
                this.reservationList.Clear();
			});
		}

		public void ClearListView()
		{
			this.listView1.Items.Clear();
		}

		/*public void UpdateCountLabel(string message)
		{
			this.Invoke(
				(MethodInvoker)delegate()
				{
					this.countLabel.Text = message;
				}
			);
		}*/

		public void UpdateMessageLabel(string message)
		{
			this.Invoke(
				(MethodInvoker)delegate()
				{
					this.messageLabel.Text = message;
				}
			);
		}

		public void UpdateProgressText(string message)
		{
			this.Invoke(
				(MethodInvoker)delegate()
				{
					this.progressText.AppendText(DateTime.Now.ToString("HH:mm:ss "));
					this.progressText.AppendText(message);
					this.progressText.AppendText("\n");
				}
			);
		}

		public void FinishSearch(string message, Dictionary<int, Dictionary<int, MatchLine>> matchLinesTable, string srcFile)
		{
			this.isProgressing = false;
			this.matchLinesTable = matchLinesTable;
			this.Invoke(
				(MethodInvoker)delegate()
				{
					this.messageLabel.Text = "検索が完了しました。";
					this.countLabel.Text = message;
					this.SaveLog(srcFile);
					this.OutputWordCountList();
				}
			);
		}

		private List<string> GetSrcFiles()
		{
			List<string> srcFiles = new List<string>();
			foreach (string srcFile in this.srcCombo.Items)
			{
				srcFiles.Add(srcFile);
			}
			return srcFiles;
		}

		private void SetSrcFiles(List<string> srcFiles)
		{
			if (null == srcFiles) return;
			this.srcCombo.Items.Clear();
			foreach (string srcFile in srcFiles)
			{
				this.srcCombo.Items.Add(srcFile);
			}
			if (0 < this.srcCombo.Items.Count) this.srcCombo.SelectedIndex = 0;
		}

		private int[] GetWordCountByRate(Dictionary<int, MatchLine> matchLines)
		{
			int[] results = new int[10];
			foreach (KeyValuePair<int, MatchLine> pair in matchLines)
			{
				MatchLine matchLine = pair.Value;
				int level = (int)(Math.Ceiling((1D - matchLine.Rate) * 10));
				if (level < 0) level = 0;
				else if (9 < level) level = 9;
				results[level] += matchLine.MatchWords;
			}
			return results;
		}

		private void OutputWordCountList()
		{
			string[] headers = new string[] { "検索元ファイル名", "検索先ファイル名", "100%", "90%以上", "80%以上", "70%以上", "60%以上", "50%以上", "合計" };
			Int32 colCount = headers.Length;
			Int32 iRow = 0;
			IRow row;
			ICell cell;
			// ワークブックオブジェクト生成
			HSSFWorkbook workbook = new HSSFWorkbook();
			// シートオブジェクト生成
			ISheet sheet1 = workbook.CreateSheet("WordCountList");
			// セルスタイル（黒線）
			ICellStyle blackBorder = workbook.CreateCellStyle();
			blackBorder.BorderBottom = BorderStyle.Thin;
			blackBorder.BorderLeft = BorderStyle.Thin;
			blackBorder.BorderRight = BorderStyle.Thin;
			blackBorder.BorderTop = BorderStyle.Thin;
			blackBorder.BottomBorderColor = HSSFColor.Black.Index;
			blackBorder.LeftBorderColor = HSSFColor.Black.Index;
			blackBorder.RightBorderColor = HSSFColor.Black.Index;
			blackBorder.TopBorderColor = HSSFColor.Black.Index;
			//ヘッダー行の作成
			row = sheet1.CreateRow(iRow);
			// セルを作成する（水平方向）
			for (int i = 0; i < colCount; i++)
			{
				cell = row.CreateCell(i);
				// セルに黒色の枠を付ける
				cell.CellStyle = blackBorder;
				cell.SetCellValue(headers[i]);
			}
			iRow++;
			string fname = Path.GetFileName(this.srcCombo.Text);
			// セルを作成する（垂直方向）
			foreach (ListViewItem lvi in this.listView1.Items)
			{
				int docId = ConvertEx.GetInt(lvi.SubItems[3].Text);
				double rate = ConvertEx.GetDouble(lvi.SubItems[0].Text);
				if (rate <= 0D) continue;
				Dictionary<int, MatchLine> matchLines = new Dictionary<int, MatchLine>();
				if (this.matchLinesTable.ContainsKey(docId)) matchLines = this.matchLinesTable[docId];
				else continue;
				int[] arr = this.GetWordCountByRate(matchLines);
				int total = 0;
				foreach (int n in arr)
				{
					total += n;
				}
				row = sheet1.CreateRow(iRow);
				// セルを作成する（水平方向）
				for (int i = 0; i < colCount; i++)
				{
					cell = row.CreateCell(i);
					// セルに黒色の枠を付ける
					cell.CellStyle = blackBorder;
					if (0 == i)
					{
						cell.SetCellValue(fname);
					}
					else if (1 == i)
					{
						cell.SetCellValue(Path.GetFileName(lvi.SubItems[2].Text));
					}
					else if (colCount - 1 == i)
					{
						cell.SetCellValue(total);
					}
					else
					{
						cell.SetCellValue(arr[i - 2]);
					}
				}
				iRow++;
			}
			// Excelファイル出力
			var timespan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
			string xlsname = Path.Combine(
				this.xlsdir,
				string.Format("{0}.{1}.xls", Path.GetFileNameWithoutExtension(this.srcCombo.Text),
				(uint)timespan.TotalSeconds)
				);
			this.OutputExcelFile(xlsname, workbook);
			this.WriteLog(string.Format("Search result was saved to {0}", xlsname));
		}

		//-------------------------------------------------
		// Excelファイル出力
		//-------------------------------------------------
		private void OutputExcelFile(String strFileName, HSSFWorkbook workbook)
		{
			FileStream file = new FileStream(strFileName, FileMode.Create);
			workbook.Write(file);
			file.Close();
		}

		//-------------------------------------------------
		// セルの位置を取得する
		//-------------------------------------------------
		private String GetCellPos(Int32 iRow, Int32 iCol)
		{
			iCol = Convert.ToInt32('A') + iCol;
			iRow = iRow + 1;
			return ((char)iCol) + iRow.ToString();
		}

		public void WriteLog(string Log)
		{
			if (string.IsNullOrEmpty(Log.Trim())) return;
			Debug.WriteLine(Log);
			this.logs.Add(string.Format("[{0}] {1}", DateTime.Now, Log));
		}

		private void WriteErrorLog()
		{
			string Log = string.Join("\r\n", this.logs.ToArray());
			string path = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "log");
			ErrorLog.Instance.WriteErrorLog(path, Log);
			this.logs.Clear();
		}

        private void StartAgentPrograms()
        {
            var assembly = System.Reflection.Assembly.GetEntryAssembly();
            // Get the full path of the assembly
            string filePath = assembly.Location;
            string dir = Path.GetDirectoryName(filePath);
            string pname = Path.Combine(dir, "Arx.DocSearch.Agent_1_8.exe");
			if (File.Exists(pname))
			{
                Process p = Process.Start(pname, "/IndexOfUser=01");
                this.agentPrograms.Add(p);
            }
			else
			{
				MessageBox.Show(string.Format("プログラム'{0}'が存在しません。", pname),
					"エラー",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error);
				return;
			}
			/*
			int total = 8;
            for (int i = 2; i <= total; i++)
            {
                pname = Path.Combine(dir, string.Format("Arx.DocSearch.Agent_{0}_{1}.exe", i, total));
				if (File.Exists(pname))
				{
					Process sub = Process.Start(pname);
					this.agentPrograms.Add(sub);
				}
				else
				{
					MessageBox.Show(string.Format("プログラム'{0}'が存在しません。", pname),
						"エラー",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);
					return;
				}
            }
			*/
        }

        private void RestartClientProgram()
        {
            var assembly = System.Reflection.Assembly.GetEntryAssembly();
            // Get the full path of the assembly
            string filePath = assembly.Location;
            string dir = Path.GetDirectoryName(filePath);
            string pname = Path.Combine(dir, "Arx.DocSearch.Client.exe");
            if (File.Exists(pname))
            {
                Process p = Process.Start(pname);
            }
            else
            {
                MessageBox.Show(string.Format("プログラム'{0}'が存在しません。", pname),
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

		private void StartNodeManager()
		{
			try {
				Console.WriteLine("StartNodeManager");
;				NMInitializeA("Both");
				NMOpenConfig(0);
				uint Cluster = 0;
				NMGetCluster(ref Cluster);
				int BoardCount = 0;
				NMGetBoardCount(Cluster, ref BoardCount);
				Debug.WriteLine(string.Format("Cluster={0} BoardCount={1}", Cluster, BoardCount));
				for (int NOBoard = 1; NOBoard <= BoardCount; NOBoard++)
                {
					uint Board = 0;
					NMGetBoard(Cluster, NOBoard, ref Board);
                    NMLogIn(Board);
                    string FileName = "Arx.DocSearch.Agent_1_8.exe";
                    uint ProcessHandle = 0;
                    NMStartProgram(1, FileName, "/IndexOfUser=01", ProcessHandle);
                    NMLogOut();

                }
                NodeManager.NMCloseConfig();
                NodeManager.NMFinalize();

			} catch (Exception ex) {
				this.WriteLog(ex.Message);
			}

        }
        private void StopNodeManager()
        {
			NMInitializeA("Both");
			NMOpenConfig(0);
            uint DResult = 0;
			uint Cluster = 0;
			NMGetCluster(ref Cluster);
			int BoardCount = 0;
			NMGetBoardCount(Cluster, ref BoardCount);
			for (int NOBoard = 1; NOBoard <= BoardCount ; NOBoard++)
            {
				uint Board = 0;
				NMGetBoard(Cluster, NOBoard, ref Board);
				NMLogIn(Board);
                NMStopProgram(1);
                NMLogOut();

            }
            NodeManager.NMCloseConfig();
            NodeManager.NMFinalize();
        }
        #endregion

    }
}
