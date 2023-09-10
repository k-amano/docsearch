using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using View = System.Windows.Forms.View;
using Xyn.Util;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Arx.DocSearch.Util;

namespace Arx.DocSearch.Client
{
	public partial class MainForm : Form
	{
		[DllImport("xd2txlib.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Cdecl)]
		public static extern int ExtractText(
				[MarshalAs(UnmanagedType.BStr)] String lpFileName,
				bool bProp,
				[MarshalAs(UnmanagedType.BStr)] ref String lpFileText);

		#region コンストラクタ
		public MainForm()
		{
			InitializeComponent();
			this.matchLinesTable = new Dictionary<int, Dictionary<int, MatchLine>>();
			this.reservationList = new List<Reservation>();
		}
		#endregion

		#region フィールド
		private Schema config;
		private string configFile;
		private readonly int LINE_LENGHTH = 70;
		private int minWords;
		private Dictionary<int, Dictionary<int, MatchLine>> matchLinesTable;
		private int lineCount = 0;
		private int charCount = 0;
		private bool isJp = false;
		private List<Reservation> reservationList;
		private SearchJob job;
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
				return srcText.Text;
			}
			set
			{
				srcText.Text = value;
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
			this.InitializeListView();
			this.configFile = Path.Combine(this.UserAppDataPath, "DocSearch.config");
			this.config = Schema.LoadSettings(this.configFile);
			this.srcText.Text = this.config.SrcFile;
			this.targetText.Text = this.config.TargetFolder;
			this.rateText.Text = this.config.Rate;
			this.wordCountText.Text = this.config.WordCount;
			this.roughLinesText.Text = this.config.RoughLines;
			this.messageLabel.Text = string.Empty;
			this.countLabel.Text = string.Empty;
			this.GetTotalCount();
			this.job = new SearchJob(this);
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (null != this.config)
			{
				this.config.SrcFile = this.srcText.Text;
				this.config.TargetFolder = this.targetText.Text;
				this.config.Rate = this.rateText.Text;
				this.config.WordCount = this.wordCountText.Text;
				this.config.RoughLines = this.roughLinesText.Text;
				this.config.SaveSettings(this.configFile);
			}
			this.job.Dispose();
		}

		private void srcButton_Click(object sender, EventArgs e)
		{
			if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				this.srcText.Text = this.openFileDialog1.FileName;
			}
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
			CompareForm f = new CompareForm(this, lvi, this.srcText.Text, this.isJp, this.lineCount, this.charCount, matchLines);
			f.Show();
		}

		private void indexButton_Click(object sender, EventArgs e)
		{
			if (!File.Exists(this.srcText.Text))
			{
				MessageBox.Show("検索元を選択してください。");
				return;
			}
			var task = Task.Factory.StartNew(() =>
			{
				string textFile = SearchJob.GetTextFileName(this.srcText.Text);
				if (string.IsNullOrEmpty(textFile))
				{
					List<string> ls = new List<string>();
					ls.Add(this.srcText.Text);
					if (".pdf".Equals(Path.GetExtension(this.srcText.Text).ToLower())) this.MakeTextFileFromPdf(ls);
					else this.MakeTextFile(ls);
				}
				this.MakeIndexFile(textFile);
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
				//Debug.WriteLine(file + "   " + Path.GetFileName(file));
				// 「.」で始まるファイルはパスする
				if (Path.GetFileName(file).StartsWith(".") || Path.GetFileName(dir).StartsWith(".")) continue;
				//Debug.WriteLine("##OK");
				if (exts.Contains(Path.GetExtension(file).ToLower()))
				{
					//if (!file.Equals(this.SrcFile)) 
					docs.Add(file);
				}
			}
			return docs;
		}

		private string GetTargetText(string doc)
		{
			string fname = SearchJob.GetTextFileName(doc);
			//Debug.WriteLine(string.Format("doc={0} fname={1}", doc, fname));
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
			StreamWriter writer = null;
			try
			{
				foreach (string srcFile in docFiles)
				{
					string dir = Path.Combine(Path.GetDirectoryName(srcFile), ".adsidx");
					string textFile = Path.Combine(dir, Path.GetFileName(srcFile) + ".txt");
					if (File.Exists(textFile)) continue;
					if (File.Exists(textFile)) File.Delete(textFile);
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
					int l = ExtractText(srcFile, false, ref fileText);
					string[] paragraphs = fileText.Split('\n');
					for (int i = 0; i < paragraphs.Length; i++)
					{
						string line = paragraphs[i].TrimEnd();
						line = this.ReplaceLine(line);
						//if (i < 5) Debug.WriteLine(string.Format("i={0} line={0}", i, line));
						writer.Write(line);
						bool startsWithCapital = false;
						if (i + 1 < paragraphs.Length && this.StartsWithCapital(paragraphs[i + 1])) startsWithCapital = true;
						//if (i < 5) Debug.WriteLine(string.Format("startsWithCapital={0} next={1}", startsWithCapital, paragraphs[i + 1]));
						if (startsWithCapital || line.TrimEnd().EndsWith(".") || line.TrimEnd().EndsWith("。"))
						{
							writer.Write("\n");//ピリオドまたは読点で終わっていれば改行する
						}
						else
						{
							writer.Write(" "); //それ以外は空白を追加。
						}
					}
					writer.Close();
				}
			}
			catch { }
			finally
			{
				if (null != writer)
				{
					writer.Close();
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
				if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
				PdfReader pdfReader = null;
				StreamWriter writer = null;
				ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
				try
				{
					pdfReader = new PdfReader(srcFile);
					writer = File.CreateText(textFile);
					writer.NewLine = "\n";
					for (int page = 1; page <= pdfReader.NumberOfPages; page++)
					{
						string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
						string[] lines = currentText.Split('\n');
						for (int i = 0; i < lines.Length; i++)
						{
							string line = lines[i].TrimEnd();
							line = this.ReplaceLine(line);
							writer.Write(line);
							bool startsWithCapital = false;
							if (i + 1 < lines.Length && this.StartsWithCapital(lines[i + 1])) startsWithCapital = true;
							if (startsWithCapital || line.TrimEnd().EndsWith(".") || line.TrimEnd().EndsWith("。"))
							{

								writer.Write("\n");//ピリオドまたは読点で終わっていれば改行する
							}
							else
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
			//センテンスの終わりで「半角スペース2個」+改行とする。
			line = Regex.Replace(line, @"\. +", ".  \n"); //ピリオド+半角スペース1個は改行
			line = Regex.Replace(line, @"。([^\n])", "。  \n($1)"); //読点。
			line = TextConverter.ZenToHan(line);
			line = TextConverter.HankToZen(line);
			//line = Regex.Replace(line, @"^([\(\[<（＜〔【≪《])([^0-9]*[0-9]*)([\)\]>）＞〕】≫》])(\s*)", "\n$1$2$3$4  \n"); //【数字】
			line = Regex.Replace(line, @"^((([\(\[<（＜〔【≪《])([^0-9]*[0-9]*)([\)\]>）＞〕】≫》])(\s*))+)", "\n$1  \n"); //【数字】
			line = Regex.Replace(line, @"^([0-9]+)(\.?)", "\n$1$2  \n"); //数字
			//line = Regex.Replace(line, @"([.,:;])", " $1 "); //半角句読点は前後にスペース
			return line;
		}

		private bool StartsWithCapital(string line)
		{
			if (Regex.IsMatch(line, @"^(\s*[A-Z])")) return true;
			else return false;
		}

		private void MakeIndexFile(string textFile)
		{
			string dir = Path.GetDirectoryName(textFile);
			string indexFile = Path.Combine(dir, Path.GetFileNameWithoutExtension(textFile) + ".idx");
			if (File.Exists(indexFile)) File.Delete(indexFile);
			//if (File.Exists(indexFile) && File.GetLastWriteTime(textFile) <= File.GetLastWriteTime(indexFile)) return;
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
			string textFile = SearchJob.GetTextFileName(this.srcText.Text);
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


		private void SaveLog()
		{
			string fname = Path.GetFileNameWithoutExtension(this.srcText.Text);
			var timespan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
			string logFile = Path.Combine(this.UserAppDataPath, string.Format("{0}.{1}.log", fname, (uint)timespan.TotalSeconds));
			Log log = new Log();

			//if (null != log)
			{
				log.SrcFile = this.srcText.Text;
				log.TargetFolder = this.targetText.Text;
				log.ListItems = this.ListViewToCsv();
				log.IsJp = this.isJp;
				log.LineCount = this.lineCount;
				log.CharCount = this.charCount;
				log.MatchLinesTable = this.matchLinesTable;
				log.SaveSettings(logFile);
				//Debug.WriteLine("ListItems:" + log.ListItems);
			}
		}

		private void LoadLog(string logFile)
		{
			Log log = null;
			//Debug.WriteLine(logFile);
			if (File.Exists(logFile)) log = Log.LoadSettings(logFile);

			if (null != log)
			{
				this.srcText.Text = log.SrcFile;
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
			if (!File.Exists(this.srcText.Text) || !Directory.Exists(this.targetText.Text))
			{
				MessageBox.Show("検索ファイルを指定してください");
				return;
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
			foreach (Reservation r in this.reservationList)
			{
				if (this.srcText.Text.Equals(r.SrcFile) && this.targetText.Text.Equals(r.TargetFolder)) return;
			}
			Reservation newReserve = new Reservation(this.srcText.Text, this.targetText.Text, this.isJp);
			this.reservationList.Add(newReserve);
		}

		private void SearchReservationList()
		{
			this.progressText.Clear();
			this.messageLabel.Text = string.Empty;
			this.countLabel.Text = string.Empty;
			var task = Task.Factory.StartNew(() =>
			{
				foreach (Reservation r in this.reservationList)
				{
					this.Invoke(
						(MethodInvoker)delegate()
						{
							this.srcText.Text = r.SrcFile;
							this.targetText.Text = r.TargetFolder;
							this.isJp = r.IsJp;
						}
					);
					List<string> docs = this.FindDocuments();
					//using (SearchJob job = new SearchJob(this))
					//{
						job.Docs = docs;
						job.SrcFile = this.SrcFile;
						job.MinWords = this.MinWords;
						job.RoughLines = ConvertEx.GetInt(this.RoughLines);
						job.WordCount = this.WordCount;
						job.IsJp = this.isJp;
						job.RateLevel = ConvertEx.GetDouble(this.rateText.Text) / 100;
						job.StartSearch();
					//}
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
			this.matchLinesTable = matchLinesTable;
			this.Invoke(
				(MethodInvoker)delegate()
				{
					this.messageLabel.Text = "検索が完了しました。";
					this.countLabel.Text = message;
					this.SaveLog();
				}
			);
		}
		#endregion
	}
}
