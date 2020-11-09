using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using Xyn.Util;
using ConvertEx = Xyn.Util.ConvertEx;
using HCInterface;
using Arx.DocSearch.Util;

namespace Arx.DocSearch.Client
{
	public class SearchJob : IDisposable 
	{
		#region コンストラクタ
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public SearchJob(MainForm mainForm)
		{
			this.mainForm = mainForm;
			this.docs = new List<string>();
			this.lines = new List<string>();
			this.linesIdx = new List<string>();
			this.startTime = DateTime.Now;
			this.matchList = new List<MatchDocument>();
			this.matchLinesTable = new Dictionary<int, Dictionary<int, MatchLine>>();
			this.DOnGetMemory = new THCGetMemoryEvent(this.DoOnGetMemory);
			this.DOnFreeMemory = new THCFreeMemoryEvent(this.DoOnFreeMemory);
			this.DOnStartOperate = new THCOperateEvent(this.DoOnStartOperate);
			this.DOnStartRace = new THCRaceEvent(this.DoOnStartRace);
			this.DOnStartBatch = new THCBatchEvent(this.DoOnStartBatch);
			this.DOnStartExecute = new THCExecuteEvent(this.DoOnStartExecute);
			this.DOnFinishExecute = new THCExecuteEvent(this.DoOnFinishExecute);
			//HarmonyCalcを初期化
			this.InitializeHC();
		}
		#endregion

		#region フィールド
		private MainForm mainForm;
		private List<string> docs;
		private List<string> lines;
		private List<string> linesIdx;
		private DateTime startTime;
		private int roughLines;
		private List<MatchDocument> matchList;
		private string srcFile;
		private int minWords;
		private string wordCount;
		private Dictionary<int, Dictionary<int, MatchLine>> matchLinesTable;
		private bool isJp = false;
		private double rateLevel;
		private const int UserIndex = 1;
		private bool Processing;
		private bool Initialized;
		private THCGetMemoryEvent DOnGetMemory;
		private THCFreeMemoryEvent DOnFreeMemory;
		private THCOperateEvent DOnStartOperate;
		private THCRaceEvent DOnStartRace;
		private THCBatchEvent DOnStartBatch;
		private THCExecuteEvent DOnStartExecute;
		private THCExecuteEvent DOnFinishExecute;
		#endregion

		#region Property
		public List<string> Docs
		{
			get
			{
				return docs;
			}
			set
			{
				docs = value;
			}
		}
		/// <summary>
		/// ログイン時にアクセスする Web ページの URL を取得または設定します。
		/// </summary>
		public string SrcFile
		{
			get
			{
				return srcFile;
			}
			set
			{
				srcFile = value;
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
				return wordCount;
			}
			set
			{
				wordCount = value;
			}
		}

		public int RoughLines
		{
			get
			{
				return roughLines;
			}
			set
			{
				roughLines = value;
			}
		}

		public double RateLevel
		{
			get
			{
				return rateLevel;
			}
			set
			{
				rateLevel = value;
			}
		}

		public Dictionary<int, Dictionary<int, MatchLine>> MatchLinesTable
		{
			get
			{
				return matchLinesTable;
			}
			set
			{
				matchLinesTable = value;
			}
		}

		public bool IsJp
		{
			get
			{
				return isJp;
			}
			set
			{
				isJp = value;
			}
		}
		#endregion

		#region メソッド
		public void Dispose() 
		{
			Debug.WriteLine("Dispose");
			if (this.Initialized) HCInterface.Client.HCFinalize();
		}

		public void StartSearch()
		{
			Debug.WriteLine("StartSearch");
			if (!File.Exists(this.SrcFile))
			{
				return;
			}
			this.mainForm.Invoke(
				(MethodInvoker)delegate()
				{
					this.mainForm.ClearListView();
				}
			);
			this.MatchLinesTable.Clear();
			try
			{
				Debug.WriteLine("HCOperate Start");
				HCInterface.Client.HCOperate(0,
						this.DOnStartOperate, null,
						null, null,
						this.DOnStartRace, null,
						null, null,
						this.DOnStartBatch, null,
						null, null,
						this.DOnStartExecute, this.DOnFinishExecute,
						null, null,
						null, null,
						null, null,
						null, null);
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.Message);
				throw;
			}
			// リストをID順でソートする
			MatchDocument[] matchArray = this.matchList.ToArray();
			Array.Sort(matchArray, (a, b) => (int)((a.Rate - b.Rate) * -1000000));
			this.mainForm.updateListView(matchArray, true);
			this.mainForm.FinishSearch(string.Format("{0} 文書中 100% 完了。開始 {1} 終了 {2}。", docs.Count, this.startTime.ToLongTimeString(), DateTime.Now.ToLongTimeString()), this.matchLinesTable);
		}

		private void showProgress()
		{
			double Progress = 0;
			if (!this.Processing)
			{
				this.Processing = true;
				HCInterface.Client.HCGetProgress(ref Progress);
				this.mainForm.UpdateMessageLabel(string.Format("{0} 文書中 {1:0.00}% 完了。開始 {2} 現在 {3}。", this.docs.Count, Progress * 100, this.startTime.ToLongTimeString(), DateTime.Now.ToLongTimeString()));
				this.Processing = false;
			}
		}

		public static string GetTextFileName(string srcFile)
		{
			if (!File.Exists(srcFile)) return string.Empty;
			string dir = Path.Combine(Path.GetDirectoryName(srcFile), ".adsidx");
			string textFile = Path.Combine(dir, Path.GetFileName(srcFile) + ".txt");
			if (File.Exists(textFile)) return textFile;
			return string.Empty;
		}

		public static string GetIndexFileName(string srcFile)
		{
			string dir = Path.Combine(Path.GetDirectoryName(srcFile), ".adsidx");
			string indexFile = Path.Combine(dir, Path.GetFileName(srcFile) + ".idx");
			if (File.Exists(indexFile) && File.GetLastWriteTime(srcFile) <= File.GetLastWriteTime(indexFile)) return indexFile;
			return string.Empty;
		}

		private object DeserializeObject(byte[] bb)
		{
			object ret = null;
			using (MemoryStream ms = new MemoryStream(bb))
			{
				BinaryFormatter f = new BinaryFormatter();
				//読み込んで逆シリアル化する
				ret = f.Deserialize(ms);
			}
			return ret;
		}

		private byte[] HCReadMemory(uint description, int length)
		{
			IntPtr address = Marshal.AllocHGlobal(length * sizeof(byte));
			HCInterface.Client.HCReadMemoryEntry(description, (uint)address.ToInt32());
			byte[] bb = new byte[length];
			Marshal.Copy(address, bb, 0, length);
			return bb;
		}

		/* セッションイベント */
		private void DoOnStartOperate(int Code, uint NodeList, ref bool DoDeliver, ref bool DoRace, ref bool DoClean)
		{
			Debug.WriteLine("DoOnStartOperate");
			DoDeliver = false;
			DoClean = false;
		}

		private void DoOnStartRace(int Code, ref int BatchCount)
		{
			BatchCount = this.docs.Count;
			Debug.WriteLine(string.Format("DoOnStartRace Code={0} BatchCount={1}", Code, BatchCount));
		}

		private void DoOnStartBatch(int Code, int BatchIndex, uint Node,
			ref bool DoSend, ref bool DoExecute, ref bool DoReceive, ref bool DoWaste,
			uint DataList, uint NewDataList)
		{
			Debug.WriteLine("DoOnStartBatch");
			DoSend = false;
		}

		private void DoOnStartExecute(int Code, int BatchIndex, uint Node, uint Description,
			uint DataList, uint NewDataList)
		{
			Debug.WriteLine(string.Format("DoOnStartExecute BatchIndex={0}", BatchIndex));
			HCInterface.Client.HCWriteAnsiStringEntry(Description, this.SrcFile);
			HCInterface.Client.HCWriteAnsiStringEntry(Description, this.docs[BatchIndex - 1]);
			HCInterface.Client.HCWriteLongEntry(Description, BatchIndex - 1);
			HCInterface.Client.HCWriteLongEntry(Description, ConvertEx.GetInt(this.WordCount));
			HCInterface.Client.HCWriteAnsiStringEntry(Description, this.wordCount);
			HCInterface.Client.HCWriteLongEntry(Description, this.roughLines);
			HCInterface.Client.HCWriteDoubleEntry(Description, this.rateLevel);
			HCInterface.Client.HCWriteBoolEntry(Description, this.isJp);
			this.mainForm.UpdateProgressText(string.Format("{0} 検索を開始しました。", this.docs[BatchIndex - 1]));
			this.showProgress();
		}

		private void DoOnFinishExecute(int Code, int BatchIndex, uint Node, uint Description,
			uint DataList, uint NewDataList)
		{
			//文書番号をHCから取得する
			int index = HCInterface.Client.HCReadLongEntry(Description);
			Debug.WriteLine(string.Format("[{0}] DoOnFinishExecute BatchIndex={1} index={2} doc={3}", DateTime.Now, BatchIndex, index, this.docs[BatchIndex - 1]));
			//バイト配列のサイズをHCから取得する
			int matchLinesLength = HCInterface.Client.HCReadLongEntry(Description);
			Debug.WriteLine(string.Format("[{0}] DoOnFinishExecute matchLinesLength={1} doc={2}", DateTime.Now, matchLinesLength, this.docs[BatchIndex - 1]));
			int matchDocumentLength = HCInterface.Client.HCReadLongEntry(Description);
			//バイト配列をHCから取得する
			byte[] bMatchLines = this.HCReadMemory(Description, matchLinesLength);
			Debug.WriteLine(string.Format("[{0}] DoOnFinishExecute bMatchLines.Length={1} doc={2}", DateTime.Now, bMatchLines.Length, this.docs[BatchIndex - 1]));
			byte[] bMatchDocument = this.HCReadMemory(Description, matchDocumentLength);
			//バイト配列を検索結果インスタンスにデシリアライズする
			Dictionary<int, MatchLine> matchLines = (this.DeserializeObject(bMatchLines)) as Dictionary<int, MatchLine>;
			if (null == matchLines) Debug.WriteLine("matchLines is null");
			else Debug.WriteLine(string.Format("[{0}] DoOnFinishExecute matchLines.Count={1} doc={2}", DateTime.Now, matchLines.Count, this.docs[BatchIndex - 1]));

			MatchDocument md = (this.DeserializeObject(bMatchDocument)) as MatchDocument;
			//デシリアライズしたインスタンスをリストに追加する。
			this.matchList.Add(md);
			this.MatchLinesTable.Add(index, matchLines);
			this.mainForm.UpdateProgressText(string.Format("{0} 検索を完了しました。", this.docs[BatchIndex - 1]));
			this.showProgress();
		}

		/* ノードイベント */
		private void DoOnGetMemory(int SlotIndex, uint Size, ref uint Address)
		{
			Debug.WriteLine("DoOnGetMemory");
			Address = Convert.ToUInt32(Marshal.AllocHGlobal((int)Size).ToInt32());
		}

		private void DoOnFreeMemory(int SlotIndex, uint Address)
		{
			Debug.WriteLine("DoOnFreeMemory");
			Marshal.FreeHGlobal(new IntPtr(Convert.ToInt32(Address)));
		}

		private void InitializeHC()
		{
			Debug.WriteLine("InitializeHC");
			this.Processing = false;
			this.Initialized = false;
			HCInterface.Client.HCInitialize(UserIndex, this.DOnGetMemory,
				this.DOnFreeMemory);
			this.Initialized = true;
		}
		#endregion
	}
}
