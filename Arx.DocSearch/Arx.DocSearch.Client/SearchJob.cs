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
			this.DOnStartDeliver = new THCDeliverEvent(this.DoOnStartDeliver);
			this.DOnFinishDeliver = new THCDeliverEvent(this.DoOnFinishDeliver);
			this.DOnStartRace = new THCRaceEvent(this.DoOnStartRace);
			this.DOnStartClean = new THCCleanEvent(this.DoOnStartClean);
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
		public uint[] TextDataArray;
		public uint[] IndexDataArray;
		private THCGetMemoryEvent DOnGetMemory;
		private THCFreeMemoryEvent DOnFreeMemory;
		private THCOperateEvent DOnStartOperate;
		private THCDeliverEvent DOnStartDeliver;
		private THCDeliverEvent DOnFinishDeliver;
		private THCRaceEvent DOnStartRace;
		private THCCleanEvent DOnStartClean;
		private THCBatchEvent DOnStartBatch;
		private THCExecuteEvent DOnStartExecute;
		private THCExecuteEvent DOnFinishExecute;
		public FileChoice fileChoice;
		private object lockObject = new object();
		#endregion

		#region Enum
		public enum FileChoice
		{
			TEXT_FILE, INDEX_FILE
		}
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

		public bool StartSearch()
		{
			try
			{
				this.WriteLog("StartSearch");
				if (!File.Exists(this.SrcFile))
				{
					return false;
				}
				this.mainForm.Invoke(
					(MethodInvoker)delegate()
					{
						this.mainForm.ClearListView();
					}
				);
				this.MatchLinesTable.Clear();
				this.matchList.Clear();
				this.CleanFile(FileChoice.TEXT_FILE);
				this.DeliverFile(FileChoice.TEXT_FILE);
				this.CleanFile(FileChoice.INDEX_FILE);
				this.DeliverFile(FileChoice.INDEX_FILE);
				try
				{
					this.WriteLog("HCOperate Begin");
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
					this.WriteLog("HCOperate End");
				}
				catch (Exception e)
				{
					this.WriteLog(e.Message);
					this.WriteLog(e.Message + e.StackTrace);
					throw;
				}
				// リストをID順でソートする
				if (0 < this.matchList.Count) {
					MatchDocument[] matchArray = this.matchList.ToArray();
					Array.Sort(matchArray, (a, b) => (int)((a.Rate - b.Rate) * -1000000));
					this.mainForm.updateListView(matchArray, true);
				}
				this.mainForm.FinishSearch(string.Format("{0} 文書中 100% 完了。開始 {1} 終了 {2}。", docs.Count, this.startTime.ToLongTimeString(), DateTime.Now.ToLongTimeString()), this.matchLinesTable, this.srcFile);
			}
			catch (Exception e)
			{
				this.WriteLog(e.Message + e.StackTrace);
				return false;
			}
			return true;
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
			else return string.Empty;
		}

		public static string GetIndexFileName(string srcFile)
		{
			string dir = Path.Combine(Path.GetDirectoryName(srcFile), ".adsidx");
			string indexFile = Path.Combine(dir, Path.GetFileName(srcFile) + ".idx");
			//if (File.Exists(indexFile) && File.GetLastWriteTime(srcFile) <= File.GetLastWriteTime(indexFile)) return indexFile;
			if (File.Exists(indexFile)) return indexFile;
			return string.Empty;
		}

		private object DeserializeObject(byte[] bb)
		{
			object ret = null;
			using (MemoryStream ms = new MemoryStream(bb))
			{
				try
				{
					BinaryFormatter f = new BinaryFormatter();
					//読み込んで逆シリアル化する
					ret = f.Deserialize(ms);
				}
				catch (SerializationException e)
				{
					this.WriteLog(string.Format("逆シリアル化に失敗しました bbLength:{0} {1}", bb.Length, e.Message));
				}
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

		private void DeliverFile(FileChoice Choice)
		{
			this.WriteLog("DeliverFile");
			this.fileChoice = Choice;
			HCInterface.Client.HCOperate(1,
								this.DOnStartOperate, null,
								this.DOnStartDeliver, this.DOnFinishDeliver,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null);
		}

		private void CleanFile(FileChoice Choice)
		{
			this.WriteLog("CleanFile");
			this.fileChoice = Choice;
			HCInterface.Client.HCOperate(2,
								this.DOnStartOperate, null,
								null, null,
								null, null,
								this.DOnStartClean, null,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null,
								null, null);
		}

		private void WriteLog(string Log)
		{
			this.mainForm.WriteLog(Log);
		}

		/* セッションイベント */
		private void DoOnStartOperate(int Code, uint NodeList, ref bool DoDeliver, ref bool DoRace, ref bool DoClean)
		{
			this.WriteLog("DoOnStartOperate");
			if ((Code == 0) || (Code == 2)) DoDeliver = false;
			if ((Code == 1) || (Code == 2)) DoRace = false;
			if ((Code == 0) || (Code == 1)) DoClean = false;
		}

		private void DoOnStartDeliver(int Code, uint DataList, uint NewDataList)
		{
			this.WriteLog("DoOnStartDeliver");
			string fname;
			if (this.fileChoice == FileChoice.TEXT_FILE) fname = SearchJob.GetTextFileName(this.SrcFile);
			else if (this.fileChoice == FileChoice.INDEX_FILE) fname = SearchJob.GetIndexFileName(this.SrcFile);
			else return;
			this.WriteLog(string.Format("DoOnStartDeliver fname={0}", fname));
			if (!File.Exists(fname)) return;
			//8文字のランダムな文字列を作成する
			Guid g = System.Guid.NewGuid();
			string guid = g.ToString("N").Substring(0, 8);
			string ext = Path.GetExtension(fname);
			string newFile = string.Format("{0}{1:X8}{2}", guid, System.Environment.TickCount, ext);
			File.Copy(fname, newFile);
			this.WriteLog(string.Format("DoOnStartDeliver newFile.Exists={0}", File.Exists(newFile)));
			uint data = 0;
			HCInterface.Client.HCCreateFileData(HCInterface.Client.HCClientNode(), newFile, ref data);
			this.WriteLog(string.Format("fileName={0} data={1:X8}", newFile, data));
			HCInterface.Client.HCAddDataToList(HCInterface.Client.HCNodeDataList(HCInterface.Client.HCClientNode()), data);
			HCInterface.Client.HCAddDataToList(DataList, data);
			HCInterface.Client.HCAddDataToList(NewDataList, data);
		}

		private void DoOnFinishDeliver(int Code, uint DataList, uint NewDataList)
		{
			this.WriteLog("DoOnFinishDeliver");
			for (int nodeIndex = 1; nodeIndex <= HCInterface.Client.HCNodeCountInList(HCInterface.Client.HCGlobalNodeList()); nodeIndex++)
			{
				if (this.fileChoice == FileChoice.TEXT_FILE) this.TextDataArray[nodeIndex - 1] = 0;
				else if (this.fileChoice == FileChoice.INDEX_FILE) this.IndexDataArray[nodeIndex - 1] = 0;
			}
			for (int dataIndex = 1; dataIndex <= HCInterface.Client.HCDataCountInList(NewDataList); dataIndex++)
			{
				uint data = HCInterface.Client.HCDataInList(NewDataList, dataIndex);
				int nodeIndex = HCInterface.Client.HCNodeIndexInList(HCInterface.Client.HCGlobalNodeList(), HCInterface.Client.HCDataNode(data));
				if (this.fileChoice == FileChoice.TEXT_FILE) this.TextDataArray[nodeIndex - 1] = data;
				else if (this.fileChoice == FileChoice.INDEX_FILE) this.IndexDataArray[nodeIndex - 1] = data;
			}
		}

		private void DoOnStartRace(int Code, ref int BatchCount)
		{
			BatchCount = this.docs.Count;
			this.WriteLog(string.Format("DoOnStartRace Code={0} BatchCount={1}", Code, BatchCount));
		}

		private void DoOnStartClean(int Code, uint DataList)
		{
			this.WriteLog(string.Format("DoOnStartClean fileChoice={0}", this.fileChoice));
			while (HCInterface.Client.HCDataCountInList(DataList) > 0)
			{
				HCInterface.Client.HCDeleteDataFromList(DataList, 1);
			}
			for (int nodeIndex = 1; nodeIndex <= HCInterface.Client.HCNodeCountInList(HCInterface.Client.HCGlobalNodeList()); nodeIndex++)
			{
				switch (this.fileChoice)
				{
					case FileChoice.TEXT_FILE:
						if (this.TextDataArray[nodeIndex - 1] != 0)
						{
							HCInterface.Client.HCAddDataToList(DataList, this.TextDataArray[nodeIndex - 1]);
						}
						break;
					case FileChoice.INDEX_FILE:
						if (this.IndexDataArray[nodeIndex - 1] != 0)
						{
							HCInterface.Client.HCAddDataToList(DataList, this.IndexDataArray[nodeIndex - 1]);
						}
						break;
					default: break;
				}
			}
			this.WriteLog("DoOnStartClean Finished");
		}

		private void DoOnStartBatch(int Code, int BatchIndex, uint Node,
			ref bool DoSend, ref bool DoExecute, ref bool DoReceive, ref bool DoWaste,
			uint DataList, uint NewDataList)
		{
			this.WriteLog("DoOnStartBatch");
			DoSend = false;
		}

		private void DoOnStartExecute(int Code, int BatchIndex, uint Node, uint Description,
			uint DataList, uint NewDataList)
		{
			try
			{
				this.WriteLog(string.Format("DoOnStartExecute BatchIndex={0}", BatchIndex));
				int nodeIndex = HCInterface.Client.HCNodeIndexInList(HCInterface.Client.HCGlobalNodeList(), Node);
				HCInterface.Client.HCClearDescription(Description);
				HCInterface.Client.HCWriteAnsiStringEntry(Description, HCInterface.Client.HCFileDataFileName(this.TextDataArray[nodeIndex - 1]));
				HCInterface.Client.HCWriteAnsiStringEntry(Description, HCInterface.Client.HCFileDataFileName(this.IndexDataArray[nodeIndex - 1]));
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
			catch (Exception e)
			{
				this.WriteLog(e.Message + e.StackTrace);
			}
		}

		private void DoOnFinishExecute(int Code, int BatchIndex, uint Node, uint Description,
			uint DataList, uint NewDataList)
		{
			lock (this.lockObject)
			{
				try
				{
					//文書番号をHCから取得する
					int index = HCInterface.Client.HCReadLongEntry(Description);
					this.WriteLog(string.Format("[{0}] DoOnFinishExecute BatchIndex={1} index={2} doc={3}", DateTime.Now, BatchIndex, index, this.docs[BatchIndex - 1]));
					//バイト配列のサイズをHCから取得する
					int matchLinesLength = HCInterface.Client.HCReadLongEntry(Description);
					this.WriteLog(string.Format("[{0}] DoOnFinishExecute matchLinesLength={1} doc={2}", DateTime.Now, matchLinesLength, this.docs[BatchIndex - 1]));
					int matchDocumentLength = HCInterface.Client.HCReadLongEntry(Description);
					//バイト配列をHCから取得する
					byte[] bMatchLines = this.HCReadMemory(Description, matchLinesLength);
					this.WriteLog(string.Format("[{0}] DoOnFinishExecute bMatchLines.Length={1} doc={2}", DateTime.Now, bMatchLines.Length, this.docs[BatchIndex - 1]));
					byte[] bMatchDocument = this.HCReadMemory(Description, matchDocumentLength);
					if (0 < bMatchLines.Length && 0 < bMatchDocument.Length)
					{
						//バイト配列を検索結果インスタンスにデシリアライズする
						Dictionary<int, MatchLine> matchLines = (this.DeserializeObject(bMatchLines)) as Dictionary<int, MatchLine>;
						this.WriteLog(string.Format("DoOnFinishExecute matchLines.Count={0} doc={1}", matchLines.Count, this.docs[BatchIndex - 1]));
						MatchDocument md = (this.DeserializeObject(bMatchDocument)) as MatchDocument;
						//デシリアライズしたインスタンスをリストに追加する。
						this.matchList.Add(md);
						this.MatchLinesTable.Add(index, matchLines);
						this.mainForm.UpdateProgressText(string.Format("{0} 検索を完了しました。", this.docs[BatchIndex - 1]));
					}
					else
					{
						this.WriteLog(string.Format("DoOnFinishExecute matchLines is null doc={0}", this.docs[BatchIndex - 1]));
					}
					this.showProgress();
				}
				catch (Exception e)
				{
					this.WriteLog(e.Message + e.StackTrace);
				}
			}
		}

		/* ノードイベント */
		private void DoOnGetMemory(int SlotIndex, uint Size, ref uint Address)
		{
			this.WriteLog("DoOnGetMemory");
			Address = Convert.ToUInt32(Marshal.AllocHGlobal((int)Size).ToInt32());
		}

		private void DoOnFreeMemory(int SlotIndex, uint Address)
		{
			this.WriteLog("DoOnFreeMemory");
			Marshal.FreeHGlobal(new IntPtr(Convert.ToInt32(Address)));
		}

		private void InitializeHC()
		{
			this.WriteLog("InitializeHC");
			this.Processing = false;
			this.Initialized = false;
			HCInterface.Client.HCInitialize(UserIndex, this.DOnGetMemory,
				this.DOnFreeMemory);
			this.Initialized = true;
			int NodeCount = HCInterface.Client.HCNodeCountInList(HCInterface.Client.HCGlobalNodeList());
			this.TextDataArray = new uint[NodeCount];
			this.IndexDataArray = new uint[NodeCount];
			for (int NodeIndex = 1; NodeIndex < NodeCount; NodeIndex++)
			{
				this.TextDataArray[NodeIndex - 1] = 0;
				this.IndexDataArray[NodeIndex - 1] = 0;
			}
		}
		#endregion
	}
}
