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
		private string TextFileName;
		private string IndexFileName;

		private THCGetMemoryEvent DOnGetMemory;
		private THCFreeMemoryEvent DOnFreeMemory;
		private THCOperateEvent DOnStartOperate;
		private THCRaceEvent DOnStartRace;
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
				this.WriteLog(string.Format("StartSearch SrcFile={0}", this.SrcFile));
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
				this.TextFileName = SearchJob.GetTextFileName(this.SrcFile);
				this.IndexFileName = SearchJob.GetIndexFileName(this.SrcFile);
				this.WriteLog(string.Format("StartSearch TextFileName={0} IndexFileName={1}", this.TextFileName, this.IndexFileName));
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

		private byte[] SerializeObject(object src)
		{
			byte[] ret = new byte[0];
			using (MemoryStream ms = new MemoryStream())
			{
				BinaryFormatter bf = new BinaryFormatter();
				bf.Serialize(ms, src);
				ret = ms.ToArray();
			}
			return ret;
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

		private void HCWriteMemory(uint description, byte[] bb)
		{
			IntPtr address = Marshal.AllocHGlobal(bb.Length * sizeof(byte));
			Marshal.Copy(bb, 0, address, bb.Length);
			HCInterface.Client.HCWriteMemoryEntry(description, (uint)address.ToInt32(), (uint)bb.Length);
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

		private void DoOnStartRace(int Code, ref int BatchCount)
		{
			BatchCount = this.docs.Count;
			this.WriteLog(string.Format("DoOnStartRace Code={0} BatchCount={1}", Code, BatchCount));
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
				this.WriteLog(string.Format("DoOnStartExecute nodeIndex={0}", nodeIndex));
				HCInterface.Client.HCClearDescription(Description);
				HCInterface.Client.HCWriteAnsiStringEntry(Description, this.TextFileName);
				this.WriteLog(string.Format("DoOnStartExecute TextFileName={0}", TextFileName));
				HCInterface.Client.HCWriteAnsiStringEntry(Description, this.IndexFileName);
				this.WriteLog(string.Format("DoOnStartExecute IndexFileName={0}", IndexFileName));
				HCInterface.Client.HCWriteAnsiStringEntry(Description, this.docs[BatchIndex - 1]);
				this.WriteLog(string.Format("DoOnStartExecute docName={0}", this.docs[BatchIndex - 1]));
				HCInterface.Client.HCWriteLongEntry(Description, BatchIndex - 1);
				HCInterface.Client.HCWriteLongEntry(Description, ConvertEx.GetInt(this.WordCount));
				HCInterface.Client.HCWriteLongEntry(Description, this.roughLines);
				HCInterface.Client.HCWriteDoubleEntry(Description, this.rateLevel);
				HCInterface.Client.HCWriteBoolEntry(Description, this.isJp);
				List<string> lines = new List<string>();
				List<string> linesIdx = new List<string>();
				using (StreamReader file = new StreamReader(this.TextFileName))
				{
					string line;
					while ((line = file.ReadLine()) != null)
					{
						lines.Add(line);
					}
				}
				this.WriteLog(string.Format("DoOnStartExecute lines.Count={0}  TextFileName={1} Exists={2}", lines.Count, this.TextFileName, File.Exists(this.TextFileName)));
				using (StreamReader file = new StreamReader(this.IndexFileName))
				{
					string line;
					while ((line = file.ReadLine()) != null)
					{
						linesIdx.Add(line);
					}
				}
				byte[] bLines = this.SerializeObject(lines);
				byte[] binesIdx = this.SerializeObject(linesIdx);
				this.WriteLog(string.Format("DoOnStartExecute bLines.Length={0}", bLines.Length));
				//バイト配列のサイズをHCに登録する
				HCInterface.Client.HCWriteLongEntry(Description, (int)bLines.Length);
				HCInterface.Client.HCWriteLongEntry(Description, (int)binesIdx.Length);
				//バイト配列をHCに登録する
				this.HCWriteMemory(Description, bLines);
				this.HCWriteMemory(Description, binesIdx);
				this.mainForm.UpdateProgressText(string.Format("{0} 検索を開始しました。", this.docs[BatchIndex - 1]));
				this.showProgress();
				this.WriteLog("DoOnStartExecute5");
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
		}
		#endregion
	}
}
