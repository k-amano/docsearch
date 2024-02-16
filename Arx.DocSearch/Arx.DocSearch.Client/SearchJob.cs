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
using DiffPlex.Model;

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
            this.startSearchTime = DateTime.Now;
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
        private DateTime startSearchTime;
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
		private List<int>[] dataPacks;
		private string TextFileName;
		private string IndexFileName;
		private const int packCount = 128;
		private int progressCount = 0;
        private double progeessValue = 0D;

        private THCGetMemoryEvent DOnGetMemory;
		private THCFreeMemoryEvent DOnFreeMemory;
		private THCOperateEvent DOnStartOperate;
		private THCRaceEvent DOnStartRace;
		private THCBatchEvent DOnStartBatch;
		private THCExecuteEvent DOnStartExecute;
		private THCExecuteEvent DOnFinishExecute;
		public FileChoice fileChoice;
		private object lockObject = new object();
        //二重解放を避けるためのフラグ
        private bool disposedValue = false;
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

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    //管理リソースの破棄処理をここに記述
                }

                //非管理リソースの破棄処理をここに記述

                disposedValue = true;
            }
        }

        // ファイナライザー
        ~SearchJob()
        {
            Dispose(false);
        }
        public void Dispose() 
		{
			Debug.WriteLine("Dispose");
			if (this.Initialized) HCInterface.Client.HCFinalize();
            Dispose(true);

            //ファイナライザーを呼ばない事を
            //ガベージコレクションに指示する
            //GC.SuppressFinalize(this);
        }

		public bool StartSearch()
		{
			try
			{
				this.WriteLog(string.Format("StartSearch SrcFile={0}", this.SrcFile));
                this.startSearchTime = DateTime.Now;
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
				this.dataPacks = this.getDataPack(packCount);
				this.WriteLog(string.Format("StartSearch TextFileName={0} IndexFileName={1}", this.TextFileName, this.IndexFileName));
				//this.showMemory("StartSearch");
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
                //this.showMemory("StartSearch");
            }
			catch (Exception e)
			{
				this.WriteLog(e.Message + e.StackTrace);
				return false;
			}
			return true;
		}

        public void showProgress()
		{
			double Progress = 0;
			if (!this.Processing)
			{
				this.Processing = true;
				HCInterface.Client.HCGetProgress(ref Progress);
                double progress2 = this.progressCount / (packCount * 11) + 0.001;
				if (Progress < 0.05 && Progress < progress2 && progress2 < 1) Progress = progress2;
				if (Progress < this.progeessValue) Progress = this.progeessValue;
                this.progeessValue = Progress;
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
			Marshal.FreeHGlobal(address);
            return bb;
		}

		private void HCWriteMemory(uint description, byte[] bb)
		{
			IntPtr address = Marshal.AllocHGlobal(bb.Length * sizeof(byte));
			Marshal.Copy(bb, 0, address, bb.Length);
			HCInterface.Client.HCWriteMemoryEntry(description, (uint)address.ToInt32(), (uint)bb.Length);
            Marshal.FreeHGlobal(address);
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
            if (this.docs.Count < packCount) BatchCount = this.docs.Count;
            else BatchCount = packCount;
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
				this.WriteLog(string.Format("DoOnStartExecute BatchIndex={0} dataPacks.Length={1} docs.Count={2}", BatchIndex, this.dataPacks.Length, this.docs.Count));
                //this.showMemory("DoOnStartExecute1");
                List<int> dataPack = new List<int>();
				if (BatchIndex <= this.dataPacks.Length) dataPack = this.dataPacks[BatchIndex - 1];
				List<string> targetDocs = new List<string>();
				for (int i = 0; i < dataPack.Count; i++)
				{
					int index = dataPack[i];
                    if (index < this.docs.Count) targetDocs.Add(this.docs[index]);
				}
				int nodeIndex = HCInterface.Client.HCNodeIndexInList(HCInterface.Client.HCGlobalNodeList(), Node);
				HCInterface.Client.HCClearDescription(Description);
				HCInterface.Client.HCWriteAnsiStringEntry(Description, this.TextFileName);
				HCInterface.Client.HCWriteAnsiStringEntry(Description, this.IndexFileName);
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
				//this.WriteLog(string.Format("DoOnStartExecute lines.Count={0}  TextFileName={1} Exists={2}", lines.Count, this.TextFileName, File.Exists(this.TextFileName)));
				using (StreamReader file = new StreamReader(this.IndexFileName))
				{
					string line;
					while ((line = file.ReadLine()) != null)
					{
						linesIdx.Add(line);
					}
				}
				byte[] bdataPack = this.SerializeObject(dataPack);
				byte[] btargetDocs = this.SerializeObject(targetDocs);
				byte[] bLines = this.SerializeObject(lines);
				byte[] binesIdx = this.SerializeObject(linesIdx);
				//バイト配列のサイズをHCに登録する
				HCInterface.Client.HCWriteLongEntry(Description, (int)bdataPack.Length);
				HCInterface.Client.HCWriteLongEntry(Description, (int)btargetDocs.Length);
				HCInterface.Client.HCWriteLongEntry(Description, (int)bLines.Length);
				HCInterface.Client.HCWriteLongEntry(Description, (int)binesIdx.Length);
                //バイト配列をHCに登録する
                //this.showMemory("DoOnStartExecute2");
                this.HCWriteMemory(Description, bdataPack);
				this.HCWriteMemory(Description, btargetDocs);
				this.HCWriteMemory(Description, bLines);
				this.HCWriteMemory(Description, binesIdx);
				this.mainForm.UpdateProgressText(string.Format("{0} 検索を開始しました。", this.docs[BatchIndex - 1]));
				this.showProgress();
				this.WriteLog(string.Format("DoOnStartExecute Finished: BatchIndex={0}", BatchIndex));
                //this.showMemory("DoOnStartExecute3");
                //GC.Collect();
                //this.showMemory("DoOnStartExecute4");
            }
			catch (Exception e)
			{
				this.WriteLog(e.Message + e.StackTrace);
			}
            this.progressCount++;
        }

		private void DoOnFinishExecute(int Code, int BatchIndex, uint Node, uint Description,
			uint DataList, uint NewDataList)
		{
			lock (this.lockObject)
			{
				try
				{
					//文書番号をHCから取得する
					//int index = HCInterface.Client.HCReadLongEntry(Description);
					this.WriteLog(string.Format("[{0}] DoOnFinishExecute BatchIndex={1} doc={2}", DateTime.Now, BatchIndex, this.docs[BatchIndex - 1]));
					//バイト配列のサイズをHCから取得する
					int dataPackLength = HCInterface.Client.HCReadLongEntry(Description);
					int mlListLength = HCInterface.Client.HCReadLongEntry(Description);
					int mdListLength = HCInterface.Client.HCReadLongEntry(Description);
					//バイト配列をHCから取得する
					byte[] bdataPack = this.HCReadMemory(Description, dataPackLength);
					byte[] bmlList = this.HCReadMemory(Description, mlListLength);
					byte[] bmdList = this.HCReadMemory(Description, mdListLength);
					//バイト配列を検索結果インスタンスにデシリアライズする
					List<int> dataPack = (this.DeserializeObject(bdataPack)) as List<int>;
					List<Dictionary<int, MatchLine>> mlList = (this.DeserializeObject(bmlList)) as List<Dictionary<int, MatchLine>>;
					List<MatchDocument> mdList = (this.DeserializeObject(bmdList)) as List<MatchDocument>;
					//this.WriteLog(string.Format("[{0}] DoOnFinishExecute bMatchLines.Length={1} doc={2}", DateTime.Now, bMatchLines.Length, this.docs[BatchIndex - 1]));
					for (int i = 0; i < dataPack.Count; i++)
					{
						int index = dataPack[i];
						Debug.WriteLine(String.Format("DoOnFinishExecute: index={0}", index));
						if (i < mlList.Count && i < mdList.Count)
						{
							Dictionary<int, MatchLine> matchLines = mlList[i];
							MatchDocument md = mdList[i];
							//デシリアライズしたインスタンスをリストに追加する。
							this.matchList.Add(md);
							this.MatchLinesTable.Add(index, matchLines);
							this.mainForm.UpdateProgressText(string.Format("{0} 検索を完了しました。", this.docs[index]));
						}
						else
						{
							this.WriteLog(string.Format("DoOnFinishExecute matchLines is null doc={0}", this.docs[index]));
						}
					}
					this.showProgress();
                    //this.showMemory("DoOnFinishExecute1");
                    //GC.Collect();
                    //this.showMemory("DoOnFinishExecute2");
                }
				catch (Exception e)
				{
					this.WriteLog(e.Message + e.StackTrace);
				}
			}
            this.WriteLog("DoOnFinishExecute end");
            this.progressCount += 10;
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
            try
            {
                this.WriteLog("InitializeHC");
                this.Processing = false;
                this.Initialized = false;
                HCInterface.Client.HCInitialize(UserIndex, this.DOnGetMemory,
                    this.DOnFreeMemory);
                this.Initialized = true;
                int NodeCount = HCInterface.Client.HCNodeCountInList(HCInterface.Client.HCGlobalNodeList());
            }
            catch (Exception e)
            {
                this.WriteLog(e.Message + e.StackTrace);
            }
        }
		private List<int>[] getDataPack(int packCount)
		{
			if (this.docs.Count < packCount) packCount = this.docs.Count;
			List<int>[] dataPacks = new List<int>[packCount];
			for (int i = 0; i < packCount; i++)
			{
				List<int> pack = new List<int>();
				dataPacks[i] = pack;
			}
			Random r = new System.Random();
			for (int i = 0; i < this.docs.Count; i++)
			{
				int n = r.Next(0, packCount - 1);
				dataPacks[n].Add(i);
			}
			return dataPacks;
		}

		private void showMemory(string message) {
            // 物理メモリ量 ex) 135,815,168
            Console.WriteLine($"{message}:System.Environment.WorkingSet:{System.Environment.WorkingSet:N0}");
            // Processインスタンスを取得
            System.Diagnostics.Process proc = System.Diagnostics.Process.GetCurrentProcess();

            // 物理メモリ量 ex) 136,282,112
            //Console.WriteLine($"proc.WorkingSet64:{proc.WorkingSet64:N0}");
            // 仮想メモリ量 ex) 290,201,600
            //Console.WriteLine($"proc.VirtualMemorySize64:{proc.VirtualMemorySize64:N0}");

        }
		#endregion
	}
}
