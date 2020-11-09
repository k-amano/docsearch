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

namespace Arx.DocSearch.MultiCore
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
			this.DOnExecuteTask = new THCExecuteTaskEvent(this.DoOnExecuteTask);
			this.DOnGetProgress = new THCGetProgressEvent(this.DoOnGetProgress);
			this.DOnInterrupt = new THCInterruptEvent(this.DoOnInterrupt);
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
		//private readonly int TEST_MAX_LINES = 5000;
		private readonly int TEST_MAX_LINES = 200;
		private readonly int ROUGH_COUNT = 20;
		private readonly double ROUGH_RATE = 0.5;
		private readonly int TARGET_ROUGH_LINES = 50;
		private const int UserIndex = 1;
		private const int SlotCount = 256;
		private int SlotIndex;
		private bool Processing;
		private bool Initialized;
		private double[] AgentProgressArray;
		private bool[] AgentInterruptedArray;
		public uint[] TextDataArray;
		public uint[] IndexDataArray;
		private THCGetMemoryEvent DOnGetMemory;
		private THCFreeMemoryEvent DOnFreeMemory;
		private THCExecuteTaskEvent DOnExecuteTask;
		private THCGetProgressEvent DOnGetProgress;
		private THCInterruptEvent DOnInterrupt;
		private THCOperateEvent DOnStartOperate;
		private THCDeliverEvent DOnStartDeliver;
		private THCDeliverEvent DOnFinishDeliver;
		private THCRaceEvent DOnStartRace;
		private THCCleanEvent DOnStartClean;
		private THCBatchEvent DOnStartBatch;
		private THCExecuteEvent DOnStartExecute;
		private THCExecuteEvent DOnFinishExecute;
		public FileChoice fileChoice;
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
			if (this.Initialized)
			{
				this.CleanFile(FileChoice.TEXT_FILE);
				this.CleanFile(FileChoice.INDEX_FILE);
				Multi.HCFinalize();
			}
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
			this.matchList.Clear();
			this.CleanFile(FileChoice.TEXT_FILE);
			this.DeliverFile(FileChoice.TEXT_FILE);
			this.CleanFile(FileChoice.INDEX_FILE);
			this.DeliverFile(FileChoice.INDEX_FILE);
			try
			{
				Debug.WriteLine("HCOperate Start");
				Multi.HCOperate(0,
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
				Multi.HCGetProgress(ref Progress);
				this.mainForm.UpdateMessageLabel(string.Format("{0} 文書中 {1:0.00}% 完了。開始 {2} 現在 {3}。", this.docs.Count, Progress * 100, this.startTime.ToLongTimeString(), DateTime.Now.ToLongTimeString()));
				this.Processing = false;
			}
		}

		private Dictionary<int, MatchLine> SearchDocument(string doc, int docId, ref int matchCount, ref double rate)
		{
			Dictionary<int, MatchLine> matchLines = new Dictionary<int, MatchLine>();
			List<int> targetLines = new List<int>();
			string targetDoc = SearchJob.GetTextFileName(doc);
			if (!File.Exists(targetDoc)) return matchLines;
			List<string> paragraphs = this.GetParagraphs(targetDoc);
			matchCount = 0;
			int total = 0;
			int characterCount = 0;
			int pos = 0;
			int linesCount = lines.Count;
#if DEBUG
			if (TEST_MAX_LINES < linesCount) linesCount = TEST_MAX_LINES;
#endif
			List<int> roughMatchLines = new List<int>();
			if (0 < this.roughLines)
			{
				this.mainForm.UpdateProgressText(string.Format("{0} ラフ検索を開始しました。", doc));
				this.SearchRough(this.linesIdx, paragraphs, roughMatchLines);
				this.mainForm.UpdateProgressText(string.Format("{0} ラフ検索を完了しました。", doc));
			}
			for (int i = 0; i < linesCount; i++)
			{
				if (string.IsNullOrEmpty((lines[i]))) continue;
				if (0 == this.roughLines || roughMatchLines.Contains(i))
				{
					int targetLine = 0;
					int totalWords = 0;
					int matchWords = 0;
					pos = 0; // ラフ検索を含めて常に最初から検索する。
					double lineRate = this.SearchLine(lines[i], this.linesIdx[i], paragraphs, targetLines, i, ref pos, ref total, ref targetLine, ref totalWords, ref matchWords);
					if (0D < lineRate)
					{
						MatchLine matchLine = new MatchLine(lineRate, targetLine, totalWords, matchWords);
						matchLines.Add(i, matchLine);
						matchCount++;
					}
				}
				else
				{
					total++;
				}
				characterCount += lines[i].Length;
			}
			rate = 0 == total ? 0D : (double)matchCount / (double)total;
			return matchLines;
		}

		public double SearchLine(string line, string lineIdx, List<string> paragraphs, List<int> targetLines, int no, ref int pos, ref int totalCount, ref int targetLine, ref int totalWords, ref int matchWords)
		{
			if (string.IsNullOrEmpty(line)) return 0;
			string[] words = lineIdx.Split(' ');
			if (words.Length < this.minWords) return 0;
			totalCount++;
			double rate = 0;
			for (int i = pos; i < paragraphs.Count; i++)
			{
				if (string.IsNullOrEmpty(paragraphs[i])) continue;
				if (targetLines.Contains(i))
				{
					continue;
				}
				string src = this.isJp ? lineIdx : line;
				rate = this.GetDiffRate(src, paragraphs[i], ref totalWords, ref matchWords);
				//指定一致率以上であればここで終了。
				if (this.rateLevel <= rate)
				{
					pos = i + 1;
					targetLine = i;
					targetLines.Add(i);
					return rate;
				}
			}
			return 0; //指定一致率以下
		}

		private double GetDiffRate(string src, string target, ref int wordCount, ref int matchCount)
		{
			var d = new Differ();
			var inlineBuilder = new InlineDiffBuilder(d);
			var result = d.CreateWordDiffs(src.Trim(), target.Trim(), true, new char[] { ' ' });
			int diffCount = 0;
			foreach (var block in result.DiffBlocks)
			{
				diffCount += block.DeleteCountA;
			}
			matchCount = result.PiecesOld.Length - diffCount;
			if (matchCount < 0) matchCount = 0;
			wordCount = Math.Max(result.PiecesOld.Length, result.PiecesNew.Length);
			double rate = (double)matchCount / (double)wordCount;
			return rate;
		}

		public void SearchRough(List<string> linesIdx, List<string> paragraphs, List<int> roughMatchLines)
		{
			int count = linesIdx.Count;
			int paraCount = paragraphs.Count;
#if DEBUG
			if (TEST_MAX_LINES < count) count = TEST_MAX_LINES;
			if (TEST_MAX_LINES < paraCount) paraCount = TEST_MAX_LINES;
#endif
			StringBuilder sbSrc = new StringBuilder();
			int offset = 0;
			if (this.roughLines < 1)
			{
				this.roughLines = 1;
			}
			List<string> roughpara = new List<string>();
			for (int i = 0; i < paraCount; i += this.TARGET_ROUGH_LINES)
			{
				StringBuilder sb = new StringBuilder();
				for (int j = i; j < i + TARGET_ROUGH_LINES + roughLines && j < paraCount; j++)
				{
					sb.Append(paragraphs[j]);
					sb.Append(" ");
				}
				roughpara.Add(sb.ToString());
			}
			for (int i = 0; i < count; i++)
			{
				sbSrc.Append(linesIdx[i]);
				sbSrc.Append(" ");
				offset++;
				if (this.roughLines <= offset)
				{
					string[] words = sbSrc.ToString().Split(new char[] { ' ' });
					for (int j = 0; j < roughpara.Count; j++)
					{
						if (this.GetRoughRate(words, roughpara[j]))
						{
							for (int k = i - offset; k <= i; k++)
							{
								if (!roughMatchLines.Contains(k)) roughMatchLines.Add(k);
							}
							break;
						}
					}
					offset = 0;
					sbSrc = new StringBuilder();
				}
			}
			if (0 < offset)
			{
				string[] words = sbSrc.ToString().Split(new char[] { ' ' });
				for (int j = 0; j < roughpara.Count; j++)
				{
					if (this.GetRoughRate(words, roughpara[j]))
					{
						for (int k = (count - 1) - offset; k < count; k++)
						{
							roughMatchLines.Add(k);
						}
					}
				}
			}
		}

		private bool GetRoughRate(string[] words, string paragraph)
		{
			if (string.IsNullOrEmpty(paragraph)) return false;
			List<int> ls = this.GetRandom(words.Length);
			int pos = 0;
			int matchCount = 0;
			for (int i = 0; i < ls.Count; i++)
			{
				int j = ls[i];
				int newPos = paragraph.IndexOf(words[j], pos);
				if (pos <= newPos)
				{
					matchCount += 1;
					pos = newPos + 1;
				}
				//20 word 進んで一致率70%未満は終了
				if ((ls.Count <= i + 1 || this.ROUGH_COUNT * this.ROUGH_RATE < i) && matchCount < i * this.ROUGH_RATE)
				{
					return false;
				}
				if (paragraph.Length <= pos) break;
			}
			return true;
		}

		private List<int> GetRandom(int max)
		{
			System.Random r = new System.Random();
			//シード値を指定しないとシード値として Environment.TickCount が使用される
			List<int> ls = new List<int>();
			do
			{
				int i1 = r.Next(max);
				if (!ls.Contains(i1)) ls.Add(i1);
			} while (ls.Count < this.ROUGH_COUNT && ls.Count < max);
			ls.Sort();
			return ls;
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

		private List<string> GetParagraphs(string fname)
		{
			string line;
			List<string> paragraphs = new List<string>();
			int i = 0;
			using (StreamReader file = new StreamReader(fname))
			{
				while ((line = file.ReadLine()) != null)
				{
					if (this.isJp) line = TextConverter.SplitWords(line);
					paragraphs.Add(line);
					i++;
				}
			}
			return paragraphs;
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
				BinaryFormatter f = new BinaryFormatter();
				//読み込んで逆シリアル化する
				ret = f.Deserialize(ms);
			}
			return ret;
		}

		private byte[] HCReadMemory(uint description, int length)
		{
			IntPtr address = Marshal.AllocHGlobal(length * sizeof(byte));
			Multi.HCReadMemoryEntry(description, (uint)address.ToInt32());
			byte[] bb = new byte[length];
			Marshal.Copy(address, bb, 0, length);
			return bb;
		}

		private void HCWriteMemory(uint description, byte[] bb)
		{
			IntPtr address = Marshal.AllocHGlobal(bb.Length * sizeof(byte));
			Marshal.Copy(bb, 0, address, bb.Length);
			Multi.HCWriteMemoryEntry(description, (uint)address.ToInt32(), (uint)bb.Length);
		}

		private void DeliverFile(FileChoice Choice)
		{
			Debug.WriteLine("DeliverFile");
			this.fileChoice = Choice;
			Multi.HCOperate(1,
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
			Debug.WriteLine("CleanFile");
			this.fileChoice = Choice;
			Multi.HCOperate(2,
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

		/* セッションイベント */
		private void DoOnStartOperate(int Code, uint NodeList, ref bool DoDeliver, ref bool DoRace, ref bool DoClean)
		{
			Debug.WriteLine("DoOnStartOperate");
			if ((Code == 0) || (Code == 2)) DoDeliver = false;
			if ((Code == 1) || (Code == 2)) DoRace = false;
			if ((Code == 0) || (Code == 1)) DoClean = false;
		}

		private void DoOnStartDeliver(int Code, uint DataList, uint NewDataList)
		{
			Debug.WriteLine("DoOnStartDeliver");
			string fname;
			if (this.fileChoice == FileChoice.TEXT_FILE) fname = SearchJob.GetTextFileName(this.SrcFile);
			else if (this.fileChoice == FileChoice.INDEX_FILE) fname = SearchJob.GetIndexFileName(this.SrcFile);
			else return;
			Debug.WriteLine(string.Format("DoOnStartDeliver fname={0}", fname));
			if (!File.Exists(fname)) return;
			//8文字のランダムな文字列を作成する
			Guid g = System.Guid.NewGuid();
			string guid = g.ToString("N").Substring(0, 8);
			string ext = Path.GetExtension(fname);
			string newFile = string.Format("{0}{1:X8}{2}", guid, System.Environment.TickCount ,ext);
			File.Copy(fname, newFile);
			Debug.WriteLine(string.Format("DoOnStartDeliver newFile.Exists={0}", File.Exists(newFile)));
			uint data = 0;
			Multi.HCCreateFileData(Multi.HCClientNode(), newFile, ref data);
			Debug.WriteLine(string.Format("fileName={0} data={1:X8}", newFile, data));
			Multi.HCAddDataToList(Multi.HCNodeDataList(Multi.HCClientNode()), data);
			Multi.HCAddDataToList(DataList, data);
			Multi.HCAddDataToList(NewDataList, data);
		}

		private void DoOnFinishDeliver(int Code, uint DataList, uint NewDataList)
		{
			Debug.WriteLine("DoOnFinishDeliver");
			for (int nodeIndex = 1; nodeIndex <= Multi.HCNodeCountInList(Multi.HCGlobalNodeList()); nodeIndex++)
			{
				if (this.fileChoice == FileChoice.TEXT_FILE) this.TextDataArray[nodeIndex - 1] = 0;
				else if (this.fileChoice == FileChoice.INDEX_FILE) this.IndexDataArray[nodeIndex - 1] = 0;
			}
			for (int dataIndex = 1; dataIndex <= Multi.HCDataCountInList(NewDataList); dataIndex++)
			{
				uint data = Multi.HCDataInList(NewDataList, dataIndex);
				int nodeIndex = Multi.HCNodeIndexInList(Multi.HCGlobalNodeList(), Multi.HCDataNode(data));
				if (this.fileChoice == FileChoice.TEXT_FILE) this.TextDataArray[nodeIndex - 1] = data;
				else if (this.fileChoice == FileChoice.INDEX_FILE) this.IndexDataArray[nodeIndex - 1] = data;
			}
		}

		private void DoOnStartRace(int Code, ref int BatchCount)
		{
			BatchCount = this.docs.Count;
			Debug.WriteLine(string.Format("DoOnStartRace Code={0} BatchCount={1}", Code, BatchCount));
		}

		private void DoOnStartClean(int Code, uint DataList)
		{
			Debug.WriteLine(string.Format("DoOnStartClean fileChoice={0}", this.fileChoice));
			while (Multi.HCDataCountInList(DataList) > 0)
			{
				Multi.HCDeleteDataFromList(DataList, 1);
			}
			for (int nodeIndex = 1; nodeIndex <= Multi.HCNodeCountInList(Multi.HCGlobalNodeList()); nodeIndex++)
			{
				switch (this.fileChoice)
				{
					case FileChoice.TEXT_FILE:
						if (this.TextDataArray[nodeIndex - 1] != 0)
						{
							Multi.HCAddDataToList(DataList, this.TextDataArray[nodeIndex - 1]);
						}
						break;
					case FileChoice.INDEX_FILE:
						if (this.IndexDataArray[nodeIndex - 1] != 0)
						{
							Multi.HCAddDataToList(DataList, this.IndexDataArray[nodeIndex - 1]);
						}
						break;
					default: break;
				}
			}
			Debug.WriteLine("DoOnStartClean Finished");
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
			int nodeIndex = Multi.HCNodeIndexInList(Multi.HCGlobalNodeList(), Node);
			Multi.HCClearDescription(Description);
			Multi.HCWriteAnsiStringEntry(Description, Multi.HCFileDataFileName(this.TextDataArray[nodeIndex - 1]));
			Multi.HCWriteAnsiStringEntry(Description, Multi.HCFileDataFileName(this.IndexDataArray[nodeIndex - 1]));
			Multi.HCWriteAnsiStringEntry(Description, this.docs[BatchIndex - 1]);
			Multi.HCWriteLongEntry(Description, BatchIndex - 1);
			Multi.HCWriteLongEntry(Description, ConvertEx.GetInt(this.WordCount));
			Multi.HCWriteAnsiStringEntry(Description, this.wordCount);
			Multi.HCWriteLongEntry(Description, this.roughLines);
			Multi.HCWriteDoubleEntry(Description, this.rateLevel);
			Multi.HCWriteBoolEntry(Description, this.isJp);
			this.mainForm.UpdateProgressText(string.Format("{0} 検索を開始しました。", this.docs[BatchIndex - 1]));
			this.showProgress();
			Debug.WriteLine("DoOnStartExecute Finished");
		}

		private void DoOnFinishExecute(int Code, int BatchIndex, uint Node, uint Description,
			uint DataList, uint NewDataList)
		{
			//文書番号をHCから取得する
			int index = Multi.HCReadLongEntry(Description);
			Debug.WriteLine(string.Format("DoOnFinishExecute BatchIndex={0} index={1}", BatchIndex, index));
			//バイト配列のサイズをHCから取得する
			int matchLinesLength = Multi.HCReadLongEntry(Description);
			int matchDocumentLength = Multi.HCReadLongEntry(Description);
			//バイト配列をHCから取得する
			byte[] bMatchLines = this.HCReadMemory(Description, matchLinesLength);
			byte[] bMatchDocument = this.HCReadMemory(Description, matchDocumentLength);
			//バイト配列を検索結果インスタンスにデシリアライズする
			Dictionary<int, MatchLine> matchLines = (this.DeserializeObject(bMatchLines)) as Dictionary<int, MatchLine>;
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

		private void DoOnExecuteTask(int SlotIndex, uint Description)
		{
			Debug.WriteLine("DoOnExecuteTask");
			//検索テキストファイル名をHCから取得する
			string textFile = HCInterface.Multi.HCReadAnsiStringEntry(Description);
			//検索インデックスファイル名をHCから取得する
			string indexFile = HCInterface.Multi.HCReadAnsiStringEntry(Description);
			//文書名をHCから取得する
			string docName = HCInterface.Multi.HCReadAnsiStringEntry(Description);
			//Debug.WriteLine(string.Format("DoOnExecuteTask SlotIndex={0} srcFile={1} docName={2}", SlotIndex, srcFile, docName));
			//文書番号をHCから取得する
			int index = HCInterface.Multi.HCReadLongEntry(Description);
			this.minWords = HCInterface.Multi.HCReadLongEntry(Description);
			this.wordCount = HCInterface.Multi.HCReadAnsiStringEntry(Description);
			this.roughLines = HCInterface.Multi.HCReadLongEntry(Description);
			this.rateLevel = HCInterface.Multi.HCReadDoubleEntry(Description);
			this.isJp = HCInterface.Multi.HCReadBoolEntry(Description);
			//Debug.WriteLine(string.Format("DoOnExecuteTask SlotIndex={0}", SlotIndex));
			byte[] bMatchLines = new byte[0];
			byte[] bMatchDocument = new byte[0];
			if (File.Exists(textFile) && File.Exists(indexFile) && File.Exists(docName))
			{
				//Debug.WriteLine(string.Format("DoOnExecuteTask File.Exists SlotIndex={0}", SlotIndex));
				this.lines.Clear();
				using (StreamReader file = new StreamReader(textFile))
				{
					string line;
					while ((line = file.ReadLine()) != null)
					{
						this.lines.Add(line);
					}
				}
				//Debug.WriteLine(string.Format("DoOnExecuteTask textFile SlotIndex={0}", SlotIndex));
				this.linesIdx.Clear();
				using (StreamReader file = new StreamReader(indexFile))
				{
					string line;
					while ((line = file.ReadLine()) != null)
					{
						this.linesIdx.Add(line);
					}
				}
				//Debug.WriteLine(string.Format("DoOnExecuteTask indexFile SlotIndex={0}", SlotIndex));
				int matchCount = 0;
				double rate = 0D;
				//検索結果インスタンスを取得する
				//Debug.WriteLine(string.Format("DoOnExecuteTask SearchDocument SlotIndex={0}", SlotIndex));
				Dictionary<int, MatchLine> matchLines = this.SearchDocument(docName, index, ref matchCount, ref rate);
				//Debug.WriteLine(string.Format("DoOnExecuteTask SlotIndex={0} matchLines.Count={1}", SlotIndex, matchLines.Count));
				MatchDocument md = new MatchDocument(rate, matchCount, docName, index);
				//検索結果インスタンスをバイト配列にシリアライズする
				bMatchLines = this.SerializeObject(matchLines);
				//this.DebugLog(string.Format("DoOnExecuteTask  bMatchLines.Length={0}", bMatchLines.Length));
				bMatchDocument = this.SerializeObject(md);
			}
			//文書番号をHCに登録する
			HCInterface.Multi.HCWriteLongEntry(Description, index);
			//バイト配列のサイズをHCに登録する
			HCInterface.Multi.HCWriteLongEntry(Description, (int)bMatchLines.Length);
			HCInterface.Multi.HCWriteLongEntry(Description, (int)bMatchDocument.Length);
			//バイト配列をHCに登録する
			this.HCWriteMemory(Description, bMatchLines);
			this.HCWriteMemory(Description, bMatchDocument);
			Debug.WriteLine("DoOnExecuteTask Finished");
		}

		private void DoOnGetProgress(int SlotIndex, ref double Progress)
		{
			Progress = this.AgentProgressArray[SlotIndex - 1];
			//Debug.WriteLine(string.Format("DoOnGetProgress SlotIndex={0}, Progress={1}", SlotIndex, Progress));
		}

		private void DoOnInterrupt(int SlotIndex)
		{
			Debug.WriteLine("DoOnInterrupt");
			this.AgentInterruptedArray[SlotIndex - 1] = true;
		}

		private void InitializeHC()
		{
			Debug.WriteLine("InitializeHC");
			this.Processing = false;
			this.Initialized = false;
			HCInterface.Multi.HCInitialize(UserIndex, this.DOnGetMemory,
				this.DOnFreeMemory,
				this.DOnExecuteTask,
				this.DOnGetProgress,
				this.DOnInterrupt);
			this.Initialized = true;
			this.AgentProgressArray = new double[SlotCount];
			this.AgentInterruptedArray = new bool[SlotCount];
			for (this.SlotIndex = 0; this.SlotIndex < SlotCount; this.SlotIndex++)
			{
				this.AgentProgressArray[SlotIndex] = 0;
				this.AgentInterruptedArray[SlotIndex] = false;
			}
			int NodeCount = Multi.HCNodeCountInList(Multi.HCGlobalNodeList());
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
