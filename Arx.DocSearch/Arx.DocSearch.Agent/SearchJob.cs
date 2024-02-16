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
using static System.Windows.Forms.LinkLabel;

namespace Arx.DocSearch.Agent
{
	public class SearchJob : IDisposable 
	{
		#region コンストラクタ
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public SearchJob(MainForm mainForm, int userIndex)
		{
			this.mainForm = mainForm;
			this.userIndex = userIndex;
			//this.lines = new List<string>();
			//this.linesIdx = new List<string>();
			this.DOnGetMemory = new THCGetMemoryEvent(this.DoOnGetMemory);
			this.DOnFreeMemory = new THCFreeMemoryEvent(this.DoOnFreeMemory);
			this.DOnExecuteTask = new THCExecuteTaskEvent(this.DoOnExecuteTask);
			this.DOnGetProgress = new THCGetProgressEvent(this.DoOnGetProgress);
			this.DOnInterrupt = new THCInterruptEvent(this.DoOnInterrupt);
			//HarmonyCalcを初期化
			this.InitializeHC();
		}
		#endregion

		#region フィールド
		private MainForm mainForm;
		//private List<string> lines;
		//private List<string> linesIdx;
		//private int roughLines;
		//private int minWords;
		//private string wordCount;
		//private bool isJp = false;
		//private double rateLevel;
		//private readonly int TEST_MAX_LINES = 5000;
		private readonly int TEST_MAX_LINES = 200;
		private readonly int ROUGH_COUNT = 20;
		private readonly double ROUGH_RATE = 0.5;
		private readonly int TARGET_ROUGH_LINES = 50;
		private int userIndex;
		private const int SlotCount = 256;
		private int SlotIndex;
		private bool Initialized;
		private double[] AgentProgressArray;
		private bool[] AgentInterruptedArray;
		private THCGetMemoryEvent DOnGetMemory;
		private THCFreeMemoryEvent DOnFreeMemory;
		private THCExecuteTaskEvent DOnExecuteTask;
		private THCGetProgressEvent DOnGetProgress;
		private THCInterruptEvent DOnInterrupt;
		#endregion

		#region メソッド
		public void Dispose() 
		{
			if (this.Initialized) HCInterface.Agent.HCFinalize();
		}

		private Dictionary<int, MatchLine> SearchDocument(string doc, int docId, ref int matchCount, ref double rate, int aMinWords, int aRoughLines, double aRateLevel, bool aIsJp, List<string> lines, List<string> linesIdx)
		{
			Dictionary<int, MatchLine> matchLines = new Dictionary<int, MatchLine>();
			List<int> targetLines = new List<int>();
			string targetDoc = SearchJob.GetTextFileName(doc);
			if (!File.Exists(targetDoc)) return matchLines;
			List<string> paragraphs = this.GetParagraphs(targetDoc, aIsJp);

			matchCount = 0;
			int total = 0;
			int characterCount = 0;
			int pos = 0;
			int linesCount = lines.Count;
#if DEBUG
			if (TEST_MAX_LINES < linesCount) linesCount = TEST_MAX_LINES;
#endif
			if (linesIdx.Count < linesCount) linesCount = linesIdx.Count;
			//this.WriteLog(string.Format("SearchDocument docId={0} linesCount={1}", docId, linesCount));
			List<int> roughMatchLines = new List<int>();
			if (0 < aRoughLines)
			{
				this.WriteLog(string.Format("{0} ラフ検索を開始しました。", doc));
				this.SearchRough(linesIdx, paragraphs, roughMatchLines, aRoughLines);
				this.WriteLog(string.Format("{0} ラフ検索を完了しました。", doc));
			}
			int i = 0;
			try
			{

				for (i = 0; i < linesCount; i++)
				{
					if (string.IsNullOrEmpty(lines[i])) continue;
					if (string.IsNullOrEmpty(linesIdx[i])) continue;
					if (0 == aRoughLines || roughMatchLines.Contains(i))
					{
						int targetLine = 0;
						int totalWords = 0;
						int matchWords = 0;
						pos = 0; // ラフ検索を含めて常に最初から検索する。
						double lineRate = this.SearchLine(lines[i], linesIdx[i], paragraphs, targetLines, i, ref pos, ref total, ref targetLine, ref totalWords, ref matchWords, aMinWords, aRateLevel, aIsJp);
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
			}
			catch (Exception ex)
			{
				this.WriteLog(String.Format("i={0}, linesCount={1},lines.Count={2} linesIdx.Count={3}", i, linesCount, lines.Count, linesIdx.Count));
				this.WriteLog(ex.StackTrace);

			}
			rate = 0 == total ? 0D : (double)matchCount / (double)total;
			//this.WriteLog(String.Format("SearchDocument: docId={0} matchCount={1} total={2} rate={3}", docId, matchCount, total, rate));
			return matchLines;
		}

		public double SearchLine(string line, string lineIdx, List<string> paragraphs, List<int> targetLines, int no, ref int pos, ref int totalCount, ref int targetLine, ref int totalWords, ref int matchWords, int aMinWords, double aRateLevel, bool aIsJp)
		{
			if (string.IsNullOrEmpty(line) || string.IsNullOrEmpty(lineIdx)) return 0;
			string[] words = lineIdx.Split(' ');
			if (words.Length < aMinWords) return 0;
			totalCount++;
			double rate = 0;
			for (int i = pos; i < paragraphs.Count; i++)
			{
				if (string.IsNullOrEmpty(paragraphs[i])) continue;
				if (targetLines.Contains(i))
				{
					continue;
				}
				string src = aIsJp ? lineIdx : line;
				rate = this.GetDiffRate(src, paragraphs[i], ref totalWords, ref matchWords);
				//指定一致率以上であればここで終了。
				if (aRateLevel <= rate && 0 < rate)
				{
					pos = i + 1;
					targetLine = i;
					targetLines.Add(i);
					return rate;
				}
			}
			return 0; //指定一致率以下
		}

		private double GetDiffRate(string src, string target, ref int totalWords, ref int matchWords)
		{
			var d = new Differ();
			var inlineBuilder = new InlineDiffBuilder(d);
			var result = d.CreateWordDiffs(src.Trim(), target.Trim(), true, new char[] { ' ' });
			int diffCount = 0;
			foreach (var block in result.DiffBlocks)
			{
				diffCount += block.DeleteCountA;
			}
			matchWords = result.PiecesOld.Length - diffCount;
			if (matchWords < 0) matchWords = 0;
			totalWords = Math.Max(result.PiecesOld.Length, result.PiecesNew.Length);
			double rate = (double)matchWords / (double)totalWords;
			return rate;
		}

		public void SearchRough(List<string> linesIdx, List<string> paragraphs, List<int> roughMatchLines, int aRoughLines)
		{
			int count = linesIdx.Count;
			int paraCount = paragraphs.Count;
#if DEBUG
			if (TEST_MAX_LINES < count) count = TEST_MAX_LINES;
			if (TEST_MAX_LINES < paraCount) paraCount = TEST_MAX_LINES;
#endif
			StringBuilder sbSrc = new StringBuilder();
			int offset = 0;
			if (aRoughLines < 1)
			{
				aRoughLines = 1;
			}
			List<string> roughpara = new List<string>();
			for (int i = 0; i < paraCount; i += this.TARGET_ROUGH_LINES)
			{
				StringBuilder sb = new StringBuilder();
				for (int j = i; j < i + TARGET_ROUGH_LINES + aRoughLines && j < paraCount; j++)
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
				if (aRoughLines <= offset)
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

		private List<string> GetParagraphs(string fname, bool isJp)
		{
			string line;
			List<string> paragraphs = new List<string>();
			int i = 0;
			using (StreamReader file = new StreamReader(fname))
			{
				while ((line = file.ReadLine()) != null)
				{
					if (isJp) line = TextConverter.SplitWords(line);
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
				try
				{
					BinaryFormatter bf = new BinaryFormatter();
					bf.Serialize(ms, src);
					ret = ms.ToArray();
				}
				catch (Exception e)
				{
					this.WriteLog("オブジェクトのシリアライズに失敗しました: " + e.Message);
				}
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
			HCInterface.Agent.HCReadMemoryEntry(description, (uint)address.ToInt32());
			byte[] bb = new byte[length];
			Marshal.Copy(address, bb, 0, length);
            Marshal.FreeHGlobal(address);
            return bb;
		}

		private void HCWriteMemory(uint description, byte[] bb)
		{
			IntPtr address = Marshal.AllocHGlobal(bb.Length * sizeof(byte));
			Marshal.Copy(bb, 0, address, bb.Length);
			HCInterface.Agent.HCWriteMemoryEntry(description, (uint)address.ToInt32(), (uint)bb.Length);
            Marshal.FreeHGlobal(address);
        }

		private void WriteLog(string Log)
		{
			Debug.WriteLine(Log);
			this.mainForm.WriteLog(Log);
		}

		/* ノードイベント */
		private void DoOnGetMemory(int SlotIndex, uint Size, ref uint Address)
		{
			Address = Convert.ToUInt32(Marshal.AllocHGlobal((int)Size).ToInt32());
		}

		private void DoOnFreeMemory(int SlotIndex, uint Address)
		{
			Marshal.FreeHGlobal(new IntPtr(Convert.ToInt32(Address)));
		}


		private void DoOnExecuteTask(int SlotIndex, uint Description)
		{
			try
			{
				this.WriteLog(string.Format("DoOnExecuteTask SlotIndex={0} Processor={1}", SlotIndex, Process.GetCurrentProcess().ProcessorAffinity.ToInt32()));
				//検索テキストファイル名をHCから取得する
				string textFile = HCInterface.Agent.HCReadAnsiStringEntry(Description);
				//this.WriteLog(string.Format("DoOnExecuteTask SlotIndex={0} textFile={1}", SlotIndex, textFile));
				//検索インデックスファイル名をHCから取得する
				string indexFile = HCInterface.Agent.HCReadAnsiStringEntry(Description);
				//this.WriteLog(string.Format("DoOnExecuteTask SlotIndex={0} indexFile={1}", SlotIndex, indexFile));
				//データパック番号をHCから取得する
				//int index = HCInterface.Agent.HCReadLongEntry(Description);
				int aMinWords = HCInterface.Agent.HCReadLongEntry(Description);
				//string wordCount = HCInterface.Agent.HCReadAnsiStringEntry(Description);
				int aRoughLines = HCInterface.Agent.HCReadLongEntry(Description);
				double aRateLevel = HCInterface.Agent.HCReadDoubleEntry(Description);
				bool aIsJp = HCInterface.Agent.HCReadBoolEntry(Description);
				//this.WriteLog(string.Format("DoOnExecuteTask SlotIndex={0}, index={1}, aMinWords={2}, aRoughLines={3}, aRateLevel={4}, aIsJp={5}", SlotIndex, index, aMinWords, aRoughLines, aRateLevel, aIsJp));
				//バイト配列のサイズをHCから取得する
				int dataPackLength = HCInterface.Agent.HCReadLongEntry(Description);
				int targetDocsLength = HCInterface.Agent.HCReadLongEntry(Description);
				int linesLength = HCInterface.Agent.HCReadLongEntry(Description);
				int linesIndxLength = HCInterface.Agent.HCReadLongEntry(Description);
				//バイト配列をHCから取得する
				byte[] bdataPack = this.HCReadMemory(Description, dataPackLength);
				byte[] btargetDocs = this.HCReadMemory(Description, targetDocsLength);
				byte[] bLines = this.HCReadMemory(Description, linesLength);
				byte[] bLinesIdx = this.HCReadMemory(Description, linesIndxLength);
				//バイト配列を検索結果インスタンスにデシリアライズする
				List<int> dataPack = (this.DeserializeObject(bdataPack)) as List<int>;
				List<string> targetDocs = (this.DeserializeObject(btargetDocs)) as List<string>;
				List<string> lines = (this.DeserializeObject(bLines)) as List<string>;
				List<string> linesIdx = (this.DeserializeObject(bLinesIdx)) as List<string>;
				List<Dictionary<int, MatchLine>> mlList = new List<Dictionary<int, MatchLine>>();
				List<MatchDocument> mdList = new List<MatchDocument>();
                this.WriteLog(String.Format("DoOnExecuteTask: dataPackLength={0} bdataPack.Length={1}, dataPack.Count={2}, targetDocs.Count={2}", dataPackLength, bdataPack.Length, dataPack.Count, targetDocs.Count));
				for (int i = 0; i < dataPack.Count && i < targetDocs.Count; i++)
				{
					int index = dataPack[i];
					string docName = targetDocs[i];

					if (File.Exists(textFile) && File.Exists(indexFile) && File.Exists(docName))
					{
						int matchCount = 0;
						double rate = 0D;
						//検索結果インスタンスを取得する
						Dictionary<int, MatchLine> matchLines = this.SearchDocument(docName, index, ref matchCount, ref rate, aMinWords, aRoughLines, aRateLevel, aIsJp, lines, linesIdx);
						//this.WriteLog(string.Format("DoOnExecuteTask SlotIndex={0} matchLines.Count={1} matchCount=[2]", SlotIndex, matchLines.Count, matchCount));
						MatchDocument md = new MatchDocument(rate, matchCount, docName, index);
						mlList.Add(matchLines);
						mdList.Add(md);
					}
				}
				//検索結果インスタンスをバイト配列にシリアライズする
				byte[] bmlList = this.SerializeObject(mlList);
				byte[] bmdList = this.SerializeObject(mdList);
				//文書番号をHCに登録する
				//HCInterface.Multi.HCWriteLongEntry(Description, index);
				//バイト配列のサイズをHCに登録する
				HCInterface.Agent.HCWriteLongEntry(Description, (int)bdataPack.Length);
				HCInterface.Agent.HCWriteLongEntry(Description, (int)bmlList.Length);
				HCInterface.Agent.HCWriteLongEntry(Description, (int)bmdList.Length);
				//バイト配列をHCに登録する
				this.HCWriteMemory(Description, bdataPack);
				this.HCWriteMemory(Description, bmlList);
				this.HCWriteMemory(Description, bmdList);
				this.WriteLog(string.Format("DoOnExecuteTask Finished SlotIndex={0}", SlotIndex));
                //GC.Collect();
            }
			catch (Exception e)
			{
				this.WriteLog(e.Message + e.StackTrace);
			}
		}


		private void DoOnGetProgress(int SlotIndex, ref double Progress)
		{
			Progress = this.AgentProgressArray[SlotIndex - 1];
			//this.WriteLog(string.Format("DoOnGetProgress SlotIndex={0}, Progress={1}", SlotIndex, Progress));
		}

		private void DoOnInterrupt(int SlotIndex)
		{
			//this.WriteLog("DoOnInterrupt");
			this.AgentInterruptedArray[SlotIndex - 1] = true;
		}

		private void InitializeHC()
		{
			this.WriteLog("並列処理の準備が出来ました。");
			this.Initialized = false;
			HCInterface.Agent.HCInitialize(this.userIndex, this.DOnGetMemory,
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
			if (2 == this.userIndex) HCInterface.Agent.HCSetDebugMode(true);
		}

		#endregion
	}
}
