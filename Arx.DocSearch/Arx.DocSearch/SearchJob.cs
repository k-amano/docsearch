using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using Xyn.Util;

namespace Arx.DocSearch
{
	public class SearchJob
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
		private string targetFile;
		private int minWords;
		private string wordCount;
		private Dictionary<int, Dictionary<int, MatchLine>> matchLinesTable;
		private bool isJp = false;
		private double rateLevel;
		private readonly int TEST_MAX_LINES = 5000;
		//private readonly int TEST_MAX_LINES = 200;
		private readonly int ROUGH_COUNT = 20;
		private readonly double ROUGH_RATE = 0.5;
		private readonly int TARGET_ROUGH_LINES = 50;
		//private readonly int LINE_LENGHTH = 70;
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

		public string TargetFile
		{
			get
			{
				return targetFile;
			}
			set
			{
				targetFile = value;
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
		public bool StartSearch()
		{
			if (!File.Exists(this.SrcFile) || !Directory.Exists(this.TargetFile))
			{
				return false;
			}
			string textFile = SearchJob.GetTextFileName(this.SrcFile);
			string indexFile = SearchJob.GetIndexFileName(this.SrcFile);
			this.MinWords = ConvertEx.GetInt(this.WordCount);
			if (string.IsNullOrEmpty(indexFile))
			{
				MessageBox.Show("検索元のインデックスファイルを作成してください。");
				return false;
			}
			this.mainForm.Invoke(
				(MethodInvoker)delegate()
				{
					this.mainForm.ClearListView();
				}
			);

			this.lines.Clear();
			using (StreamReader file = new StreamReader(textFile))
			{
				string line;
				while ((line = file.ReadLine()) != null)
				{
					//if (string.IsNullOrEmpty(line)) continue;
					this.lines.Add(line);
				}
			}
			this.linesIdx.Clear();
			using (StreamReader file = new StreamReader(indexFile))
			{
				string line;
				while ((line = file.ReadLine()) != null)
				{
					this.linesIdx.Add(line);
				}
			}
			this.MatchLinesTable.Clear();
			for (int j = 0; j < this.docs.Count; j++)
			{
				this.mainForm.UpdateCountLabel(string.Format("{0} 文書中 {1} 文書目。開始 {2} 終了 {3}。", docs.Count, j + 1, this.startTime.ToLongTimeString(), DateTime.Now.ToLongTimeString()));
				int matchCount = 0;
				double rate = 0D;
				Dictionary<int, MatchLine> matchLines = this.SearchDocument(this.docs[j], j, ref matchCount, ref rate);
				this.matchList.Add(new MatchDocument(rate, matchCount, this.docs[j], j));
				this.MatchLinesTable.Add(j, matchLines);
				if (0 < this.matchList.Count)
				{
					MatchDocument md = this.matchList[this.matchList.Count - 1];
					this.mainForm.updateListView(new MatchDocument[] { md }, false);
				}
			}
			// リストをID順でソートする
			MatchDocument[] matchArray = this.matchList.ToArray();
			Array.Sort(matchArray, (a, b) => (int)((a.Rate - b.Rate) * -1000000));
			this.mainForm.updateListView(matchArray, true);
			this.mainForm.FinishSearch(string.Format("{0} 文書 {0} 文書目。開始 {1} 終了 {2}。", docs.Count, this.startTime.ToLongTimeString(), DateTime.Now.ToLongTimeString()), this.matchLinesTable, this.srcFile);
			return true;
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
				this.mainForm.UpdateMessageLabel(string.Format("{0} ラフ検索中です。", doc));
				this.SearchRough(this.linesIdx, paragraphs, roughMatchLines);
			}
			for (int i = 0; i < linesCount; i++)
			{
				if (string.IsNullOrEmpty((lines[i]))) continue;
				this.mainForm.UpdateMessageLabel(string.Format("{0} {1}行目を検索中です。", doc, i + 1));
				if (0 == this.roughLines || roughMatchLines.Contains(i))
				{
					//if (72==i) Debug.WriteLine("### i={0}\n{1}", i, lines[i]);
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
            if (this.isJp && string.IsNullOrEmpty(lineIdx)) return 0;
            string[] words = lineIdx.Split(' ');
			if (words.Length < this.minWords) return 0;
			totalCount++;
			double rate = 0;
			for (int i = pos; i < paragraphs.Count; i++)
			{
				//if (45 == no && 44 == i) Debug.WriteLine("### SearchDocument i= {0}", i);
				if (string.IsNullOrEmpty(paragraphs[i])) continue;
				if (targetLines.Contains(i))
				{
					//if (45 == no && 44 == i) Debug.WriteLine("ContainsKey");
					continue;
				}
				string src = this.isJp ? lineIdx : line;
				rate = this.GetDiffRate(src, paragraphs[i], ref totalWords, ref matchWords);
				//if (45==no && 44==i) Debug.WriteLine("### paragraphs[{0}] rate ~= {1}\n{2}", i, rate, paragraphs[i]);
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
			//wordCount = result.PiecesOld.Length;
			foreach (var block in result.DiffBlocks)
			{
				//diffCount += block.DeleteCountA + block.InsertCountB;
				diffCount += block.DeleteCountA;
			}
			matchCount = result.PiecesOld.Length - diffCount;
			if (matchCount < 0) matchCount = 0;
			wordCount = Math.Max(result.PiecesOld.Length, result.PiecesNew.Length);
			double rate = (double)matchCount / (double)wordCount;
            //if (0.7 < rate && rate < 1.0) Debug.WriteLine(string.Format("src={0}\ntarget={1}", src, target));
            //if (0.4 < rate && src.Contains("operably coupled")) Debug.WriteLine(string.Format("src={0}\ntarget={1}\n matchCount={2}", src, target, matchCount));
            //Debug.WriteLine(string.Format("src={0}\ntarget={1}\n matchCount={2}", src, target, matchCount));
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
						//if (72 == i) Debug.WriteLine(string.Format("i={0}\n{1}\n{2}", i, sbSrc, roughpara[j]));
						if (this.GetRoughRate(words, roughpara[j]))
						{
							//if (72 == i) Debug.WriteLine("###Matched");
							for (int k = i - offset; k <= i; k++)
							{
								if (!roughMatchLines.Contains(k)) roughMatchLines.Add(k);
								//if (72 == i) Debug.WriteLine(String.Format("###Matched k={0}", k));
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
			//Debug.WriteLine(string.Format("##paragraph={0}", paragraph));
			List<int> ls = this.GetRandom(words.Length);
			int pos = 0;
			int matchCount = 0;
			for (int i = 0; i < ls.Count; i++)
			{
				int j = ls[i];
				int newPos = paragraph.IndexOf(words[j], pos);
				//Debug.WriteLine(string.Format("### j={0} words[j]={1} pos={2} newPos={3}", j, words[j], pos, newPos));
				if (pos <= newPos)
				{
					matchCount += 1;
					pos = newPos + 1;
				}
				//20 word 進んで一致率70%未満は終了
				if ((ls.Count <= i + 1 || this.ROUGH_COUNT * this.ROUGH_RATE < i) && matchCount < i * this.ROUGH_RATE)
				{
					//if ("患者の" == words[0]) Debug.WriteLine(string.Format("### not matched i={0} matchCount={1}", i, matchCount));
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
			//if (File.Exists(indexFile) && File.GetLastWriteTime(srcFile) <= File.GetLastWriteTime(indexFile)) return indexFile;
			if (File.Exists(indexFile)) return indexFile;
			return string.Empty;
		}

		private List<string> GetParagraphs(string fname)
		{
            //Debug.WriteLine("#GetParagraphs:" + fname);
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
		#endregion
	}
}

