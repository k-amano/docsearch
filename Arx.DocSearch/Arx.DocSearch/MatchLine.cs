namespace Arx.DocSearch
{
	public class MatchLine
	{
		public double Rate { get; set; }
		public int TargetLine { get; set; }
		public int TotalWords { get; set; }
		public int MatchWords { get; set; }

		public MatchLine()
		{
			this.Rate = 0D;
			this.TargetLine = 0;
			this.TotalWords = 0;
			this.MatchWords = 0;
		}

		public MatchLine(double rate, int targetLine, int totalWords, int matchWords)
		{
			this.Rate = rate;
			this.TargetLine = targetLine;
			this.TotalWords = totalWords;
			this.MatchWords = matchWords;
		}
	}
}
