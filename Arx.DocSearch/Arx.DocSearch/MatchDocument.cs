namespace Arx.DocSearch
{
	public class MatchDocument
	{
		public double Rate { get; set; }
		public int MatchCount { get; set; }
		public string Doc { get; set; }
		public int DocId { get; set; }

		public MatchDocument(double rate, int matchCount, string doc, int docId)
		{
			this.Rate = rate;
			this.MatchCount = matchCount;
			this.Doc = doc;
			this.DocId = docId;
		}
	}
}
