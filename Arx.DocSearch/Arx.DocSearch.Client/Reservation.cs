namespace Arx.DocSearch.Client
{
	public class Reservation
	{
		public string SrcFile { get; set; }
		public string TargetFolder { get; set; }
		public bool IsJp { get; set; }

		public Reservation()
		{
			this.SrcFile = string.Empty;
			this.TargetFolder = string.Empty;
			this.IsJp = false;
		}

		public Reservation(string srcFile, string targetFolder, bool isJp)
		{
			this.SrcFile = srcFile;
			this.TargetFolder = targetFolder;
			this.IsJp = isJp;
		}
	}
}
