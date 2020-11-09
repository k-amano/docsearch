using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;
using System.Runtime.Serialization;
using System.Xml;

namespace Arx.DocSearch
{
	public class Log
	{
		#region Constructor
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public Log()
		{
			this.srcFile = "";
			this.targetFolder = "";
			this.listItems = "";
			this.isJp = false;
			this.lineCount = 0;
			this.charCount = 0;
			this.matchLinesTable = new Dictionary<int, Dictionary<int, MatchLine>>();
		}
		#endregion

		#region Field
		private string srcFile;
		private string targetFolder;
		private string listItems;
		private bool isJp;
		private int lineCount;
		private int charCount;
		private Dictionary<int, Dictionary<int, MatchLine>> matchLinesTable;
		#endregion

		#region Property
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

		/// <summary>
		/// ログイン時にアクセスする WebService の URL を取得または設定します。
		/// </summary>
		public string TargetFolder
		{
			get
			{
				return targetFolder;
			}
			set
			{
				targetFolder = value;
			}
		}

		/// <summary>
		/// ログイン時にアクセスする WebService の URL を取得または設定します。
		/// </summary>
		public string ListItems
		{
			get
			{
				return listItems;
			}
			set
			{
				listItems = value;
			}
		}
		/// <summary>
		/// ログイン時にアクセスする WebService の URL を取得または設定します。
		/// </summary>
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
		public int LineCount
		{
			get
			{
				return lineCount;
			}
			set
			{
				lineCount = value;
			}
		}
		public int CharCount
		{
			get
			{
				return charCount;
			}
			set
			{
				charCount = value;
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
		#endregion

		#region Method
		/// <summary>
		/// config ファイルから設定内容を読み込み、その値を書き込んだ Schema インスタンスを取得します。
		/// </summary>
		/// <param name="configFile">設定ファイル。</param>
		/// <returns>取得した FelicaDemo オブジェクト。</returns>
		static public Log LoadSettings(string configFile)
		{
			Log log = new Log();
			try
			{
				if (File.Exists(configFile))
				{
					//＜XMLファイルから読み込む＞
					//XmlSerializerオブジェクトの作成
					DataContractSerializer serializer = new DataContractSerializer(typeof(Log));
					//ファイルを開く
					using (FileStream fs = new FileStream(configFile, FileMode.Open))
						//XMLファイルから読み込み、逆シリアル化する
						log = serializer.ReadObject(fs) as Log;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.ToString());
			}
			return log;
		}

		/// <summary>
		/// config ファイルに設定内容を書き込みます。
		/// </summary>
		/// <param name="configFile">設定ファイル。</param> 
		public void SaveSettings(string configFile)
		{
			try
			{
				//＜XMLファイルに書き込む＞
				//XmlSerializerオブジェクトを作成
				//書き込むオブジェクトの型を指定する
				DataContractSerializer serializer = new DataContractSerializer(typeof(Log));
				//ファイルを開く
				using (FileStream fs = new FileStream(configFile, FileMode.Create))
					//シリアル化し、XMLファイルに保存する
					serializer.WriteObject(fs, this);
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.ToString());
			}
		}
		#endregion
	}
}

