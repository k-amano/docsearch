using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;

namespace Arx.DocSearch.MultiCore
{
	public class Schema
	{
		#region Constructor
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public Schema()
		{
			this.srcFile = "";
			this.targetFolder = "";
			this.rate = "60";
			this.wordCount = "10";
			this.xlsdir = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
		}
		#endregion

		#region Field
		private string srcFile;
		private List<string> srcFiles;
		private string targetFolder;
		private string rate;
		private string wordCount;
		private string roughLines;
		private string xlsdir;
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


		public List<string> SrcFiles
		{
			get
			{
				return srcFiles;
			}
			set
			{
				srcFiles = value;
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
		public string Rate
		{
			get
			{
				return rate;
			}
			set
			{
				rate = value;
			}
		}

		/// <summary>
		/// ログイン時にアクセスする WebService の URL を取得または設定します。
		/// </summary>
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

		/// <summary>
		/// ログイン時にアクセスする WebService の URL を取得または設定します。
		/// </summary>
		public string RoughLines
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


		/// <summary>
		/// ログイン時にアクセスする Web ページの URL を取得または設定します。
		/// </summary>
		public string Xlsdir
		{
			get
			{
				return xlsdir;
			}
			set
			{
				xlsdir = value;
			}
		}

		#endregion

		#region Method
		/// <summary>
		/// config ファイルから設定内容を読み込み、その値を書き込んだ Schema インスタンスを取得します。
		/// </summary>
		/// <param name="configFile">設定ファイル。</param>
		/// <returns>取得した FelicaDemo オブジェクト。</returns>
		static public Schema LoadSettings(string configFile)
		{
			Schema schema = new Schema();
			try
			{
				if (File.Exists(configFile))
				{
					//＜XMLファイルから読み込む＞
					//XmlSerializerオブジェクトの作成
					XmlSerializer serializer = new XmlSerializer(typeof(Schema));
					//ファイルを開く
					using (FileStream fs = new FileStream(configFile, FileMode.Open))
						//XMLファイルから読み込み、逆シリアル化する
						schema = serializer.Deserialize(fs) as Schema;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.Message + ex.StackTrace);
			}
			return schema;
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
				XmlSerializer serializer = new XmlSerializer(typeof(Schema));
				//ファイルを開く
				using (FileStream fs = new FileStream(configFile, FileMode.Create))
					//シリアル化し、XMLファイルに保存する
					serializer.Serialize(fs, this);
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.ToString());
			}
		}
		#endregion
	}
}
