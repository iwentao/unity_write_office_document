using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI;

public class OpWord : MonoBehaviour
{
	/// <summary>
	/// 存放word中需要替换的关键字以及对应要更改的内容
	/// </summary>
	public Dictionary<string, string> DicWord = new Dictionary<string, string>();

	private string path = Application.streamingAssetsPath + "/1.docx";
	private string targetPath = Application.streamingAssetsPath + "/3.docx";

	public Text text;

	private void Start()
	{
		if (!File.Exists(path))
		{
			File.Create(path).Dispose();
		}

		DicWord.Add("     1", "00000");
		DicWord.Add("     2", "11111");
		DicWord.Add("     3", "33333");
		DicWord.Add("   0  ", "oh God");
		//Read();
		//WriteDoc();
		ReplaceKeyword();
	}



	private void ReplaceKeyword()
	{
		using (FileStream stream = File.OpenRead(path))
		{
			FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
			XWPFDocument doc = new XWPFDocument(fs);

			foreach (var para in doc.Paragraphs)
			{
				string oldText = para.ParagraphText;
				if (oldText != "" && oldText != string.Empty && oldText != null)
				{
					string tempText = para.ParagraphText;

					foreach (KeyValuePair<string, string> kvp in DicWord)
					{                     
						if (tempText.Contains(kvp.Key))
						{
							tempText = tempText.Replace(kvp.Key, kvp.Value);                         
						}
					}
					para.ReplaceText(oldText, tempText);
					Debug.Log(tempText);
					Debug.Log(para.ParagraphText);
				}
			}

			//遍历表格      
			var tables = doc.Tables;
			foreach (var table in tables)
			{
				foreach (var row in table.Rows)
				{
					foreach (var cell in row.GetTableCells())
					{
						foreach (var para in cell.Paragraphs)
						{
							string oldText = para.ParagraphText;
							if (oldText != "" && oldText != string.Empty && oldText != null)
							{
								//记录段落文本
								string tempText = para.ParagraphText;
								foreach (KeyValuePair<string, string> kvp in DicWord)
								{
									if (tempText.Contains(kvp.Key))
									{
										tempText = tempText.Replace(kvp.Key, kvp.Value);

										//替换内容
										para.ReplaceText(oldText, tempText);
									}
								}
							}
						}
					}
				}
			}

			//生成指定文件
			//XWPFDocument doc0 = doc;
			FileStream output = new FileStream(targetPath, FileMode.Create);
			//将文档信息写入文件
			doc.Write(output);

			//关闭释放
			fs.Close();
			fs.Dispose();
			output.Close();
			output.Dispose();

			System.Diagnostics.Process.Start(targetPath);//打开指定文件
		}
	}


	private void WriteDoc()
	{
		FileStream file = new FileStream(targetPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);

		XWPFDocument doc = new XWPFDocument(file);

		XWPFParagraph p2 = doc.CreateParagraph();
		p2.Alignment = ParagraphAlignment.CENTER;//居中

		XWPFRun r2 = p2.CreateRun();   //插入一行
		r2.SetText("000casdasda000sdasd");
		//r2.SetColor("00FFFF");
		//r2.SetTextPosition(1);
		//r2.SetUnderline(UnderlinePatterns.Single);
		//r2.SetStrike(value: true);
		//r2.SetFontFamily("宋体", FontCharRange.None);

		//插入一张照片
		XWPFParagraph p3 = doc.CreateParagraph();
		p3.Alignment = ParagraphAlignment.CENTER;
		XWPFRun r6 = p3.CreateRun();
		XWPFTable table = doc.CreateTable(1, 4);//创建1*4的表

		table.SetColumnWidth(0, 6 * 256);
		table.SetColumnWidth(1, 10 * 256);
		table.SetColumnWidth(2, 6 * 256);
		table.SetColumnWidth(3, 10 * 256);


		table.GetRow(0).GetCell(0).SetText("11111");
		table.GetRow(0).GetCell(0).SetColor("00FFFF");
		table.GetRow(0).GetCell(1).SetText("11111");
		table.GetRow(0).GetCell(2).SetText("11111");
		table.GetRow(0).GetCell(3).SetText("11111");
		table.GetRow(0).GetCell(3).SetColor("00FFFF");


		doc.Write(file);
		file.Close();
		System.Diagnostics.Process.Start(targetPath);
	}


	void Read()
	{
		FileStream file = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);
		XWPFDocument word = new XWPFDocument(file);

		Debug.Log(word.Paragraphs.Count);
		foreach (XWPFParagraph paragraph in word.Paragraphs)
		{
			string message = paragraph.ParagraphText;//获取段落内容
			Debug.Log(message);
			text.text += message;
		}
		file.Close();
	}
}
