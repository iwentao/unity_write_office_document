/// --------------------------------------------- ///
/// Word manager
/// -- auxiliary class to write article.
/// Date: 2022-10-30
/// --------------------------------------------- ///

using System.IO;
using NPOI.XWPF.UserModel;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class WordManager : Singleton<WordManager>
{
	/// <summary>
	/// Target path for office document. 
	/// i.e. C:/a/b/c.docx
	/// </summary>
    public string TargetPath;

	/// <summary>
	/// Source path for document to be replaced. 
	/// </summary>
    public string SourcePath;

	/// <summary>
	/// Replace table used to replace old text in word template with given new text. 
	/// Need to be initialized before usage.
	/// /// </summary>
	public Dictionary<string, string> ReplaceTable;

	public bool AutoOpenDocument = false;

	public const string Default_Font_Color = "000000";
	public const string Font_Color_Red = "FF0000";
	public const string Font_Color_Green = "00FF00";
	public const string Font_Color_Blue = "0000FF";

	public const string Default_Font_Famility = "宋体";
	public const int Default_Font_Size = 16;

	public const int Default_Font_Paragraph_Size = 14;

	private XWPFDocument doc = new XWPFDocument();

	/// <summary>
	/// Re-generate XWPFDocument, it means a new office document.
	/// </summary>
	public void RegenDocument()
	{
		if(doc != null)
		{
			doc.Close();
			doc = null;
		}
		doc = new XWPFDocument();
	}

	public void WriteArticle(string[] content, bool autoSave = true, bool addWhiteSpaceBeforeParagraph = false)
	{
		WriteArticle_Impl(content, autoSave, addWhiteSpaceBeforeParagraph);
	}

	public void WriteArticle(string content, char splitter, bool autoSave = true, bool addWhiteSpaceBeforeParagraph = false)
	{
		string[] cc = content.Split(splitter);
		WriteArticle_Impl(cc, autoSave, addWhiteSpaceBeforeParagraph);
	}

	private void WriteArticle_Impl(string[] contents, bool autoSave, bool addWhiteSpaceBeforeParagraph)
	{
		if(contents == null || contents.Length == 0)
		{
			Debug.LogError("Empty article");
			return;
		}

		foreach(var c in contents)
			AddParagraph(c, ParagraphAlignment.LEFT, Default_Font_Color, Default_Font_Paragraph_Size, false, addWhiteSpaceBeforeParagraph);

		if(autoSave)
			Save();
	}

	public void AddTitle(string title, string font_color = Default_Font_Color, int font_size = Default_Font_Size, bool isBold = true)
	{
		AddParagraph(title, ParagraphAlignment.CENTER, font_color, font_size, isBold);
	}

	public void AddParagraph(string paragraph_text, ParagraphAlignment alignment, string font_color = Default_Font_Color, int font_size = Default_Font_Size, bool isBold = false, bool addWhiteSpace = false)
	{
		XWPFParagraph p = doc.CreateParagraph();
		p.Alignment = alignment;

		XWPFRun run = p.CreateRun();
		string text = addWhiteSpace ? "    " + paragraph_text : paragraph_text;
		SetRun(run, text , font_color , Default_Font_Famility, font_size , false, isBold);
	}

	public void AddTable(int row_count, int column_count, float[] column_width_array, string[] table_content)
	{
		Debug.Assert(table_content.Length == row_count * column_count);
		Debug.Assert(column_width_array.Length == column_count);

		XWPFParagraph p3 = doc.CreateParagraph();
		p3.Alignment = ParagraphAlignment.CENTER;
		XWPFRun r6 = p3.CreateRun();
		XWPFTable table = doc.CreateTable(row_count, column_count);

		for(int i = 0; i < column_count; ++i)
			table.SetColumnWidth(i, (ulong)column_width_array[i] * 256);

		for(int r = 0; r < row_count; ++r)
			for(int c = 0; c < column_count; ++c)
				table.GetRow(r).GetCell(c).SetText(table_content[r * column_count + c]);

		// table.SetColumnWidth(0, 6 * 256);
		// table.SetColumnWidth(1, 10 * 256);
		// table.SetColumnWidth(2, 6 * 256);
		// table.SetColumnWidth(3, 10 * 256);

		// table.GetRow(0).GetCell(0).SetText("11111");
		// table.GetRow(0).GetCell(0).SetColor("00FFFF");
		// table.GetRow(0).GetCell(1).SetText("11111");
		// table.GetRow(0).GetCell(2).SetText("11111");
		// table.GetRow(0).GetCell(3).SetText("11111");
		// table.GetRow(0).GetCell(3).SetColor("00FFFF");
	}

	public void AddImage(string image_folder, string image_name, float image_size_scale)
	{
		XWPFParagraph paragraph = doc.CreateParagraph();
		paragraph.Alignment = ParagraphAlignment.CENTER;

		XWPFRun run = paragraph.CreateRun();
		SetImage(run, image_size_scale, image_folder, image_name);
	}

	public void AddImages(string image_folder, string[] image_names, float[] image_size_scales)
	{
		Debug.Assert(image_names != null && image_names.Length > 0);
		Debug.Assert(image_size_scales != null && image_size_scales.Length > 0);
		Debug.Assert(image_names.Length == image_size_scales.Length);

		for(int i =0; i < image_names.Length; ++i)
		{
			XWPFParagraph paragraph = doc.CreateParagraph();
			paragraph.Alignment = ParagraphAlignment.CENTER;

			XWPFRun run = paragraph.CreateRun();
			SetImage(run, image_size_scales[i], image_folder, image_names[i]);
		}
	}

	public void Save()
	{
		FileStream fs = new FileStream(TargetPath, FileMode.Create);
		doc.Write(fs);
		fs.Close();
		fs.Dispose();
		Debug.Log($"Created {TargetPath} succeed");
	}

	public void ReplaceWord(bool autoSave = false)
	{
		Debug.Log("Regenerate a doc to replace words");
		if(doc != null)
		{
			doc.Close();
			doc = null;
		}

		FileStream fs = new FileStream(SourcePath, FileMode.Open, FileAccess.Read);
		doc = new XWPFDocument(fs);

		XWPFDocument replace_doc = doc;

		foreach (var para in replace_doc.Paragraphs)
		{
			string oldText = para.ParagraphText;
			if (oldText != "" && oldText != string.Empty && oldText != null)
			{
				string tempText = para.ParagraphText;

				foreach (KeyValuePair<string, string> kvp in ReplaceTable)
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
		var tables = replace_doc.Tables;
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
							foreach (KeyValuePair<string, string> kvp in ReplaceTable)
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

		fs.Close();
		fs.Dispose();

		if(autoSave)
			Save();

		if(AutoOpenDocument)
			System.Diagnostics.Process.Start(TargetPath);
	}

	public void Test_Write2()
	{
		string _content = "test words";
		XWPFParagraph paragraph = doc.CreateParagraph(); 
		paragraph.Alignment = ParagraphAlignment.CENTER;
		paragraph.SetNumID("1");

		XWPFRun run = paragraph.CreateRun();
		run.FontSize = 20;
		run.SetColor("33CC00");
		run.SetText(_content);

		FileStream fs = new FileStream(TargetPath, FileMode.Create);
		doc.Write(fs);
		fs.Close();
		fs.Dispose();
		Debug.Log($"Created {TargetPath} succeed");
	}

	public void Test_Write3()
	{
		XWPFParagraph paragraph = doc.CreateParagraph();
		paragraph.Alignment = ParagraphAlignment.CENTER;

		XWPFRun run1 = paragraph.CreateRun();
		SetRun(run1, "Hello 测试", "FF0000", "宋体", 16, false, true);
		// SetImage(run1, 100.0f, 100.0f, Application.dataPath + "/../test_images", "test3.png");

		FileStream fs = new FileStream(TargetPath, FileMode.Create);
		doc.Write(fs);
		fs.Close();
		fs.Dispose();
		Debug.Log($"Created {TargetPath} succeed");
	}

	/// <summary>
	/// Add image to paragraph.
	/// </summary>
	/// <param name="run">The paragraph to insert image</param>
	/// <param name="scale">The scale factor that multple the original size</param>
	/// <param name="image_folder">Images folder</param>
	/// <param name="image_name">The image name to insert</param>
	private void SetImage(XWPFRun run, float scale, string image_folder, string image_name)
	{
		// int width = (int)(sizeX * 9525);
		// int height = (int)(sizeY * 9525);
		try
		{
			string image_path = System.IO.Path.Combine(image_folder, image_name);
			if(File.Exists(image_path))
			{
				byte[] bytes;
				Vector2 image_size;
				FileUtility.FileInfo(image_path, out bytes, out image_size);
				Debug.Log($"image size, width = {image_size.x}, height = {image_size.y}");

				using(FileStream fs = new FileStream(image_path, FileMode.Open, FileAccess.Read))
					run.AddPicture(fs, (int)PictureType.PNG, image_name, (int)(image_size.x * 9525 * scale), (int)(image_size.y * 9525 * scale));
			}
			else Debug.Log($"{image_path} is not exist");
		}
		catch
		{
			Debug.LogError("write image failed");
		}
	}

	private void SetRun(XWPFRun run, string text, string color, string fontfamily, int size, bool isItalic, bool isBold)
	{
		run.SetText(text);
		run.FontFamily = fontfamily;
		run.FontSize = size;
		run.SetColor(color);
		run.IsItalic = isItalic;
		run.IsBold = isBold;
	}

    public void Test_Write()
    {
        string targetPath = TargetPath;
		Debug.Log($"tar path = {targetPath}");

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
		// XWPFParagraph p3 = doc.CreateParagraph();
		// p3.Alignment = ParagraphAlignment.CENTER;
		// XWPFRun r6 = p3.CreateRun();
		// XWPFTable table = doc.CreateTable(1, 4);//创建1*4的表

		// table.SetColumnWidth(0, 6 * 256);
		// table.SetColumnWidth(1, 10 * 256);
		// table.SetColumnWidth(2, 6 * 256);
		// table.SetColumnWidth(3, 10 * 256);


		// table.GetRow(0).GetCell(0).SetText("11111");
		// table.GetRow(0).GetCell(0).SetColor("00FFFF");
		// table.GetRow(0).GetCell(1).SetText("11111");
		// table.GetRow(0).GetCell(2).SetText("11111");
		// table.GetRow(0).GetCell(3).SetText("11111");
		// table.GetRow(0).GetCell(3).SetColor("00FFFF");

		doc.Write(file);
		file.Close();
		file.Dispose();
		// System.Diagnostics.Process.Start(targetPath);
    }
}
