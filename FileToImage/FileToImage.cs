using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;


namespace FileToImage
{
    public class FileTransform
    {
		public void ExcelToJPG(string sFileName)
		{
			try
			{
				Excel.Application xlsApp = new Excel.Application();
				Excel.Workbook xlsBook = xlsApp.Workbooks.Open(sFileName);
				xlsBook.SaveAs(Application.StartupPath + @"\" + sFileName, Excel.XlFileFormat.xlHtml);
				xlsBook.Close(false, Type.Missing, Type.Missing);
				xlsApp.Quit();
			}
			catch (Exception ex)
			{
				throw new Exception(ex.StackTrace.ToString());
			}

		}

		public void WordToJPG(string sPath)
		{
			Word.Application docApp = new Word.Application();
			Word.Document doc = docApp.Documents.Open(@"D:\ImgW.docx");
			doc.SaveAs2(Application.StartupPath + @"\ImgW.html", Word.WdSaveFormat.wdFormatHTML);
			doc.Close();
			docApp.Quit();
		}

		//xlsbook.LoadFromFile(@"D:\Images.xlsx");
		

		

    }
}
