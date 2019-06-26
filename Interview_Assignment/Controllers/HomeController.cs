using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.OleDb;

namespace Interview_Assignment.Controllers
{
	public class HomeController : Controller
	{
		[HttpGet]
		public ActionResult Index()
		{
			return View();
		}

		
		[HttpPost]
		public ActionResult Index(HttpPostedFileBase file)
		{
			ViewBag.success = null;
			try
			{
				if (file.ContentLength <= 0)
				{
					ViewBag.error = "Please Select the Excel or CVS  file";
				}
				else
				{
					string[] validFileTypes = { ".xls", ".xlsx", ".csv" };
					string extension = System.IO.Path.GetExtension(file.FileName);
					if (validFileTypes.Contains(extension))
					{
						string pathFile = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), file.FileName);
						Boolean result = false;
						if (System.IO.File.Exists(pathFile))
						{
							System.IO.File.Delete(pathFile);
						}
						file.SaveAs(pathFile);
						if (extension == ".csv")
						{
							result = ConvertCSVToJSON(pathFile, Server.MapPath("~/Content/Uploads"), file.FileName);
						}
						else if (extension.Trim() == ".xls")
						{
							ConvertXLSToJSON(pathFile);
						}
						else
						{
							ConvertXLSXToJSON(pathFile);
						}

						if (result)
						{
							ViewBag.success = "Successful upload files! The file will be stored in \" ~/Content/Uploads\" ";
						}
					}
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}
			
			return View();
		}

		private Boolean ConvertCSVToJSON(string pathFile,string pathFolder,string fileName)
		{
			using (StreamReader sr = new StreamReader(pathFile))
			{
				List<Dictionary<string,string>> fileList = new List<Dictionary<string, string>>();
				string[] headerArray = sr.ReadLine().Split(',');
				
				while (!sr.EndOfStream)
				{
					string[] line = sr.ReadLine().Split(',');
					Dictionary<string, string> dict = new Dictionary<string, string>();
					for (int i = 0; i < headerArray.Length; i++)
					{
						if (line.Length <= i)
						{
							dict.Add(headerArray[i], null);
						}
						else
						{
							dict.Add(headerArray[i], line[i]);
						}
						
					}
					fileList.Add(dict);
				}
				return SaveAsJson(fileList,pathFolder, fileName);
			}
		}

		private void ConvertXLSToJSON(string pathFile)
		{
			string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathFile + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
			OleDbConnection oleDbConnection = new OleDbConnection(connString);
			try
			{
				oleDbConnection.Open();
				using (OleDbCommand cmd = new OleDbCommand())
				{

				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}
			finally
			{
				oleDbConnection.Close();
			}
		}

		private void ConvertXLSXToJSON(string pathFile)
		{
			//Excel.Application application = new Excel.Application();
			//Excel.Workbook workbook = application.Workbooks.Open(pathFile);
			//Excel.Worksheet worksheet = workbook.ActiveSheet;
			//Excel.Range range = worksheet.UsedRange;
			//foreach (var a in range)
			//{
			//	Console.WriteLine(a);
			//}

		}
		private Boolean SaveAsJson(List<Dictionary<string, string>> fileList, string pathFolder, string fileName)
		{
			try
			{
				string json = JsonConvert.SerializeObject(fileList);
				System.IO.File.WriteAllText(pathFolder + "\\json_" + fileName.Split('.')[0]+ ".json", json);
				return true;
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}
			return false;
		}
	}
}