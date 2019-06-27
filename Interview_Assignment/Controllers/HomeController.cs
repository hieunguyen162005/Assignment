using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelDataReader;

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
				// check file is existed or not
				if (file == null || file.ContentLength <= 0)
				{
					ViewBag.error = "Please Select the Excel or CVS  file";
				}
				else
				{
					string[] validFileTypes = { ".xls", ".xlsx", ".csv" };
					string extension = System.IO.Path.GetExtension(file.FileName);

					// check file extension is csv, xls, xlsx
					if (validFileTypes.Contains(extension))
					{
						// store file in local system
						string pathFile = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), file.FileName);
						Boolean result = false;
						// remove existed file before create new one
						if (System.IO.File.Exists(pathFile))
						{
							System.IO.File.Delete(pathFile);
						}
						file.SaveAs(pathFile);
						//convert CSV file to JSON and get result is True if  convert success
						if (extension == ".csv")
						{
							result = ConvertCSVToJSON(pathFile, Server.MapPath("~/Content/Uploads"), file.FileName);
						}
						else
						{
							result = ConvertXLSXToJSON(pathFile, Server.MapPath("~/Content/Uploads"), file.FileName);
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
				List<Dictionary<object, object>> fileList = new List<Dictionary<object, object>>();
				string[] headerArray = sr.ReadLine().Split(',');
				
				while (!sr.EndOfStream)
				{
					string[] line = sr.ReadLine().Split(',');
					Dictionary<object, object> dict = new Dictionary<object, object>();
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

		private Boolean ConvertXLSXToJSON(string pathFile, string pathFolder, string fileName)
		{
			using (var stream = System.IO.File.Open(pathFile, FileMode.Open, FileAccess.Read))
			{
				using (var reader = ExcelReaderFactory.CreateReader(stream))
				{
					var dataset = reader.AsDataSet();
					var table = dataset.Tables[0];
					System.Data.DataColumnCollection columns = table.Columns;
					object[] headerArray =  table.Rows[0].ItemArray;
					
					List<Dictionary<object, object>> fileList = new List<Dictionary<object, object>>();
					for(int rowIndex = 1; rowIndex < table.Rows.Count;rowIndex++)
					{
						object[] line = table.Rows[rowIndex].ItemArray;
						Dictionary<object, object> dict = new Dictionary<object, object>();
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
					
					reader.Close();
					return SaveAsJson(fileList, pathFolder, fileName);
				}

			}

		}
		private Boolean SaveAsJson(List<Dictionary<object, object>> fileList, string pathFolder, string fileName)
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