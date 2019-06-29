namespace Interview_Assignment.Controllers
{
	using System;
	using System.Linq;
	using System.Web;
	using System.Web.Mvc;
	using Interview_Assignment.Utility;

	public class HomeController : Controller
	{
		
		[HttpGet]
		public ActionResult Index()
		{
			return View();
		}

		/// <summary>
		/// Get file is converted to JSON or Error when Uploaded File is wrong 
		/// </summary>
		/// <param name="postedFileBase">uploaded Excel or CSV file</param>
		/// <returns> return to View Success</returns>
		[HttpPost]
		public ActionResult Index(HttpPostedFileBase postedFileBase)
		{
			ViewBag.success = null;
			try
			{
				// Check file is existed or not
				if (postedFileBase == null || postedFileBase.ContentLength <= 0)
				{
					ViewBag.error = "The File could be not Excel or CSV file, or empty file. Please Select the correct Excel or CVS  file";
				}
				else
				{
					string[] validFileTypes = { ".xls", ".xlsx", ".csv" };
					string extension = System.IO.Path.GetExtension(postedFileBase.FileName);

					// Check file extension is csv, xls, xlsx
					if (validFileTypes.Contains(extension))
					{
						// Store file in local system
						// The path of converted folder
						string pathconvertedFolder = Server.MapPath("~/Content");
						// The path of uploaded File
						string pathUploadedFile = string.Format("{0}/{1}", pathconvertedFolder, postedFileBase.FileName);
						bool result = false;

						// Remove existed file before create new one
						if (System.IO.File.Exists(pathUploadedFile))
						{
							System.IO.File.Delete(pathUploadedFile);
						}

						postedFileBase.SaveAs(pathUploadedFile);
						ConvertationFiles covertationFiles = new ConvertationFiles();
						if (extension.Equals(".csv"))
						{
							//Convert CSV file to JSON and get result is True if  convertion successes
							result = covertationFiles.ConvertCSVToJSON(pathUploadedFile, pathconvertedFolder, postedFileBase.FileName);
						}
						else
						{
							//Convert EXCEL file to JSON and get result is True if  convertion successes
							result = covertationFiles.ConvertXLSXToJSON(pathUploadedFile, pathconvertedFolder, postedFileBase.FileName);
						}

						if (result)
						{
							ViewBag.success = "Successful upload files! The file will be stored in \" ~/Content ";
						}
					}
					else
					{
						ViewBag.error = "The File could be not Excel or CSV file, or empty file. Please Select the correct Excel or CVS  file";
					}
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				ViewBag.error = "The File could be not Excel or CSV file, or empty file. Please Select the correct Excel or CVS  file";
			}
			
			return View();
		}
	}
}