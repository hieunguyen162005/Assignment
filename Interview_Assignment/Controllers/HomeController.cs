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
		/// return to View Success when File is converted to JSON or Error when Uploaded File is wrong 
		/// </summary>
		/// <param name="postedFileBase">uploaded Excel or CSV file</param>
		/// <returns></returns>
		[HttpPost]
		public ActionResult Index(HttpPostedFileBase postedFileBase)
		{
			ViewBag.success = null;
			try
			{
				// check file is existed or not
				if (postedFileBase == null || postedFileBase.ContentLength <= 0)
				{
					ViewBag.error = "Please Select the Excel or CVS  file";
				}
				else
				{
					string[] validFileTypes = { ".xls", ".xlsx", ".csv" };
					string extension = System.IO.Path.GetExtension(postedFileBase.FileName);

					// check file extension is csv, xls, xlsx
					if (validFileTypes.Contains(extension))
					{
						// store file in local system
						// the path of converted folder
						string pathconvertedFolder = Server.MapPath("~/Content");
						// the path of uploaded File
						string pathUploadedFile = string.Format("{0}/{1}", pathconvertedFolder, postedFileBase.FileName);
						bool result = false;

						// remove existed file before create new one
						if (System.IO.File.Exists(pathUploadedFile))
						{
							System.IO.File.Delete(pathUploadedFile);
						}
						postedFileBase.SaveAs(pathUploadedFile);

						
						ConvertationFiles covertationFiles = new ConvertationFiles();
						if (extension == ".csv")
						{
							//convert CSV file to JSON and get result is True if  convert success
							result = covertationFiles.ConvertCSVToJSON(pathUploadedFile, pathconvertedFolder, postedFileBase.FileName);
						}
						else
						{
							//convert EXCEL file to JSON and get result is True if  convert success
							result = covertationFiles.ConvertXLSXToJSON(pathUploadedFile, pathconvertedFolder, postedFileBase.FileName);
						}

						if (result)
						{
							ViewBag.success = "Successful upload files! The file will be stored in \" ~/Content ";
						}
					}
					else
					{
						ViewBag.error = "Please Select the Excel or CVS  file";
					}
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				ViewBag.error = "Please Select the Excel or CVS  file";
			}
			
			return View();
		}

		
	}
}