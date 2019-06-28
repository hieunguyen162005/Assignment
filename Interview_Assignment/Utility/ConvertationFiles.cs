namespace Interview_Assignment.Utility
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Data;
	using ExcelDataReader;
	using Newtonsoft.Json;

	/// <summary>
	/// Convert files to other files
	/// </summary>
	public class ConvertationFiles
	{
		/// <summary>
		/// Convert CSV file to JSON file
		/// </summary>
		/// <param name="pathFile"> the path of uploaded file in local system</param>
		/// <param name="pathconvertedFolder"> The path of folder for converted file</param>
		/// <param name="fileName"> the uploaded file name</param>
		/// <returns>return True if convertation is successful</returns>
		public bool ConvertCSVToJSON(string pathFile, string pathconvertedFolder, string uploadedFileName)
		{
			using (StreamReader streamReader = new StreamReader(pathFile))
			{
				// list stores every data from uploaded file
				List<Dictionary<object, object>> list = new List<Dictionary<object, object>>();
				// the header of CSV file 
				string[] headerArray = streamReader.ReadLine().Split(',');

				// each line of uploaded file will be stored as dictionary and store in list
				while (!streamReader.EndOfStream)
				{
					string[] line = streamReader.ReadLine().Split(',');
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
					list.Add(dict);
				}
				return SaveAsJson(list, pathconvertedFolder, "FormCSV_"+ uploadedFileName.Split('.')[0] + ".json");
			}
		}

		/// <summary>
		/// Convert EXCEL file ( including ".xlsx", and ".xls" ) to JSON
		/// </summary>
		/// <param name="pathFile"> the path of uploaded file in local system</param>
		/// <param name="pathconvertedFolder"> The path of folder for converted file</param>
		/// <param name="fileName"> the uploaded file name</param>
		/// <returns>return True if convertation is successful</returns>
		public bool ConvertXLSXToJSON(string pathFile, string pathconvertedFolder, string uploadedFileName)
		{
			using (FileStream stream = System.IO.File.Open(pathFile, FileMode.Open, FileAccess.Read))
			{
				// using excel package to read data in uploaded file and store it in dataset
				// each row of dataset will be stored as dictionary and store in list
				using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
				{
					DataSet dataset = reader.AsDataSet();
					DataTable table = dataset.Tables[0];
					System.Data.DataColumnCollection columns = table.Columns;
					object[] headerArray = table.Rows[0].ItemArray;

					List<Dictionary<object, object>> list = new List<Dictionary<object, object>>();
					for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
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
						list.Add(dict);
					}

					reader.Close();
					return SaveAsJson(list, pathconvertedFolder, "FromExcel_" + uploadedFileName.Split('.')[0] + ".json");
				}

			}

		}

		/// <summary>
		/// Save Excel or CSV file as JSOn in specific folder
		/// </summary>
		/// <param name="pathFile"> the path of uploaded file in local system</param>
		/// <param name="pathconvertedFolder"> The path of folder for converted file</param>
		/// <param name="fileName"> the uploaded file name</param>
		/// <returns>return True if convertation is successful</returns>
		private bool SaveAsJson(List<Dictionary<object, object>> list, string pathconvertedFolder, string uploadedFileName)
		{
			try
			{
				string json = JsonConvert.SerializeObject(list);
				System.IO.File.WriteAllText(pathconvertedFolder + "\\Json_" + uploadedFileName, json);
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