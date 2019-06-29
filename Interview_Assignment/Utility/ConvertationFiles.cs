namespace Interview_Assignment.Utility
{
	using ExcelDataReader;
	using Newtonsoft.Json;
	using System;
	using System.Collections.Generic;
	using System.Data;
	using System.IO;

	/// <summary>
	/// Convert files to other files
	/// </summary>
	public class ConvertationFiles
	{
		/// <summary>
		/// Convert CSV file to JSON file
		/// </summary>
		/// <param name="pathFile"> The path of uploaded file in local system</param>
		/// <param name="pathconvertedFolder"> The path of folder for converted file</param>
		/// <param name="fileName"> The uploaded file name</param>
		/// <returns>Return True if convertation is successful</returns>
		public bool ConvertCSVToJSON(string pathFile, string pathconvertedFolder, string uploadedFileName)
		{
			using (StreamReader streamReader = new StreamReader(pathFile))
			{
				// list stores every data from uploaded file
				List<Dictionary<object, object>> list = new List<Dictionary<object, object>>();
				// The header of CSV file 
				string[] headerArray = streamReader.ReadLine().Split(',');
				if (headerArray.Length <=0) {
					return false;
				}
				// Each line of uploaded file will be stored as dictionary and store in list
				while (!streamReader.EndOfStream)
				{
					string[] line = streamReader.ReadLine().Split(',');
					Dictionary<object, object> dict = new Dictionary<object, object>();
					for (int i = 0; i < headerArray.Length; i++)
					{
						// If header item is empty, ignore this column
						if (headerArray[i].Trim().Equals(""))
						{
							continue;
						}
						// fill remain items in row by null if this line length is less than header length
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
				string jsonFileName = "FromCSV_" + uploadedFileName.Substring(0, uploadedFileName.LastIndexOf('.')) + ".json";
				return this.SaveAsJson(list, pathconvertedFolder, jsonFileName);
			}
		}

		/// <summary>
		/// Convert EXCEL file ( including ".xlsx", and ".xls" ) to JSON
		/// </summary>
		/// <param name="pathFile"> The path of uploaded file in local system</param>
		/// <param name="pathconvertedFolder"> The path of folder for converted file</param>
		/// <param name="fileName"> The uploaded file name</param>
		/// <returns>Return True if convertation is successful</returns>
		public bool ConvertXLSXToJSON(string pathFile, string pathconvertedFolder, string uploadedFileName)
		{
			using (FileStream stream = System.IO.File.Open(pathFile, FileMode.Open, FileAccess.Read))
			{
				// Using excel package to read data in uploaded file and store it in dataset
				// Each row of dataset will be stored as dictionary and store in list
				using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
				{
					DataSet dataset = reader.AsDataSet();
					
					// Check Tables are created or not and at least 1 tables in table set
					if (dataset.Tables == null || dataset.Tables.Count <= 0)
					{
						return false;
					}

					// Get first sheet in EXCEL
					DataTable table = dataset.Tables[0];
					DataColumnCollection columns = table.Columns;
					if (table.Rows.Count <= 0 || table.Rows[0].ItemArray.Length <=0) {
						return false;
					}
					object[] headerArray = table.Rows[0].ItemArray;
					
					List<Dictionary<object, object>> list = new List<Dictionary<object, object>>();
					for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
					{
						object[] line = table.Rows[rowIndex].ItemArray;
						Dictionary<object, object> dict = new Dictionary<object, object>();
						for (int i = 0; i < headerArray.Length; i++)
						{
							// If header item is empty, ignore this column
							if (headerArray[i].ToString().Trim().Equals("")) {
								continue;
							}
							// fill remain items in row by null if this line length is less than header length
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
					string jsonFileName = "FromExcel_" + uploadedFileName.Substring(0,uploadedFileName.LastIndexOf('.'))+ ".json";
					return this.SaveAsJson(list, pathconvertedFolder, jsonFileName);
				}

			}

		}

		/// <summary>
		/// Save Excel or CSV file as JSOn in specific folder
		/// </summary>
		/// <param name="pathFile"> the path of uploaded file in local system</param>
		/// <param name="pathconvertedFolder"> The path of folder for converted file</param>
		/// <param name="jsonFileName"> the Json file name</param>
		/// <returns>return True if convertation is successful</returns>
		private bool SaveAsJson(List<Dictionary<object, object>> list, string pathconvertedFolder, string jsonFileName)
		{
			try
			{
				string json = JsonConvert.SerializeObject(list);
				File.WriteAllText(pathconvertedFolder + "\\Json_" + jsonFileName, json);
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