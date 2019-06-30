namespace Convertation_Test
{
	using System.Collections.Generic;
	using System.IO;
	using Interview_Assignment.Utility;
	using Microsoft.VisualStudio.TestTools.UnitTesting;
	
	/// <summary>
	/// Test cases for convertation from CSV and EXCEL to JSON file
	/// </summary>
	[TestClass]
	public class Convertation_Function_Test
	{
		/// <summary>
		/// Test function for converting CSV to JSON file
		/// </summary>
		[TestMethod]
		public void ConvertCSVToJSONTest()
		{
			// Expect Value
			string expectedJson = "[{\"header1\":\"1\",\"header2\":\"2\",\"header3\":\"3\"}]";
			ConvertationFiles convertationFiles = new ConvertationFiles();
			bool result = convertationFiles.ConvertCSVToJSON(@"../../File_Test/input_CSV_test.csv", @"../../File_Test", "output_CSV_test.csv");
			Assert.AreEqual<string>(expectedJson, this.ReadJson(@"../../File_Test/Json_FromCSV_output_CSV_test.json"));
		}

		/// <summary>
		/// Test function for converting EXCEL to JSON file
		/// </summary>
		[TestMethod]
		public void ConvertXLSXToJSON()
		{
			// Expect Value
			string expectedJson = "[{\"header1\":1.0,\"header2\":2.0,\"header3\":3.0}]";
			ConvertationFiles convertationFiles = new ConvertationFiles();
			bool result = convertationFiles.ConvertXLSXToJSON(@"../../File_Test/input_Excel_test.xlsx", @"../../File_Test", "output_Excel_test.csv");
			Assert.AreEqual<string>(expectedJson, this.ReadJson(@"../../File_Test/Json_FromExcel_output_Excel_test.json"));
		}

		/// <summary>
		/// Test function for saving from list to JSON
		/// </summary>
		[TestMethod]
		public void SaveAsJsonTest()
		{
			// Expect Value
			string expectedJson = "[{\"header1\":\"1\",\"header2\":\"2\",\"header3\":\"3\"}]";
			ConvertationFiles convertationFiles = new ConvertationFiles();
			List<Dictionary<object, object>> list = new List<Dictionary<object, object>>();
			Dictionary<object, object> dict = new Dictionary<object, object>();
			dict.Add("header1","1");
			dict.Add("header2", "2");
			dict.Add("header3", "3");
			list.Add(dict);
			bool result = convertationFiles.SaveAsJson(list, @"../../File_Test", "output_SaveAsJson_test.json");
			Assert.AreEqual<string>(expectedJson, this.ReadJson(@"../../File_Test/Json_output_SaveAsJson_test.json"));
		}

		/// <summary>
		/// Read Json from file path
		/// </summary>
		/// <param name="pathFile"> The path of file</param>
		/// <returns>Get Json String</returns>
		public string ReadJson(string pathFile)
		{
			using (StreamReader r = new StreamReader(pathFile))
			{
				string json = r.ReadToEnd();
				return json;
			}
		}
	}
}
