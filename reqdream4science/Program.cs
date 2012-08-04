/*
 * Created by SharpDevelop.
 * User: Richard
 * Date: 8/3/2012
 * Time: 4:24 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Linq;
using ExcelLibrary.SpreadSheet;

namespace reqdream4science
{
	class Program
	{
		public static void Main(string[] args)
		{
			ProcessFiles(args[0], args[1], args[2], args[3]);
		}
		
		static void ProcessFiles(string imageName, string results1path, string results2path, string outputXlsPath)
		{
			Workbook workbook = null;
			Worksheet sheet = null;
			
			// parse results 1: reads lines into an array and selects the 2nd element (element index 1)
			// which contains the actual data
			string[] results1 = File.ReadAllLines(results1path)[1].Split('\t');
			string[][] results2 = File.ReadAllLines(results2path).Skip(1).Select(line => line.Split('\t')).ToArray();
						
			// create output xl, if it's not there
			// or load and grab the next sheet:
			if(!File.Exists(outputXlsPath))
			{
				workbook = new Workbook();
				sheet = new Worksheet("Results");
				workbook.Worksheets.Add(sheet);
			}
			else
			{
				workbook = Workbook.Load(outputXlsPath);	
				sheet = workbook.Worksheets[0];
			}
			
			// find the next empty columns to write in:
			int column = 0;			
			bool found = string.IsNullOrWhiteSpace(
					(sheet.Cells[9, column].Value == null ? string.Empty :
					sheet.Cells[9, column].Value.ToString()));
			
			while(!found)
			{
				// if column and column + 1 are blank, then we found en empty spot:
				found = string.IsNullOrWhiteSpace(
					(sheet.Cells[9, column].Value == null ? string.Empty :
					sheet.Cells[9, column].Value.ToString()))
					&& string.IsNullOrWhiteSpace(
					(sheet.Cells[9, column + 1].Value == null ? string.Empty :
					sheet.Cells[9, column + 1].Value.ToString()));
					
				column++;
			}
						
			// start writing in the data:			
			// the image name:
			sheet.Cells[0, column] = new Cell(imageName);
			
			// the translated data from results1.xls
			sheet.Cells[2, column] = new Cell(results1[0]);
			sheet.Cells[3, column] = new Cell(results1[1]);
			sheet.Cells[4, column] = new Cell(Double.Parse(results1[2]));
			sheet.Cells[5, column] = new Cell(Int32.Parse(results1[3]));
			sheet.Cells[6, column] = new Cell(Int32.Parse(results1[4]));
			
			// the 2 columns from results2.xls, sans headers:
			for(int r2index = 0; r2index < results2.Length; r2index++)
			{
				sheet.Cells[r2index + 8, column] = new Cell(Int32.Parse(results2[r2index][0]));
				sheet.Cells[r2index + 8, column + 1] = new Cell(Int32.Parse(results2[r2index][1]));
			}
			
			workbook.Save(outputXlsPath);
		}	
	}
}