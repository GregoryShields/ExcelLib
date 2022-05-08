using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
	/// <summary>
	/// Instantiated only from ExcelApp by one of these two methods:
	/// NewWorkbookWithSingleSheet()
	/// OpenWorkbookReadOnly(string workbookPath)
	/// </summary>
	public class ExcelWorkbook : IDisposable
	{
		//public ExcelWorkbook(ExcelApp excelApp, string[] workSheetNames)
		//{
		//	ExcelApp = excelApp;
		//	// Create a new Workbook with a single Worksheet:
		//	ExcelInteropWorkbook = excelApp.ExcelInteropApp.Workbooks.Add(ExcelInterop.XlWBATemplate.xlWBATWorksheet);

		//	var theFirstInteropWorksheet = ExcelInteropWorkbook.Worksheets[1]; // This worksheet was automatically created when the workbook was created above.
		//	theFirstInteropWorksheet.Name = workSheetNames[0];

		//	ExcelWorksheets = new Dictionary<string, ExcelWorksheet>
		//	{
		//		{ workSheetNames[0], new ExcelWorksheet(this, theFirstInteropWorksheet) }
		//	};

		//	for (int i = 1; i < workSheetNames.Length; i++) // Skip the first name since it was already dealt with above.
		//	{
		//		// Add another worksheet to the physical workbook:
		//		ExcelInteropWorkbook.Worksheets.Add(
		//		ExcelWorksheets.Add(workSheetNames[i], new ExcelWorksheet(this, theFirstInteropWorksheet));
				
		//	}
		//}

		/// <summary>
		/// Constructor that receives a name that it uses to name the Worksheet in a new Workbook that it creates.
		/// </summary>
		/// <param name="excelApp"></param>
		/// <param name="workSheetName"></param>
		public ExcelWorkbook(ExcelApp excelApp, string workSheetName)
		{
			ExcelApp = excelApp;
			// Create a new Workbook with a single Worksheet:
			ExcelInteropWorkbook = excelApp.ExcelInteropApp.Workbooks.Add(ExcelInterop.XlWBATemplate.xlWBATWorksheet);

			var theOnlyInteropWorksheet = ExcelInteropWorkbook.Worksheets[1];
			theOnlyInteropWorksheet.Name = workSheetName;

			ExcelWorksheets = new Dictionary<string, ExcelWorksheet>
			{
				{ workSheetName, new ExcelWorksheet(this, theOnlyInteropWorksheet) }
			};
		}

		/// <summary>
		/// Constructor that receives a reference to a Workbook opened here:
		/// ExcelApp.OpenWorkbookReadOnly(string workbookPath, string workbookName)
		/// </summary>
		/// <param name="excelApp"></param>
		/// <param name="interopWorkbook"></param>
		public ExcelWorkbook(ExcelApp excelApp, ExcelInterop.Workbook interopWorkbook)
		{
			ExcelApp = excelApp;
			ExcelInteropWorkbook = interopWorkbook;

			InitializeExcelWorksheets();
		}

		void InitializeExcelWorksheets()
		{
			ExcelWorksheets = new Dictionary<string, ExcelWorksheet>();

			foreach (ExcelInterop.Worksheet interopWorksheet in ExcelInteropWorkbook.Worksheets)
			{
				var ews = new ExcelWorksheet(this, interopWorksheet);
				ExcelWorksheets.Add(interopWorksheet.Name, ews);
			}
		}

		public ExcelWorksheet AddWorksheet(string worksheetName)
		{
			ExcelInteropWorkbook.Sheets.Add(After: ExcelInteropWorkbook.Sheets[ExcelInteropWorkbook.Sheets.Count]);
			ExcelInteropWorkbook.Sheets[ExcelInteropWorkbook.Sheets.Count].Name = worksheetName;
			var ews = new ExcelWorksheet(this, ExcelInteropWorkbook.Sheets[ExcelInteropWorkbook.Sheets.Count]);
			ExcelWorksheets.Add(worksheetName, ews);
			return ews;
		}

		public ExcelApp ExcelApp { get; private set; }
		public ExcelInterop.Workbook ExcelInteropWorkbook { get; private set; }
		public Dictionary<string, ExcelWorksheet> ExcelWorksheets { get; private set; }

		public void SaveAs(string workbookPathAndName)
		{
			ExcelApp.ExcelInteropApp.DisplayAlerts = false;
			ExcelInteropWorkbook.SaveAs(Filename: workbookPathAndName);
			//ExcelInteropWorkbook.SaveAs(workbookPathAndName, ExcelInterop.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, 
			//	Missing.Value, false, false, ExcelInterop.XlSaveAsAccessMode.xlNoChange, 
			//	ExcelInterop.XlSaveConflictResolution.xlUserResolution, true, 
			//	Missing.Value, Missing.Value, Missing.Value);
			ExcelApp.ExcelInteropApp.DisplayAlerts = true;
		}

		#region IDisposable member

		public void Dispose()
		{
			if (ExcelInteropWorkbook != null)
			{
				// First Dispose all the Worksheet objects.
				foreach (var excelWorksheet in ExcelWorksheets)
					excelWorksheet.Value.Dispose();

				GC.Collect();
				GC.WaitForPendingFinalizers();
				ExcelInteropWorkbook.Close(false, false, Type.Missing); // CHANGE THIS LATER!?
				Marshal.FinalReleaseComObject(ExcelInteropWorkbook);
			}
			//ExcelApp.Dispose(); We don't want this because we may have multiple workbooks under the app object.
		}

		#endregion IDisposable member

		#region I've disabled these because they destroy encapsulation of the interop objects.
		//public ExcelInterop.Worksheet FirstInteropWorksheet
		//{
		//	get { return ExcelInteropWorkbook.Worksheets[1]; }
		//}

		//public ExcelInterop.Sheets InteropWorksheets
		//{
		//	get { return ExcelInteropWorkbook.Worksheets; }
		//}
		#endregion
	}
}

