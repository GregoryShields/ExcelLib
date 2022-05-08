using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
	/// <summary>
	/// Represents an Excel Interop Application object.
	/// Use NewWorkbookWithSingleSheet() if you want to create a new workbook.
	/// Use OpenWorkbookReadOnly(string workbookPath, string workbookName) to open an existing workbook.
	/// Then you'll have a reference to an ExcelWorkbook object that wraps the Workbook object and it can be disposed of automatically.
	/// </summary>
	public class ExcelApp : IDisposable
	{
		/// <summary>
		/// Either start and make visible a new instance of Excel which contains a new Workbook, or connect to an existing instance in memory. 
		/// Usually you would set <see cref="startNewExcelInstance"/> to true so that you are not dealing with Workbook objects that are not wrapped in an ExcelWorkbook object. Therefore, it defaults to true.
		/// </summary>
		/// <param name="startNewExcelInstance">
		/// </param>
		public ExcelApp(bool startNewExcelInstance = true)
		{
			ExcelWorkbooks = new Dictionary<string, ExcelWorkbook>();
			if (startNewExcelInstance)
				ExcelInteropApp = new ExcelInterop.Application();
			else
				ConnectToExcel();
		}

		public ExcelInterop.Application ExcelInteropApp { get; private set; }
 
		/// <summary>
		/// Unsaved Workbooks do not have distinct names in the interop API.
		/// Therefore, this property allows us to have a unique name for Workbooks so we can identify them, even if they're not saved to disk.
		/// </summary>
		public Dictionary<string, ExcelWorkbook> ExcelWorkbooks { get; private set; }

		#region These three methods each add a Workbook to the private ExcelInterop.Application object (ExcelInteropApp) and return an ExcelWorkbook object. This provides some consistency to the API.

		public ExcelWorkbook NewWorkbookWithSingleSheet(string workbookAndWorksheetName)
		{
			var ewb = new ExcelWorkbook(this, workbookAndWorksheetName);
			ExcelWorkbooks.Add(workbookAndWorksheetName, ewb); // Name the Workbook the same as its single Worksheet.
			return ewb;
		}

		public ExcelWorkbook NewWorkbookWithSingleSheet(string workbookName, string worksheetName)
		{
			var ewb = new ExcelWorkbook(this, worksheetName);
			ExcelWorkbooks.Add(workbookName, ewb);
			return ewb;
		}

		//public ExcelWorkbook NewWorkbookWithMultipleSheets(string workbookName, string[] worksheetNames)
		//{
			
		//}

		public ExcelWorkbook OpenWorkbookReadOnly(string workbookPath, string workbookName)
		{
			// If no valid file extension is provided, assume "xlsx" is desired.
			if (
				!(
					workbookName.EndsWith(".xlsx") ||
					workbookName.EndsWith(".xlsm") ||
					workbookName.EndsWith(".xls")
				)
			)
				workbookName += ".xlsx";

			ExcelInterop.Workbook interopWorkbook =
				ExcelInteropApp.Workbooks.Open(
					Filename: workbookPath + workbookName,
					UpdateLinks: 0,
					ReadOnly: true,
					Notify: false
				);

			var ewb = new ExcelWorkbook(this, interopWorkbook);
			ExcelWorkbooks.Add(workbookName, ewb);
			return ewb;
		}

		#endregion

		public ExcelInterop.Workbook FirstWorkbook => ExcelInteropApp.Workbooks[1];

		public void MakeVisible()
		{
			ExcelInteropApp.Visible = true;
			ExcelInteropApp.DisplayFullScreen = false;
			ExcelInteropApp.WindowState = ExcelInterop.XlWindowState.xlNormal;
		}

		void ConnectToExcel()
		{
			try
			{
				if (ExcelInteropApp == null)
					ExcelInteropApp = (ExcelInterop.Application)Marshal2.GetActiveObject("Excel.Application");
			}
			catch(COMException)
			{
				throw new Exception("Could not retrieve an active instance of Excel.");
			}
		}

		#region IDisposable member

		// CAN'T DO THIS BECAUSE OF COM ERROR:
		//~ExcelApp()
		//{
		//    // We're being finalized (i.e. destroyed), so call Dispose.
		//    Dispose();
		//}
		public void Dispose()
		{
			if (ExcelInteropApp == null) return;

			// First Dispose all the Workbook objects, which in turn dispose all of their Worksheet objects.
			foreach (var excelWorkbook in ExcelWorkbooks)
				excelWorkbook.Value.Dispose();

			GC.Collect();
			GC.WaitForPendingFinalizers();
			ExcelInteropApp.Quit();
			Marshal.FinalReleaseComObject(ExcelInteropApp);
		}

		#endregion IDisposable member
	}
}
