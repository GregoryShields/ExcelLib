using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
	public class ExcelWorksheet : IDisposable
	{
		readonly ExcelInterop.Worksheet _interopWorksheet;

		/// <summary>
		/// Constructor that automatically sets the _ExcelWorkbook and _interopWorksheet private fields as valid references.
		/// </summary>
		/// <param name="excelWorkbook">
		/// Either start a new instance of Excel which contains a new Workbook with a single Worksheet, or connect to an existing instance in memory.
		/// </param>
		/// <param name="interopWorksheet">
		/// If <see cref="excelWorkbook"/> is null, the new Worksheet will be automatically named with this value. Otherwise, there is an attempt 
		/// to locate a Worksheet by this name in the first Workbook of the existing Excel instance. If successful, _interopWorksheet is set as a reference to it. 
		/// Otherwise, an exception is thrown.
		/// </param>
		public ExcelWorksheet(ExcelWorkbook excelWorkbook, ExcelInterop.Worksheet interopWorksheet)
		{
			ExcelWorkbook = excelWorkbook;
			_interopWorksheet = interopWorksheet;
		}

		public void Dispose()
		{
			if (_interopWorksheet != null)
			{
				GC.Collect();
				GC.WaitForPendingFinalizers();
				// Marshal.FinalReleaseComObject(_Rng); // Release a Range object if you have one.
				Marshal.FinalReleaseComObject(_interopWorksheet);
			}
			//ExcelWorkbook.Dispose(); We don't want this because we may have multiple worksheets under the workbook object.
		}

		public ExcelWorkbook ExcelWorkbook { get; private set; }

		public ExcelInterop.Worksheet InteropWorksheet => _interopWorksheet;

		public bool DisplayGridlines
		{
			set => InteropWorksheet.Parent.Windows(1).DisplayGridlines = value;
		}

		// I commented this out until I could get around to adding the proper references. I believe I made a note about it in the BKEP SecRev project.
		//public void AddTextBox(float left, float top, float width)
		//{
		//	InteropWorksheet.Shapes.AddTextbox(Orientation: MsoTextOrientation.msoTextOrientationDownward, Left: left, Top: top, Width: width, Height: 100.0f);
		//}

		public int RowCount =>
			InteropWorksheet.Cells
				.Find(
					"*",
					Type.Missing,
					Type.Missing,
					Type.Missing,
					ExcelInterop.XlSearchOrder.xlByRows,
					ExcelInterop.XlSearchDirection.xlPrevious,
					Type.Missing,
					Type.Missing,
					Type.Missing
				).Row;
	}
}

