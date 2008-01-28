
#region namespaces

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;

#endregion

namespace Excel.Linq {

	#region XlsWorkbookクラス

	/// <summary>
	/// Excelワークブックに対する操作を提供するクラス
	/// </summary>
	public class XlsWorkbook : IDisposable {

		#region const

		/// <summary>
		/// 省略可能引数に渡す値
		/// </summary>
		private static readonly object None = Type.Missing;

		#endregion

		#region fields

		private Application xlsApp;
		private Workbook workbook;

		private XlsWorksheets worksheets;

		#endregion

		#region properties

		/// <summary>
		/// ワークシートのコレクションを取得します。
		/// </summary>
		public XlsWorksheets Worksheets {
			get { return worksheets; }
		}

		#endregion

		#region constructors

		/// <summary>
		/// 指定したExcelファイルを開くコンストラクタ
		/// </summary>
		/// <param name="fileName">ファイル名</param>
		public XlsWorkbook(string fileName) {
			xlsApp = new ApplicationClass();
			workbook = xlsApp.Workbooks.Open(Path.GetFullPath(fileName),
				None, true, None, None, None, None, None, None, None, None, None, None, None, None
			);
			worksheets = new XlsWorksheets(workbook.Sheets);
		}

		/// <summary>
		/// デストラクタ
		/// </summary>
		~XlsWorkbook() {
		}

		#endregion

		#region IDisposable member

		/// <summary>
		/// リソースを開放します。
		/// </summary>
		/// <param name="disposing">破棄するかどうか</param>
		protected virtual void Dispose(bool disposing) {
			if(!disposing) return;

			Worksheets.Dispose();

			if(workbook != null) {
				workbook.Close(false, None, None);
				Marshal.ReleaseComObject(workbook);
				workbook = null;
			}
			if(xlsApp != null) {
				xlsApp.Quit();
				Marshal.ReleaseComObject(xlsApp);
				xlsApp = null;
			}
		}

		/// <summary>
		/// <see cref="IDisposable.Dispose()"/>
		/// </summary>
		public void Dispose() {
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		#endregion

	}

	#endregion

}
