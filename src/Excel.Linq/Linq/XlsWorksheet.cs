
#region namespaces

using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;
using System.Collections.Generic;

#endregion

namespace Excel.Linq {

	#region XlsWorksheetクラス

	/// <summary>
	/// Excelワークシートに対する操作を提供するクラス
	/// </summary>
	public class XlsWorksheet : IDisposable {

		#region fields

		private Worksheet worksheet;

		#endregion

		#region properties

		/// <summary>
		/// ワークシート名を取得、設定します。
		/// </summary>
		public string Name {
			get { return worksheet.Name; }
			set { worksheet.Name = value; }
		}

		#endregion

		#region constructors

		/// <summary>
		/// 生のWorksheetオブジェクトを設定するコンストラクタ
		/// </summary>
		/// <param name="worksheet">Worksheetオブジェクト</param>
		protected internal XlsWorksheet(Worksheet worksheet) {
			this.worksheet = worksheet;
		}

		#endregion

		#region methods

		/// <summary>
		/// 指定したセル範囲のセルのコレクションを取得します。
		/// </summary>
		/// <param name="startCell">開始セル</param>
		/// <param name="endCell">終了セル</param>
		/// <returns>セルのコレクション</returns>
		public XlsCells Range(string startCell, string endCell) {
			Range range = worksheet.get_Range(startCell, endCell) as Range;

			return new XlsCells(range);
		}

		/// <summary>
		/// 指定したセル範囲のセルのコレクションを取得します。
		/// </summary>
		/// <param name="startRow"></param>
		/// <param name="endRow"></param>
		/// <param name="startCol"></param>
		/// <param name="endCol"></param>
		/// <returns></returns>
		public IEnumerable<XlsCell> Range(int startRow, int endRow, int startCol, int endCol) {
			for(int row = startRow; row < endRow; row++) {
				for(int col = startCol; col < endCol; col++) {
					yield return new XlsCell(worksheet.Cells[row, col] as Range);
				}
			}
		}

		#endregion

		#region IDisposable メンバ

		/// <summary>
		/// リソースを開放します。
		/// </summary>
		/// <param name="disposing">破棄するかどうか</param>
		protected virtual void Dispose(bool disposing) {
			if(!disposing) return;

			if(worksheet != null) {
				Marshal.ReleaseComObject(worksheet);
				worksheet = null;
			}
		}

		/// <summary>
		/// <see cref="IDisposable.Dispose()"/>
		/// </summary>
		public void Dispose() {
			lock(this) {
				Dispose(true);

				GC.SuppressFinalize(this);
			}
		}

		#endregion

	}

	#endregion

}
