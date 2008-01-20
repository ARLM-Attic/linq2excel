
#region namespaces

using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;

#endregion

namespace Excel.Linq {

	#region XlsCellクラス

	/// <summary>
	/// Excelワークシートのセルに対する操作を提供するクラス
	/// </summary>
	[DebuggerStepThrough]
	public class XlsCell : IDisposable {

		#region fields

		private Range cell;

		#endregion

		#region properties

		/// <summary>
		/// 行番号を取得します。
		/// </summary>
		public int Row {
			get { return cell.Row - 1; }
		}

		/// <summary>
		/// 列番号を取得します。
		/// </summary>
		public int Column {
			get { return cell.Column - 1; }
		}

		/// <summary>
		/// テキストを取得します。
		/// </summary>
		public string Text {
			get { return cell.Text.ToString(); }
		}

		#endregion

		#region constructors

		/// <summary>
		/// 生のRangeオブジェクトを設定するコンストラクタ
		/// </summary>
		/// <param name="cell">Rangeオブジェクト</param>
		protected internal XlsCell(Range cell) {
			this.cell = cell;
		}

		#endregion

		#region methods

		/// <summary>
		/// <see cref="object.ToString()"/>
		/// </summary>
		/// <returns></returns>
		public override string ToString() {
			return string.Format("Row = {0}, Column = {1}, Text = {2}", Row, Column, Text);
		}

		#endregion

		#region IDisposable メンバ

		/// <summary>
		/// リソースを開放します。
		/// </summary>
		/// <param name="disposing">破棄するかどうか</param>
		protected virtual void Dispose(bool disposing) {
			if(!disposing) return;

			if(cell != null) {
				Marshal.ReleaseComObject(cell);
				cell = null;
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
