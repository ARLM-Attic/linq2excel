#region license
/*
Copyright (c) 2007, Cozy Yamaguchi
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

* Neither the name of cozy nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY,
OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
#endregion

#region namespaces

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;

#endregion

namespace Excel.Linq {

	#region XlsWorksheet class

	/// <summary>
	/// Excelワークシートに対する操作を提供するクラス
	/// </summary>
	public class XlsWorksheet : IDisposable {

		#region fields

		private Worksheet worksheet;
		/// <summary>
		/// シートの全セル
		/// </summary>
		private XlsCells cells;

		#endregion

		#region properties

		/// <summary>
		/// ワークシート名を取得、設定します。
		/// </summary>
		public string Name {
			get { return worksheet.Name; }
			set { worksheet.Name = value; }
		}

		/// <summary>
		/// セルのコレクションを取得します。
		/// </summary>
		public XlsCells Cells {
			get { return cells; }
		}

		#endregion

		#region constructors

		/// <summary>
		/// 生のWorksheetオブジェクトを設定するコンストラクタ
		/// </summary>
		/// <param name="worksheet">Worksheetオブジェクト</param>
		protected internal XlsWorksheet(Worksheet worksheet) {
			this.worksheet = worksheet;
			this.cells = new XlsCells(worksheet.Cells);
		}

		/// <summary>
		/// デストラクタ
		/// </summary>
		~XlsWorksheet() {
			Dispose();
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

		#region IDisposable member

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
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		#endregion

	}

	#endregion

}
