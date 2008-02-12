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
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;

#endregion

namespace Excel.Linq {

	#region XlsCell class

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

		#region IDisposable member

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
