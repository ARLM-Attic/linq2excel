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
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;

#endregion

namespace Excel.Linq {

	#region XlsWorkbook class

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
