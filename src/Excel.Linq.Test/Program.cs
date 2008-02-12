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

using System;

namespace Excel.Linq {
	class Program {
		static void Main(string[] args) {
			// Worksheetの検索
			Action method1 = () => {
				using(XlsWorkbook book = new XlsWorkbook("TestData\\100.xls")) {

					var sheets = from s in book.Worksheets
								 where s.Name == "100"
								 select s;

					foreach(var sheet in sheets) Console.WriteLine(sheet.Name);
				}
			};
			Action method2 = () => {
				using(XlsWorkbook book = new XlsWorkbook("TestData\\100.xls")) {
					using(XlsWorksheet sheet = book.Worksheets[0]) {
						var cells = from c in sheet.Cells
									where c.Text.Contains("hoge")
									select c;

						foreach(var cell in cells) Console.WriteLine(cell.Text);
					}
				}
			};
			//method1();
			method2();

			Console.Read();
		}
	}
}
