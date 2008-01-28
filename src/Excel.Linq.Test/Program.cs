﻿
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
