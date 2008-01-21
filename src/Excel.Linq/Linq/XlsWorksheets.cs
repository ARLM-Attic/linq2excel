
#region namespaces

using System;
using System.Linq;
using System.Linq.Expressions;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;

#endregion

namespace Excel.Linq {

	#region XlsWorksheetsクラス

	/// <summary>
	/// Excelワークシートのコレクションに対する操作を提供するクラス
	/// </summary>
	public class XlsWorksheets : IQueryable<XlsWorksheet>, IQueryProvider, IDisposable {

		#region fields

		private Sheets worksheets;

		#endregion

		#region properties

		/// <summary>
		/// 指定したインデックスのワークシートを取得します。
		/// </summary>
		/// <param name="index">インデックス</param>
		/// <returns>ワークシート</returns>
		/// <exception cref="IndexOutOfRangeException">インデックスの範囲が領域外の時</exception>
		public XlsWorksheet this[int index] {
			get {
				Worksheet worksheet = worksheets[index + 1] as Worksheet;
				// nullの場合は例外が飛ばずにnull参照が返ってくる。
				if(worksheet == null) {
					throw new IndexOutOfRangeException();
				}
				return new XlsWorksheet(worksheet);
			}
		}

		#endregion

		#region constructors

		/// <summary>
		/// 生のWorksheetオブジェクトを設定するコンストラクタ
		/// </summary>
		/// <param name="worksheets">Worksheetオブジェクト</param>
		protected internal XlsWorksheets(Sheets worksheets) {
			this.worksheets = worksheets;
		}

		/// <summary>
		/// デストラクタ
		/// </summary>
		~XlsWorksheets() {
			Dispose();
		}

		#endregion

		#region IEnumerable<XlsWorksheet> メンバ

		/// <summary>
		/// <see cref="IEnumerable&lt;XlsWorksheet&gt;.GetEnumerator()"/>
		/// </summary>
		/// <returns></returns>
		public IEnumerator<XlsWorksheet> GetEnumerator() {
			foreach(Worksheet sheet in worksheets) yield return new XlsWorksheet(sheet);
		}

		#endregion

		#region IEnumerable メンバ

		/// <summary>
		/// <see cref="IEnumerable.GetEnumerator()"/>
		/// </summary>
		/// <returns></returns>
		IEnumerator IEnumerable.GetEnumerator() {
			return this.GetEnumerator();
		}

		#endregion

		#region IQueryable メンバ

		/// <summary>
		/// <see cref="IQueryable.ElementType"/>
		/// </summary>
		Type IQueryable.ElementType {
			get { return typeof(XlsWorksheet); }
		}

		/// <summary>
		/// <see cref="IQueryable.Expression"/>
		/// </summary>
		Expression IQueryable.Expression {
			get { return null; }
		}

		/// <summary>
		/// <see cref="IQueryable.Provider"/>
		/// </summary>
		IQueryProvider IQueryable.Provider {
			get { return null; }
		}

		#endregion

		#region IQueryProvider メンバ

		/// <summary>
		/// <see cref="IQueryProvider.CreateQuery&lt;TElement&gt;.CreateQuery(Expression)"/>
		/// </summary>
		/// <typeparam name="TElement"></typeparam>
		/// <param name="expression"></param>
		/// <returns></returns>
		public IQueryable<TElement> CreateQuery<TElement>(Expression expression) {
			return null;
		}

		/// <summary>
		/// <see cref="IQueryProvider.CreateQuery(Expression)"/>
		/// </summary>
		/// <param name="expression"></param>
		/// <returns></returns>
		public IQueryable CreateQuery(Expression expression) {
			return null;
		}

		/// <summary>
		/// <see cref="IQueryProvider&lt;TResult&gt;.Execute(Expression)"/>
		/// </summary>
		/// <typeparam name="TResult"></typeparam>
		/// <param name="expression"></param>
		/// <returns></returns>
		public TResult Execute<TResult>(Expression expression) {
			throw new NotImplementedException();
		}

		/// <summary>
		/// <see cref="IQueryProvider.Execute(Expression)"/>
		/// </summary>
		/// <param name="expression"></param>
		/// <returns></returns>
		public object Execute(Expression expression) {
			throw new NotImplementedException();
		}

		#endregion

		#region IDisposable メンバ

		/// <summary>
		/// リソースを開放します。
		/// </summary>
		/// <param name="disposing">破棄するかどうか</param>
		protected virtual void Dispose(bool disposing) {
			if(!disposing) return;

			if(worksheets != null) {
				Marshal.ReleaseComObject(worksheets);
				worksheets = null;
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
