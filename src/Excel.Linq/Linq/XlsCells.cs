
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

	#region XlsCellsクラス

	/// <summary>
	/// Excelワークシートのセルのコレクションに対する操作を提供するクラス
	/// </summary>
	public class XlsCells : IQueryable<XlsCell>, IQueryProvider, IDisposable {

		#region fields

		private Range cells;

		private Expression expression;

		#endregion

		#region properties

		/// <summary>
		/// 指定した行、列にあるセルを取得します。
		/// </summary>
		/// <param name="row">行番号</param>
		/// <param name="column">列番号</param>
		/// <returns>セルオブジェクト</returns>
		public XlsCell this[int row, int column] {
			get {
				// Excel COM は1ベースの為、+1しとく
				Range cell = cells[row + 1, column + 1] as Range;

				return new XlsCell(cell);
			}
		}

		#endregion

		#region constructors

		/// <summary>
		/// 生のRangeオブジェクトを設定するコンストラクタ
		/// </summary>
		/// <param name="cells">Rangeオブジェクト</param>
		protected internal XlsCells(Range cells) {
			this.cells = cells;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cells"></param>
		/// <param name="expression"></param>
		private XlsCells(Range cells, Expression expression)
			: this(cells) {
			this.expression = expression;
		}

		/// <summary>
		/// デストラクタ
		/// </summary>
		~XlsCells() {
			Dispose();
		}

		#endregion

		#region methods

		/// <summary>
		/// Expressionなしの結果セットを取得します。
		/// </summary>
		/// <returns>結果セット</returns>
		private IEnumerable<XlsCell> ExecuteNone() {
			foreach(Range cell in cells) yield return new XlsCell(cell);
		}

		/// <summary>
		/// 指定したExpressionを実行します。
		/// </summary>
		/// <param name="expression">Expressionオブジェクト</param>
		/// <returns>結果セット</returns>
		protected virtual IEnumerable<XlsCell> ExecuteExpression(Expression expression) {
			foreach(Range cell in cells) {
			}
			yield break;
		}

		#endregion

		#region IEnumerable<XlsCell> メンバ

		/// <summary>
		/// <see cref="IEnumerable&lt;XlsCell&gt;.GetEnumerator()"/>
		/// </summary>
		/// <returns></returns>
		public IEnumerator<XlsCell> GetEnumerator() {
			return this.Provider.Execute<IEnumerator<XlsCell>>(Expression);
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
		public Type ElementType {
			get { return typeof(XlsCell); }
		}

		/// <summary>
		/// <see cref="IQueryable.Expression"/>
		/// </summary>
		public Expression Expression {
			get { return expression; }
		}

		/// <summary>
		/// <see cref="IQueryable.Provider"/>
		/// </summary>
		public IQueryProvider Provider {
			get { return this; }
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
			return (IQueryable<TElement>)CreateQuery(expression);
		}

		/// <summary>
		/// <see cref="IQueryProvider.CreateQuery(Expression)"/>
		/// </summary>
		/// <param name="expression"></param>
		/// <returns></returns>
		public IQueryable CreateQuery(Expression expression) {
			return new XlsCells(this.cells, expression);
		}

		/// <summary>
		/// <see cref="IQueryProvider&lt;TResult&gt;.Execute(Expression)"/>
		/// </summary>
		/// <typeparam name="TResult"></typeparam>
		/// <param name="expression"></param>
		/// <returns></returns>
		public TResult Execute<TResult>(Expression expression) {
			return (TResult)Execute(expression);
		}

		/// <summary>
		/// <see cref="IQueryProvider.Execute(Expression)"/>
		/// </summary>
		/// <param name="expression"></param>
		/// <returns></returns>
		public object Execute(Expression expression) {
			return (expression != null ?
				ExecuteExpression(expression) : ExecuteNone()
			).GetEnumerator();
		}

		#endregion

		#region IDisposable メンバ

		/// <summary>
		/// リソースを開放します。
		/// </summary>
		/// <param name="disposing">破棄するかどうか</param>
		protected virtual void Dispose(bool disposing) {
			if(!disposing) return;

			if(cells != null) {
				Marshal.ReleaseComObject(cells);
				cells = null;
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
