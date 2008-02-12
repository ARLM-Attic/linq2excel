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
using System.Linq;
using System.Linq.Expressions;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;

using Excel.Interop;

#endregion

namespace Excel.Linq {

	#region XlsCells class

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
		private IEnumerable<XlsCell> ForEachWithoutExpression() {
			foreach(Range cell in cells) yield return new XlsCell(cell);
		}

		/// <summary>
		/// 指定したExpressionを実行します。
		/// </summary>
		/// <param name="expression">Expressionオブジェクト</param>
		/// <returns>結果セット</returns>
		protected virtual IEnumerable<XlsCell> ExecuteExpression(Expression expression) {
			yield break;
		}

		#endregion

		#region IEnumerable<XlsCell> member

		/// <summary>
		/// <see cref="IEnumerable&lt;XlsCell&gt;.GetEnumerator()"/>
		/// </summary>
		/// <returns></returns>
		public IEnumerator<XlsCell> GetEnumerator() {
			return this.Provider.Execute<IEnumerator<XlsCell>>(Expression);
		}

		#endregion

		#region IEnumerable member

		/// <summary>
		/// <see cref="IEnumerable.GetEnumerator()"/>
		/// </summary>
		/// <returns></returns>
		IEnumerator IEnumerable.GetEnumerator() {
			return this.GetEnumerator();
		}

		#endregion

		#region IQueryable member

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

		#region IQueryProvider member

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
				ExecuteExpression(expression) : ForEachWithoutExpression()

			).GetEnumerator();
		}

		#endregion

		#region IDisposable member

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
