
#region namespaces

using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
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

		#region innerClass

		#region ExpressionRebuilder クラス

		/// <summary>
		/// 式を再構築するクラス
		/// </summary>
		private class ExpressionRebuilder {

			#region fields

			/// <summary>
			/// XlsWorksheet型のParameterExpressionと置換するParameter
			/// </summary>
			private ParameterExpression paramExpr;

			#endregion

			#region constructors

			/// <summary>
			/// Worksheetのパラメータ名を設定するコンストラクタ
			/// </summary>
			/// <param name="name"></param>
			public ExpressionRebuilder(string name) {
				paramExpr = Expression.Parameter(typeof(Worksheet), name);
			}

			#endregion

			#region methods

			/// <summary>
			/// 指定した式を解析して、ExpressionTreeを再構築します。
			/// </summary>
			/// <param name="expr"></param>
			/// <returns></returns>
			public Expression Rebuild(Expression expr) {
				switch(expr.NodeType) {
					case ExpressionType.Lambda:
						return Expression.Lambda<Predicate<Worksheet>>(
							Rebuild(((LambdaExpression)expr).Body),
							((LambdaExpression)expr).Parameters.ToList().ConvertAll<ParameterExpression>(
								p => (ParameterExpression)Rebuild(p)
							)
						);

					case ExpressionType.Equal:
					case ExpressionType.NotEqual:
					case ExpressionType.GreaterThan:
					case ExpressionType.GreaterThanOrEqual:
					case ExpressionType.LessThan:
					case ExpressionType.LessThanOrEqual:
						BinaryExpression bexpr = (BinaryExpression)expr;

						return (Expression)typeof(Expression).InvokeMember(expr.NodeType.ToString(),
							BindingFlags.Public | BindingFlags.Static | BindingFlags.InvokeMethod,
							null, null,
							new object[] {
								Rebuild(bexpr.Left),
								Rebuild(bexpr.Right),
								bexpr.IsLiftedToNull, bexpr.Method
							}
						);

					case ExpressionType.Parameter:
						return ((ParameterExpression)expr).Type == typeof(XlsWorksheet) ? paramExpr : expr;

					case ExpressionType.MemberAccess:
						MemberExpression mexpr = (MemberExpression)expr;
						// XlsWorksheetであればメンバ名をマッピングする。
						MemberInfo member = mexpr.Member.DeclaringType != typeof(XlsWorksheet) ?
							mexpr.Member :
							((Func<MemberInfo, MemberInfo>)((m) => {
								if(m.Name == "Name") return typeof(_Worksheet).GetProperty("Name");

								return typeof(_Worksheet).GetProperty(m.Name);

							}))(mexpr.Member);

						return Expression.Property(
							Rebuild(mexpr.Expression), (PropertyInfo)member
						);
				}
				return expr;
			}

			#endregion

		}

		#endregion

		#endregion

		#region fields

		private Sheets worksheets;

		private Expression expression;

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
		/// 生のWorksheetオブジェクトと列挙結果をフィルタリングする式を設定するコンストラクタ
		/// </summary>
		/// <param name="worksheets">Worksheetsオブジェクト</param>
		/// <param name="expression">式</param>
		private XlsWorksheets(Sheets worksheets, Expression expression)
			: this(worksheets) {
			this.expression = expression;
		}

		/// <summary>
		/// デストラクタ
		/// </summary>
		~XlsWorksheets() {
			Dispose();
		}

		#endregion

		#region methods

		/// <summary>
		/// 指定した式を条件として、オブジェクトの列挙を行います。
		/// </summary>
		/// <param name="expression">式</param>
		/// <returns>コレクション</returns>
		private IEnumerable<XlsWorksheet> ExecuteExpression(Expression expression) {
			Predicate<Worksheet> predicate = ParseExpression(expression);

			foreach(Worksheet worksheet in worksheets) {
				if(predicate(worksheet)) yield return new XlsWorksheet(worksheet);
				else {
					Marshal.ReleaseComObject(worksheet);
				}
			}
		}

		/// <summary>
		/// 指定した式を解析して、適切なデリゲートに変換します。
		/// </summary>
		/// <param name="expression">式</param>
		/// <returns>デリゲート</returns>
		private Predicate<Worksheet> ParseExpression(Expression expression) {
			Expression lexpr = RebuildExpression(expression);

			return (Predicate<Worksheet>)((LambdaExpression)lexpr).Compile();
		}

		/// <summary>
		/// 指定した式を解析して、ExpressionTreeを再構築します。
		/// </summary>
		/// <param name="expression">式</param>
		/// <returns>Expression</returns>
		private Expression RebuildExpression(Expression expression) {
			return new ExpressionRebuilder("s").Rebuild(expression);
		}

		/// <summary>
		/// 式無しでアイテムの列挙のみを行います。
		/// </summary>
		/// <returns>コレクション</returns>
		private IEnumerable<XlsWorksheet> ForEachWithoutExpression() {
			foreach(Worksheet sheet in worksheets) yield return new XlsWorksheet(sheet);
		}

		#endregion

		#region IEnumerable<XlsWorksheet> メンバ

		/// <summary>
		/// <see cref="IEnumerable&lt;XlsWorksheet&gt;.GetEnumerator()"/>
		/// </summary>
		/// <returns></returns>
		public IEnumerator<XlsWorksheet> GetEnumerator() {
			return Provider.Execute<IEnumerator<XlsWorksheet>>(Expression);
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
			get { return typeof(XlsWorksheet); }
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
			return new XlsWorksheets(this.worksheets, expression);
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
