
#region namespaces

using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq.Expressions;

#endregion

namespace Excel.Linq {

	#region XlsExtensionクラス

	/// <summary>
	/// Excel.Linq の各クラスに対して拡張メソッドを提供するクラス
	/// </summary>
	public static class XlsExtension {

		#region methods

		/// <summary>
		/// 
		/// </summary>
		/// <param name="collection"></param>
		/// <param name="expression"></param>
		/// <returns></returns>
		public static IQueryable<XlsWorksheet> Where(
			this IQueryable<XlsWorksheet> collection, Expression<Predicate<XlsWorksheet>> expression) {
			return collection.Provider.CreateQuery<XlsWorksheet>(expression);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="collection"></param>
		/// <param name="expression"></param>
		/// <returns></returns>
		public static IQueryable<XlsCell> Where(
			this IQueryable<XlsCell> collection, Expression<Predicate<XlsCell>> expression) {
			return collection.Provider.CreateQuery<XlsCell>(expression);
		}

		#endregion

	}

	#endregion

}
