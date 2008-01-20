
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
	/// ...のクラス
	/// </summary>
	public static class XlsExtension {

		#region methods

		/// <summary>
		/// 
		/// </summary>
		/// <param name="collection"></param>
		/// <param name="predicate"></param>
		/// <returns></returns>
		public static IEnumerable<XlsWorksheet> Where(this IEnumerable<XlsWorksheet> collection, Predicate<XlsWorksheet> predicate) {
			foreach(var o in collection) {
				if(predicate(o)) yield return o;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="collection"></param>
		/// <param name="expression"></param>
		/// <returns></returns>
		public static IQueryable<XlsCell> Where(this IQueryable<XlsCell> collection, Expression<Predicate<XlsCell>> expression) {
			return collection.Provider.CreateQuery<XlsCell>(expression);
		}

		#endregion

	}

	#endregion

}
