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
using System.Collections.Generic;
using System.Diagnostics;

using Coma2n.Commons;

#endregion

namespace Excel.Linq {

	#region XlsExtension class

	/// <summary>
	/// Excel.Linq の各クラスに対して拡張メソッドを提供するクラス
	/// </summary>
	public static class XlsExtension {

		#region methods

		/// <summary>
		/// 指定した述語に基づいて、シーケンスをフィルタリングします。
		/// </summary>
		/// <param name="collection">シーケンス</param>
		/// <param name="expression">Expression</param>
		/// <returns>フィルタリングされたシーケンス</returns>
		/// <exception cref="ArgumentNullException">引数がnullの時</exception>
		public static IQueryable<XlsWorksheet> Where(
			this IQueryable<XlsWorksheet> collection, Expression<Predicate<XlsWorksheet>> expression) {
			#region ArgumentValidation
			ArgumentValidation.CheckForNullReference(collection, "collection");
			ArgumentValidation.CheckForNullReference(expression, "expression");
			#endregion

			return collection.Provider.CreateQuery<XlsWorksheet>(expression);
		}

		/// <summary>
		/// 指定した述語に基づいて、シーケンスをフィルタリングします。
		/// </summary>
		/// <param name="collection">シーケンス</param>
		/// <param name="expression">Expression</param>
		/// <returns>フィルタリングされたシーケンス</returns>
		/// <exception cref="ArgumentNullException">引数がnullの時</exception>
		public static IQueryable<XlsCell> Where(
			this IQueryable<XlsCell> collection, Expression<Predicate<XlsCell>> expression) {
			#region ArgumentValidation
			ArgumentValidation.CheckForNullReference(collection, "collection");
			ArgumentValidation.CheckForNullReference(expression, "expression");
			#endregion

			return collection.Provider.CreateQuery<XlsCell>(expression);
		}

		#endregion

	}

	#endregion

}
