using Dapper;
using System;
using System.Collections.Generic;
using System.Data;

/*
 * from:  http://stackoverflow.com/questions/28678442/how-can-i-make-dapper-net-throw-when-result-set-has-unmapped-columns/39490419#39490419
 * solution by Richardissimo
 */

namespace SafeDapper
{
    public static class DapperExtensions
    {
        private static readonly HashSet<Type> TypesThatHaveMapper = new HashSet<Type>();
        private static readonly object _lock = new object();

        /// <summary>
        /// Extension to the Dapper methods, SafeQuery will throw an exception if any column in the query 
        /// is not mapped to a property of the type. This prevents silent failure/defaulted property values
        /// in the case that the names of columns/properties are misspelled, or changed in one place but not another.
        /// </summary>
        public static IEnumerable<T> SafeQuery<T>(this IDbConnection cnn, string sql, object param = null,
            IDbTransaction transaction = null, bool buffered = true, int? commandTimeout = default(int?),
            CommandType? commandType = default)
        {
            AssertSafeTypeMap<T>();

            return cnn.Query<T>(sql, param, transaction, buffered, commandTimeout, commandType);
        }


        /// <summary>
        /// Extension to the Dapper methods, SafeQuery will throw an exception if any column in the query 
        /// is not mapped to a property of the type. This prevents silent failure/defaulted property values
        /// in the case that the names of columns/properties are misspelled, or changed in one place but not another.
        /// </summary>
        public static T SafeQuerySingle<T>(this IDbConnection cnn, string sql, object param = null, IDbTransaction transaction = null, int? commandTimeout = default(int?), CommandType? commandType = default(CommandType?))
        {
            AssertSafeTypeMap<T>();

            return cnn.QuerySingle<T>(sql, param, transaction, commandTimeout, commandType);
        }

        /// <summary>
        /// Specifies a safe type mapper (<see cref="ThrowWhenNullTypeMap{T}"/>) for <typeparamref name="T"/>.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static void AssertSafeTypeMap<T>()
        {
            // First do cheap check without  locking
            if (!TypesThatHaveMapper.Contains(typeof(T))){

                // Then lock
                lock (_lock)
                {
                    //And check again
                    if (!TypesThatHaveMapper.Contains(typeof(T)))
                    {
                        SqlMapper.SetTypeMap(typeof(T), new ThrowWhenNullTypeMap<T>());

                        TypesThatHaveMapper.Add(typeof(T));
                    }
                }
            }
        }

    }
}
