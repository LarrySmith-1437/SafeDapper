using System;
using System.Reflection;
using Dapper;

namespace SafeDapper
{
    class ThrowWhenNullTypeMap<T>:SqlMapper.ITypeMap
    {
        private readonly SqlMapper.ITypeMap _defaultTypeMap = new DefaultTypeMap(typeof(T));

        public ConstructorInfo FindConstructor(string[] names, Type[] types)
        {
            return _defaultTypeMap.FindConstructor(names, types);
        }

        public ConstructorInfo FindExplicitConstructor()
        {
            return _defaultTypeMap.FindExplicitConstructor();
        }

        public SqlMapper.IMemberMap GetConstructorParameter(ConstructorInfo constructor, string columnName)
        {
            return _defaultTypeMap.GetConstructorParameter(constructor, columnName);
        }

        public SqlMapper.IMemberMap GetMember(string columnName)
        {
            var member = _defaultTypeMap.GetMember(columnName);
            if (member == null)
            {
                throw new DapperObjectMappingException($"Column {columnName} could not be mapped to object.");
            }
            return member;
        }
    }
}
