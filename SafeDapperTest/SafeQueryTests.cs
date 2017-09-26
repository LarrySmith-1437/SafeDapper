using System;
using System.CodeDom;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using SafeDapper;

namespace SafeDapperTest
{
    [TestClass]
    public class SafeQueryTests
    {
        private string _testFile = "TestData.xlsx";
        [TestMethod]
        [ExpectedException(typeof(DapperObjectMappingException))]
        public void SafeQueryThrowsExceptionWhenColumnNotMapped()
        {
            var currentAssemblyDirectory =Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); // System.Reflection.Assembly.GetEntryAssembly();
            var fullPath = Path.Combine(currentAssemblyDirectory, _testFile);
            var connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{fullPath}\";Extended Properties=\"Excel 12.0 Xml;HDR=Yes;IMEX=1;\"";

            var exists = File.Exists(fullPath);
            using (var conn = new OleDbConnection(connString))
            {
                conn.Open();
                var results = conn.SafeQuery<TestDataPropertyMismatch>($"Select * from [Sheet1$]", null, null, false,
                    commandType: CommandType.Text).ToList();

            }
        }

        [TestMethod]
        public void SafeQueryExecutesAndGetsData()
        {
            var currentAssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); // System.Reflection.Assembly.GetEntryAssembly();
            var fullPath = Path.Combine(currentAssemblyDirectory, _testFile);
            var connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{fullPath}\";Extended Properties=\"Excel 12.0 Xml;HDR=Yes;IMEX=1;\"";

            var exists = File.Exists(fullPath);
            using (var conn = new OleDbConnection(connString))
            {
                conn.Open();
                var results = conn.SafeQuery<TestData>($"Select * from [Sheet1$]", null, null, false,
                    commandType: CommandType.Text).ToList();

                Assert.AreEqual(3, results.Count);
            }
        }

    }
}
