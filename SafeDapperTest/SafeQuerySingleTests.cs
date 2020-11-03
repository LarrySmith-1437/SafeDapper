using Microsoft.VisualStudio.TestTools.UnitTesting;
using SafeDapper;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Reflection;

/// <summary>
/// If you don't have MS office installed, you'll likely need to install the office drivers to read the excel sheet test file w/out having actual office installed
/// https://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
/// or included in this checkin, a version that works on windows 7 x64 through Windows 10.
/// </summary>

namespace SafeDapperTest
{
    [TestClass]
    public class SafeQuerySingleTests
    {
        private string _testFile = "TestData.xlsx";
        [TestMethod]
        [ExpectedException(typeof(DapperObjectMappingException))]
        public void SafeQuerySingleThrowsExceptionWhenColumnNotMapped()
        {
            var currentAssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); // System.Reflection.Assembly.GetEntryAssembly();
            var fullPath = Path.Combine(currentAssemblyDirectory, _testFile);
            var connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{fullPath}\";Extended Properties=\"Excel 12.0 Xml;HDR=Yes;IMEX=1;\"";

            Assert.IsTrue(File.Exists(fullPath));
            using (var conn = new OleDbConnection(connString))
            {
                conn.Open();
                var result = conn.SafeQuerySingle<TestDataPropertyMismatch>($"Select top 1 * from [Sheet1$]", null, null,
                    commandType: CommandType.Text);

            }
        }

        [TestMethod]
        public void SafeQuerySingleExecutesAndGetsData()
        {
            var currentAssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); // System.Reflection.Assembly.GetEntryAssembly();
            var fullPath = Path.Combine(currentAssemblyDirectory, _testFile);
            var connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{fullPath}\";Extended Properties=\"Excel 12.0 Xml;HDR=Yes;IMEX=1;\"";

            Assert.IsTrue(File.Exists(fullPath));
            using (var conn = new OleDbConnection(connString))
            {
                conn.Open();
                var result = conn.SafeQuerySingle<TestData>($"Select top 1 * from [Sheet1$]", null, null,
                    commandType: CommandType.Text);

                Assert.IsNotNull(result);
                Assert.AreEqual(1, result.ColumnA);
                Assert.AreEqual(2, result.ColumnB);
                Assert.AreEqual(3, result.ColumnC);
            }
        }

    }
}
