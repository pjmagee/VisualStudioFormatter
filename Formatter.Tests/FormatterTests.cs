namespace Formatter.Tests
{
    using System;
    using System.IO;
    using System.Linq;
    using global::Formatter;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [TestClass]
    public class FormatterTests
    {
        private VisualStudioFormatter visualStudioFormatter;

        private string GetPathFor(string fileName)
        {
            return Path.Combine(Environment.CurrentDirectory, "..\\..", "Files", fileName);
            // return Path.Combine(Environment.CurrentDirectory, "Files", fileName);
        }
        
        [TestInitialize]
        public void Initialize()
        {
            //visualStudioFormatter = new VisualStudioFormatter();
            visualStudioFormatter = new VisualStudioFormatter(version: 14);
        }

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            
        }

        [TestMethod]
        public void FormatDirectory()
        {
            visualStudioFormatter.FormatDirectory(@"C:\\Projects\\Time247\\Website\\@Modules");
        }

        [TestMethod]
        public void RazorFormatTest()
        {
            // Arrange
            var original = File.ReadAllText(GetPathFor("RazorFile.cshtml"));
            var modifiedPath = GetPathFor("RazorFile.Modified.cshtml");
            var modified = original.Replace("<p>", "\t<p>").Replace("var x = 1;", "var x = \t1");

            // Assert files aren't the same
            Assert.AreNotEqual(notExpected: original, actual: modified, ignoreCase: false);
            Assert.AreNotEqual(notExpected: original.Length, actual: modified.Length);

            // Act format modified
            File.WriteAllText(modifiedPath, contents: modified);
            visualStudioFormatter.Format(modifiedPath);
            modified = File.ReadAllText(modifiedPath);

            // Assert same as original
            Assert.AreEqual(expected: original, actual: modified, ignoreCase: false);
            Assert.AreEqual(expected: original.Length, actual: modified.Length);
        }

        [TestMethod]
        public void HtmlFormatTest()
        {
            // Arrange
            var original = File.ReadAllText(GetPathFor("HtmlFile.html"));
            var modifiedPath = GetPathFor("HtmlFile.Modified.html");
            var modified = original.Replace("<head>", "\t<head>").Replace("<title>", "\t<title>");

            // Assert files aren't the same
            Assert.AreNotEqual(notExpected: original, actual: modified, ignoreCase: false);
            Assert.AreNotEqual(notExpected: original.Length, actual: modified.Length);

            // Act format modified
            File.WriteAllText(modifiedPath, contents: modified);
            visualStudioFormatter.Format(modifiedPath);
            modified = File.ReadAllText(modifiedPath);

            // Assert same as original
            Assert.AreEqual(expected: original, actual: modified, ignoreCase: false);
            Assert.AreEqual(expected: original.Length, actual: modified.Length);
        }

        [TestMethod]
        public void RazorWithJsFormatTest()
        {
            // Arrange
            var original = File.ReadAllText(GetPathFor("RazorJsFile.cshtml"));
            var modifiedPath = GetPathFor("RazorJsFile.Modified.cshtml");

            var modified = original
                .Replace("var helloWorld = function()", "    var     helloWorld =    function()")
                .Replace("<p style=", "<p       style=")
                .Replace("@{", "    @{");

            // Assert files aren't the same
            Assert.AreNotEqual(notExpected: original, actual: modified, ignoreCase: false);
            Assert.AreNotEqual(notExpected: original.Length, actual: modified.Length);

            // Act format modified
            File.WriteAllText(modifiedPath, contents: modified);
            visualStudioFormatter.Format(modifiedPath);
            modified = File.ReadAllText(modifiedPath);

            // Assert same as original
            Assert.AreEqual(expected: original, actual: modified, ignoreCase: false);
            Assert.AreEqual(expected: original.Length, actual: modified.Length);
        }

        [TestMethod]
        public void CSharpTest()
        {
            // Arrange
            var original = File.ReadAllText(GetPathFor("CSharpFile.cs"));
            var modifiedPath = GetPathFor("CSharpFile.Modified.cs");
            var modified = original.Replace("public", "\tpublic").Replace("get; set;", "get;\tset;");

            // Assert files aren't the same
            Assert.AreNotEqual(notExpected: original, actual: modified, ignoreCase: false);
            Assert.AreNotEqual(notExpected: original.Length, actual: modified.Length);

            // Act format modified
            File.WriteAllText(modifiedPath, contents: modified);
            visualStudioFormatter.Format(modifiedPath);
            modified = File.ReadAllText(modifiedPath);

            // Assert same as original
            Assert.AreEqual(expected: original, actual: modified, ignoreCase: false);
            Assert.AreEqual(expected: original.Length, actual: modified.Length);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (visualStudioFormatter == null) return;
            visualStudioFormatter.Dispose();
            visualStudioFormatter = null;
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            var processes = System.Diagnostics.Process.GetProcesses().Where(p => p.ProcessName.Contains("devenv")).ToList();

            foreach (System.Diagnostics.Process process in processes)
            {
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }
    }
}
