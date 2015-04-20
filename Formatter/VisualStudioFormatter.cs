namespace Formatter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using EnvDTE;
    using EnvDTE80;

    /// <summary>
    /// https://msdn.microsoft.com/en-us/library/ms228755.aspx
    /// https://gist.github.com/pjmagee/d052fd1fe4c6995f6699
    /// http://www.pinvoke.net/default.aspx/user32.getwindowthreadprocessid
    /// https://msdn.microsoft.com/en-us/library/ms228772(v=VS.100).aspx
    /// https://msdn.microsoft.com/en-us/library/ms165618.aspx
    /// https://msdn.microsoft.com/en-us/library/envdte._dte.executecommand.aspx?cs-save-lang=1&cs-lang=csharp#code-snippet-1
    /// https://msdn.microsoft.com/en-us/library/za2b25t3.aspx
    /// http://blogs.msdn.com/b/kirillosenkov/archive/2011/08/10/how-to-get-dte-from-visual-studio-process-id.aspx
    /// </summary>
    public class VisualStudioFormatter : IDisposable
    {
        private readonly int _version;
        private Solution2 solution;
        private DTE2 dte2;
        private Project project;
        private int processId;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        private string TempPath
        {
            get { return Path.Combine(Path.GetTempPath(), "@MSharp.Temp", "Formatting"); }
        }

        private string VsVersion
        {
            get { return string.Format("VisualStudio.Solution.{0}.0", _version); }
        }

        private string DteVersion
        {
            get { return string.Format("VisualStudio.DTE.{0}.0", _version); }
        }

        private DTE2 DTE2
        {
            get
            {
                if(dte2 != null) return dte2;

                dte2 = GetDTE();
                dte2.MainWindow.Activate();

                return dte2;
            }
        }

        private Project Project
        {
            get { return project ?? (project = CreateProject()); }
        }

        private CommandWindow Command
        {
            get { return DTE2.ToolWindows.CommandWindow; }
        }

        private VisualStudioFormatter()
        {
            MessageFilter.Register();
        }

        public VisualStudioFormatter(int version) : this()
        {
            _version = version;
        }

        public void Format(string file)
        {
            Open(file);
            FormatDocument(file);
            SaveDocuments();
        }

        public void FormatDirectory(string directory)
        {
            var di = new DirectoryInfo(directory);
            if (!di.Exists) return;

            var files = di.GetFiles("*.*", SearchOption.AllDirectories).ToList();

            CreateSolution();
            CreateProject();
            OpenFiles(files);
            FormatDocuments();
            SaveDocuments();
        }

        private void CreateSolution()
        {
            DTE2.Solution.Create(TempPath, "FormattingSolution");
        }

        private void OpenFiles(List<FileInfo> files)
        {
            foreach (FileInfo file in files)
            {
                Project.ProjectItems.AddFromTemplate(file.FullName, file.Name);
            }

            while (DTE2.Documents.Count < files.Count)
            {
                // wait to open them all
            }
        }

        private Project CreateProject()
        {
            var template = @"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\ProjectTemplates\CSharp\Windows\1033\ConsoleApplication\consoleapplication.csproj";

            return DTE2.Solution.AddFromTemplate(template, TempPath, "Formatting", false);
        }

        private void SaveDocuments()
        {
            Command.SendInput("File.SaveAll", true);
            Command.SendInput("Window.CloseAllDocuments", true);

            while (DTE2.Documents.Count != 0)
            {
                // Wait to save and close all
            }
        }

        private void FormatDocument(string file)
        {
            Console.WriteLine("Formatting: {0}", file);
            
            while (DTE2.ActiveDocument.Saved) // try format document
            {
                DTE2.ActiveDocument.Activate();
                Command.SendInput("Edit.FormatDocument", true);
            }
        }

        private void FormatDocuments()
        {
            foreach (ProjectItem item in Project.ProjectItems)
            {
                var document = item.Document;
                
                while (document.Saved)
                {
                    document.Activate();
                    Console.WriteLine("Formatting: {0}", document.Name);
                    Command.SendInput("Edit.FormatDocument", true);
                }
            }
        }

        private void Open(string file)
        {
            Console.WriteLine("Opening: {0}.", file);

            // File.Open in 12, File.OpenFile in 14
            Command.SendInput(string.Format("File.OpenFile {0}", file), true);

            Console.WriteLine("Waiting for Visual Studio to open file.");

            while (DTE2.Documents.Count == 0)
            {
            }

            Console.WriteLine("File open.");
        }

        /// <summary>
        /// Create a new Visual Studio Instance for File formatting.
        /// </summary>
        private void LoadVisualStudio()
        {
            if (dte2 != null) 
                throw new Exception("DTE2 already exists for an existing Visual Studio Instance");

            solution = Activator.CreateInstance(Type.GetTypeFromProgID(VsVersion)) as Solution2;
            Console.WriteLine("VS HWND: {0}", solution.DTE.MainWindow.HWnd);
            GetWindowThreadProcessId(new IntPtr(solution.DTE.MainWindow.HWnd), out processId);
            Console.WriteLine("PID: {0}", processId);
        }

        /// <summary>
        /// Get the DTE which is associated to the new visual studio instance we created using Monkiker with ROT
        /// </summary>
        public DTE2 GetDTE()
        {
            LoadVisualStudio();

            string rotEntry = String.Format("!{0}:{1}", DteVersion, processId);
            object runningObject = null;

            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(0, out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;

                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];
                    string name = null;

                    try
                    {
                        if (runningObjectMoniker != null)
                        {
                            runningObjectMoniker.GetDisplayName(bindCtx, null, out name);
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Do nothing, there is something in the ROT that we do not have access to.
                    }

                    if (!string.IsNullOrEmpty(name) && string.Equals(name, rotEntry, StringComparison.Ordinal))
                    {
                        Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
                        break;
                    }
                }
            }
            finally
            {
                if (enumMonikers != null) Marshal.ReleaseComObject(enumMonikers);
                if (rot != null) Marshal.ReleaseComObject(rot);
                if (bindCtx != null) Marshal.ReleaseComObject(bindCtx);
            }

            return (DTE2)runningObject;
        }

        public void Dispose()
        {
            Command.SendInput("File.Exit", true);
            // DTE2.Quit();

            MessageFilter.Revoke();
        }
    }
}
