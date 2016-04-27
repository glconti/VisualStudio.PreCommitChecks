#region using

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using EnvDTE;
using LibGit2Sharp;
using Microsoft.VisualStudio.Shell;

#endregion

namespace VSPreCommitChecks.Command
{
    /// <summary>
    ///   Command handler
    /// </summary>
    internal sealed class CleanupDirtyFilesCommand
    {
        /// <summary>
        ///   Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        ///   Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("476c4c16-bfa5-4e13-9e55-3b213cefa85e");

        private static readonly string[] ExtensionsToFormat = { ".cs", ".xaml", ".resx", ".xml", ".config" };
        private static readonly string[] ExtensionsNotToFormat = { ".designer.cs" };
        private readonly Dictionary<string, DateTime> _lastFormattedFiles = new Dictionary<string, DateTime>();

        /// <summary>
        ///   VS Package that provides this command, not null.
        /// </summary>
        private readonly Package _package;
        private string _lastFormattedBranch;

        /// <summary>
        ///   Initializes a new instance of the <see cref="CleanupDirtyFilesCommand" /> class.
        ///   Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CleanupDirtyFilesCommand(Package package)
        {
            if (package == null) throw new ArgumentNullException(nameof(package));

            _package = package;

            var commandService = ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

            commandService?.AddCommand(new MenuCommand(MenuItemCallback, new CommandID(CommandSet, CommandId)));
        }

        /// <summary>
        ///   Gets the instance of the command.
        /// </summary>
        public static CleanupDirtyFilesCommand Instance { get; private set; }

        /// <summary>
        ///   Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider => _package;

        private string LastFormattedBranch
        {
            get { return _lastFormattedBranch; }
            set
            {
                if (value == _lastFormattedBranch) return;
                _lastFormattedBranch = value;

                _lastFormattedFiles.Clear();
            }
        }

        /// <summary>
        ///   Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new CleanupDirtyFilesCommand(package);
        }

        /// <summary>
        ///   This function is the callback used to execute the command when the menu item is clicked.
        ///   See the constructor to see how the menu item is associated with this function using
        ///   OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte = (DTE)ServiceProvider.GetService(typeof(DTE));

            var directoryName = Path.GetDirectoryName(dte.Solution.FullName);
            if (directoryName == null) return;

            var solutionPath = directoryName.ToLower();
            if (!Repository.IsValid(solutionPath)) return;

            dte.Documents.SaveAll();

            var filesToCleanup = GetDirtyFiles(solutionPath);
            if (filesToCleanup.Count < 1) return;

            var previouslyActiveDocument = dte.ActiveDocument;

            foreach (var file in filesToCleanup)
            {
                var isClosed = !dte.ItemOperations.IsFileOpen(file);

                var document = isClosed
                                   ? dte.ItemOperations.OpenFile(file, Constants.vsViewKindTextView).Document
                                   : dte.Documents.Cast<Document>().FirstOrDefault(d => d.FullName.ToLower() == file);

                if (document == null) continue;

                document.Activate();

                bool canSave;

                try
                {
                    dte.ExecuteCommand("ReSharper.ReSharper_SilentCleanupCode");
                    canSave = true;
                }
                catch
                {
                    canSave = false;
                }

                if (canSave && !document.Saved) document.Save();
                if (isClosed) document.Close();

                _lastFormattedFiles[file] = DateTime.Now;
            }

            previouslyActiveDocument?.Activate();
        }

        private bool IsNotFormatted(string file)
        {
            var lastModified = File.GetLastWriteTime(file);

            DateTime lastFormatted;
            if (_lastFormattedFiles.TryGetValue(file, out lastFormatted)) return lastModified > lastFormatted;

            _lastFormattedFiles.Add(file, DateTime.MinValue);
            return true;
        }

        private List<string> GetDirtyFiles(string solutionPath)
        {
            using (var repo = new Repository(solutionPath))
            {
                var repositoryStatus = repo.RetrieveStatus();
                if (!repositoryStatus.IsDirty) return new List<string>();

                LastFormattedBranch = repo.Head.FriendlyName.ToLower();

                return repositoryStatus.Staged
                                       .Union(repositoryStatus.Untracked)
                                       .Union(repositoryStatus.Added)
                                       .Union(repositoryStatus.Modified)
                                       .Select(statusEntry => statusEntry.FilePath.ToLower())
                                       .Distinct()
                                       .Where(f => !ExtensionsNotToFormat.Any(f.EndsWith) && ExtensionsToFormat.Any(f.EndsWith))
                                       .Select(f => Path.Combine(solutionPath, f))
                                       .Where(IsNotFormatted)
                                       .ToList();
            }
        }
    }
}