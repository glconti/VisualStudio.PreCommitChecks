﻿#region using

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

        /// <summary>
        ///   VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        ///   Initializes a new instance of the <see cref="CleanupDirtyFilesCommand" /> class.
        ///   Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CleanupDirtyFilesCommand(Package package)
        {
            if (package == null) throw new ArgumentNullException(nameof(package));

            this.package = package;

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
        private IServiceProvider ServiceProvider => package;

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
            var extensionToUpdate = new[] { ".cs", ".xaml", ".resx", ".config" };

            var dte = (DTE)ServiceProvider.GetService(typeof(DTE));

            var solutionPath = Path.GetDirectoryName(dte.Solution.FullName);

            if (!Repository.IsValid(solutionPath)) return;

            dte.Documents.SaveAll();

            var filesToCleanUp = new HashSet<string>();

            using (var repo = new Repository(solutionPath))
            {
                var repositoryStatus = repo.RetrieveStatus();

                if (!repositoryStatus.IsDirty) return;

                var dirtyFiles = repositoryStatus.Added.Union(repositoryStatus.Modified).Select(statusEntry => statusEntry.FilePath);

                foreach (var dirtyFile in dirtyFiles)
                {
                    filesToCleanUp.Add(dirtyFile);
                }
            }

            var enumerable = filesToCleanUp.Where(f => extensionToUpdate.Any(ext => f.EndsWith(ext, StringComparison.InvariantCultureIgnoreCase))).ToList();

            foreach (var file in enumerable)
            {
                dte.ExecuteCommand("ReSharper_SilentCleanupCode");
            }

            dte.Documents.SaveAll();
        }
    }
}