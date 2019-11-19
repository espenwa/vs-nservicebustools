﻿using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace NServiceBusNavigator
{
    internal sealed class FindIAmStartedBy
    {
        public const int CommandId = 256;

        public static readonly Guid CommandSet = new Guid("f01665b9-2905-4505-afcd-0a671e28e57c");

        private readonly AsyncPackage package;

        private FindIAmStartedBy(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static FindIAmStartedBy Instance
        {
            get;
            private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in FindIAmStartedBy's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new FindIAmStartedBy(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Find all lines having "publish" or "Send"
            // ... followed by any number of spaces
            // then <
            // then any number of spaces
            // then the selected text
            // then any number of spaces
            // then >
            // then any text 
            // then end of line
            ServiceHelper.FindPattern($"^.*\\.(IAmStartedBy)\\s*<\\s*{ServiceHelper.GetSelectedText(ServiceProvider)}\\s*>*.$", ServiceProvider);
        }
    }
}
