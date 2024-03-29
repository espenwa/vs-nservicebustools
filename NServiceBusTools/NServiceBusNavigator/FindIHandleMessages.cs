﻿using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace NServiceBusNavigator
{
    internal sealed class FindIHandleMessages
    {
        public const int CommandId = 0x0102;

        public static readonly Guid CommandSet = new Guid("6850066e-e30e-49cc-a144-066cc01f2944");

        private readonly AsyncPackage package;

        private FindIHandleMessages(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static FindIHandleMessages Instance
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
            // Switch to the main thread - the call to AddCommand in FindIAmStartedByMessages's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new FindIHandleMessages(package, commandService);
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
            ServiceHelper.FindPattern($"IHandleMessages\\s*<\\s*{ServiceHelper.GetSelectedText(ServiceProvider)}\\s*>", ServiceProvider);
        }
    }
}
