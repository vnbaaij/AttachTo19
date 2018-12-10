//------------------------------------------------------------------------------
// <copyright file="AttachTo.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace AttachTo
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AttachTo
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;
        public const int AttachToIISId = 0x100;
        public const int AttachToIISExpressId = 0x101;
        public const int AttachToNUnitId = 0x102;
        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid AttachToPackageCmdSet = new Guid("03b7f10f-7fbb-45f5-88cb-c9d93a45e7be");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="AttachTo"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private AttachTo(Package package)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.package = package ?? throw new ArgumentNullException("package");
            if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {

                AddAttachToCommand(commandService, AttachToIISId, gop => gop.ShowAttachToIIS, "w3wp.exe");
                AddAttachToCommand(commandService, AttachToIISExpressId, gop => gop.ShowAttachToIISExpress, "iisexpress.exe");
                AddAttachToCommand(commandService, AttachToNUnitId, gop => gop.ShowAttachToNUnit, "nunit-agent.exe", "nunit.exe", "nunit-console.exe", "nunit-agent-x86.exe", "nunit-x86.exe", "nunit-console-x86.exe");

                //var menuCommandID = new CommandID(AttachToPackageCmdSet, AttachToIISId);
                //var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                //commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AttachTo Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new AttachTo(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "AttachTo";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private void AddAttachToCommand(OleMenuCommandService commandService, uint commandId, Func<GeneralOptionsPage, bool> isVisible, params string[] programsToAttach)
        {
            OleMenuCommand menuItemCommand = new OleMenuCommand(
                delegate (object sender, EventArgs e)
                {
                    ThreadHelper.ThrowIfNotOnUIThread();
                    DTE dte = (DTE)ServiceProvider.GetService(typeof(DTE));
                    Assumes.Present(dte);
                    foreach (Process process in dte.Debugger.LocalProcesses)
                    {
                        if (programsToAttach.Any(p => { Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread(); return process.Name.EndsWith(p); }))
                        {
                            process.Attach();
                        }
                    }
                },
                new CommandID(AttachToPackageCmdSet, (int)commandId));
            menuItemCommand.BeforeQueryStatus += (s, e) => menuItemCommand.Visible = isVisible((GeneralOptionsPage)package.GetDialogPage(typeof(GeneralOptionsPage)));
            commandService.AddCommand(menuItemCommand);
        }
    }
}
