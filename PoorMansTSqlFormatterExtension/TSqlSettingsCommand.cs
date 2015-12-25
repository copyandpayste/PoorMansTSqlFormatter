//------------------------------------------------------------------------------
// <copyright file="TSqlSettingsCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using PoorMansTSqlFormatterPluginShared;
using System.Windows.Forms;
using System.Reflection;
using EnvDTE;
using System.Resources;
using EnvDTE80;

namespace PoorMansTSqlFormatterExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TSqlSettingsCommand
    {
        private ResourceManager _generalResourceManager = new ResourceManager("PoorMansTSqlFormatterExtension.GeneralLanguageContent", Assembly.GetExecutingAssembly());
        private PoorMansTSqlFormatterLib.SqlFormattingManager _formattingManager = null;
        DTE _applicationObject = Package.GetGlobalService(typeof(DTE)) as DTE;


        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("c5c4ef0c-afc8-44ab-a333-1168a23caa38");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TSqlSettingsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TSqlSettingsCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TSqlSettingsCommand Instance
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
            Instance = new TSqlSettingsCommand(package);
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
            //TODO: pass the current settings to this form
            GetFormatHotkey();
            SettingsForm settings = new SettingsForm(Properties.Settings.Default, Assembly.GetExecutingAssembly(), _generalResourceManager.GetString("ProjectAboutDescription"), new SettingsForm.GetTextEditorKeyBindingScopeName(GetTextEditorKeyBindingScopeName));
            if (settings.ShowDialog() == DialogResult.OK)
            {
                //TODO: change this to use the current settings
                SetFormatHotkey();
                _formattingManager = Utils.GetFormattingManager(Properties.Settings.Default);
            }
            settings.Dispose();
        }

        private void GetFormatHotkey()
        {
            try
            {
                Command formatCommand = _applicationObject.Commands.Item("Tools.FormatTSql");
                string flatBindingsValue = "";
                var bindingArray = formatCommand.Bindings as object[];
                if (bindingArray != null && bindingArray.Length > 0)
                    flatBindingsValue = bindingArray[0].ToString();

                if (Properties.Settings.Default.Hotkey != flatBindingsValue)
                {
                    Properties.Settings.Default.Hotkey = flatBindingsValue;
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format(_generalResourceManager.GetString("HotkeyRetrievalFailureMessage"), Environment.NewLine, e.ToString()));
            }
        }

        private void SetFormatHotkey()
        {
            try
            {
                Command formatCommand = _applicationObject.Commands.Item("Tools.FormatTSql");
                if (Properties.Settings.Default.Hotkey == null || Properties.Settings.Default.Hotkey.Trim() == "")
                    formatCommand.Bindings = new object[0];
                else
                    formatCommand.Bindings = Properties.Settings.Default.Hotkey;
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format(_generalResourceManager.GetString("HotkeyBindingFailureMessage"), Environment.NewLine, e.ToString()));
            }
        }

        private string GetTextEditorKeyBindingScopeName()
        {
            string strScope = null;
            try
            {
                //"dirty hack" (as its author puts it) to get localized Text Editor scope name in 
                // non-english instalations - but it works! (without having access to "IVsShell")
                // Thank you Roland Weigelt! http://weblogs.asp.net/rweigelt/archive/2006/07/16/458634.aspx
                Command cmd = _applicationObject.Commands.Item("Edit.DeleteBackwards", -1);
                object[] arrBindings = (object[])cmd.Bindings;
                string strBinding = (string)arrBindings[0];
                strScope = strBinding.Substring(0, strBinding.IndexOf("::"));
            }
            catch (Exception ex)
            {
                //I know, general catch blocks are evil - but honestly, if that failed, what can we do?? I have no idea what types of issues to expect!
                MessageBox.Show(string.Format(_generalResourceManager.GetString("TextEditorScopeNameRetrievalFailureMessage"), Environment.NewLine, ex.ToString()));
            }
            return strScope;
        }
    }
}
