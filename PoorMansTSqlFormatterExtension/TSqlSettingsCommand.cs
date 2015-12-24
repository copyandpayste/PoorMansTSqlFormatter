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

namespace PoorMansTSqlFormatterExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TSqlSettingsCommand
    {
        private Command _formatCommand;

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
            //GetFormatHotkey();
            SettingsForm settings = new SettingsForm(Properties.Settings.Default, Assembly.GetExecutingAssembly(), _generalResourceManager.GetString("ProjectAboutDescription"), new SettingsForm.GetTextEditorKeyBindingScopeName(GetTextEditorKeyBindingScopeName));
            if (settings.ShowDialog() == DialogResult.OK)
            {
                //SetFormatHotkey();
                _formattingManager = Utils.GetFormattingManager(Properties.Settings.Default);
            }
            settings.Dispose();
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


        /// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
        /// <param term='commandName'>The name of the command to execute.</param>
        /// <param term='executeOption'>Describes how the command should be run.</param>
        /// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
        /// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
        /// <param term='handled'>Informs the caller if the command was handled or not.</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if (commandName.Equals("PoorMansTSqlFormatterExtension.AddinConnector.FormatSelectionOrActiveWindow"))
                {
                    string fileExtension = System.IO.Path.GetExtension(_applicationObject.ActiveDocument.FullName);
                    bool isSqlFile = fileExtension.ToUpper().Equals(".SQL");

                    if (isSqlFile ||
                        MessageBox.Show(_generalResourceManager.GetString("FileTypeWarningMessage"), _generalResourceManager.GetString("FileTypeWarningMessageTitle"), MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        string fullText = SelectAllCodeFromDocument(_applicationObject.ActiveDocument);
                        TextSelection selection = (TextSelection)_applicationObject.ActiveDocument.Selection;
                        if (!selection.IsActiveEndGreater)
                            selection.SwapAnchor();
                        if (selection.Text.EndsWith(Environment.NewLine) || selection.Text.EndsWith(" "))
                            selection.CharLeft(true, 1); //newline counts as a distance of one.
                        string selectionText = selection.Text;
                        bool formatSelectionOnly = selectionText.Length > 0 && selectionText.Length != fullText.Length;
                        int cursorPoint = selection.ActivePoint.AbsoluteCharOffset;

                        string textToFormat = formatSelectionOnly ? selectionText : fullText;
                        bool errorsFound = false;
                        string formattedText = _formattingManager.Format(textToFormat, ref errorsFound);

                        bool abortFormatting = false;
                        if (errorsFound)
                            abortFormatting = MessageBox.Show(_generalResourceManager.GetString("ParseErrorWarningMessage"), _generalResourceManager.GetString("ParseErrorWarningMessageTitle"), MessageBoxButtons.YesNo) != DialogResult.Yes;

                        if (!abortFormatting)
                        {
                            if (formatSelectionOnly)
                            {
                                //if selection just delete/insert, so the active point is at the end of the selection
                                selection.Delete(1);
                                selection.Insert(formattedText, (int)EnvDTE.vsInsertFlags.vsInsertFlagsContainNewText);
                            }
                            else
                            {
                                //if whole doc then replace all text, and put the cursor approximately where it was (using proportion of text total length before and after)
                                int newPosition = (int)Math.Round(1.0 * cursorPoint * formattedText.Length / textToFormat.Length, 0, MidpointRounding.AwayFromZero);
                                ReplaceAllCodeInDocument(_applicationObject.ActiveDocument, formattedText);
                                ((TextSelection)(_applicationObject.ActiveDocument.Selection)).MoveToAbsoluteOffset(newPosition, false);
                            }
                        }
                    }

                    handled = true;
                    return;
                }
                if (commandName.Equals("PoorMansTSqlFormatterExtension.AddinConnector.FormattingOptions"))
                {
                    GetFormatHotkey();
                    SettingsForm settings = new SettingsForm(Properties.Settings.Default, Assembly.GetExecutingAssembly(), _generalResourceManager.GetString("ProjectAboutDescription"), new SettingsForm.GetTextEditorKeyBindingScopeName(GetTextEditorKeyBindingScopeName));
                    if (settings.ShowDialog() == DialogResult.OK)
                    {
                        SetFormatHotkey();
                        _formattingManager = Utils.GetFormattingManager(Properties.Settings.Default);
                    }
                    settings.Dispose();
                }
            }
        }

        private void GetFormatHotkey()
        {
            try
            {
                //TODO: Add support for multiple keybindings.
                string flatBindingsValue = "";
                var bindingArray = _formatCommand.Bindings as object[];
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
                //TODO: Add support for multiple keybindings.
                if (Properties.Settings.Default.Hotkey == null || Properties.Settings.Default.Hotkey.Trim() == "")
                    _formatCommand.Bindings = new object[0];
                else
                    _formatCommand.Bindings = Properties.Settings.Default.Hotkey;
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format(_generalResourceManager.GetString("HotkeyBindingFailureMessage"), Environment.NewLine, e.ToString()));
            }
        }

        //Nice clean methods avoiding slow selection-editing, from online post at:
        //  http://www.visualstudiodev.com/visual-studio-extensibility/how-can-i-edit-documents-programatically-22319.shtml
        private static string SelectAllCodeFromDocument(Document targetDoc)
        {
            string outText = "";
            TextDocument textDoc = targetDoc.Object("TextDocument") as TextDocument;
            if (textDoc != null)
                outText = textDoc.StartPoint.CreateEditPoint().GetText(textDoc.EndPoint);
            return outText;
        }

        private static void ReplaceAllCodeInDocument(Document targetDoc, string newText)
        {
            TextDocument textDoc = targetDoc.Object("TextDocument") as TextDocument;
            if (textDoc != null)
            {
                textDoc.StartPoint.CreateEditPoint().Delete(textDoc.EndPoint);
                textDoc.StartPoint.CreateEditPoint().Insert(newText);
            }
        }

    }
}
