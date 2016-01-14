//------------------------------------------------------------------------------
// <copyright file="TSqlFormatCommand.cs" company="Company">
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
using PoorMansTSqlFormatterLib.Formatters;

namespace PoorMansTSqlFormatterExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TSqlFormatCommand
    {
        public static DTE applicationObject = Package.GetGlobalService(typeof(DTE)) as DTE;
        public static Events events = applicationObject.Events;
        public static DocumentEvents documentEvents = applicationObject.Events.get_DocumentEvents(null);

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4129;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("c5c4ef0c-afc8-44ab-a333-1168a23caa38");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TSqlFormatCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TSqlFormatCommand(Package package)
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
        public static TSqlFormatCommand Instance
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

        private static bool SavingDocument = false;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new TSqlFormatCommand(package);
            documentEvents.DocumentSaved += new _dispDocumentEvents_DocumentSavedEventHandler(DocumentEvents_DocumentSaved);
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
            if (applicationObject.ActiveDocument != null)
            {
                FormatDocument(applicationObject.ActiveDocument);
            }
        }

        public static void FormatDocument(Document document) {
            PoorMansTSqlFormatterLib.SqlFormattingManager formattingManager = Utils.GetFormattingManager(Properties.Settings.Default);
            ResourceManager generalResourceManager = new ResourceManager("PoorMansTSqlFormatterExtension.GeneralLanguageContent", Assembly.GetExecutingAssembly());

        string fileExtension = System.IO.Path.GetExtension(document.FullName);
            bool isSqlFile = fileExtension.ToUpper().Equals(".SQL");

            if (isSqlFile ||
                MessageBox.Show(generalResourceManager.GetString("FileTypeWarningMessage"), generalResourceManager.GetString("FileTypeWarningMessageTitle"), MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string fullText = SelectAllCodeFromDocument(document);
                TextSelection selection = (TextSelection)document.Selection;
                if (!selection.IsActiveEndGreater)
                    selection.SwapAnchor();
                if (selection.Text.EndsWith(Environment.NewLine) || selection.Text.EndsWith(" "))
                    selection.CharLeft(true, 1); //newline counts as a distance of one.
                string selectionText = selection.Text;
                bool formatSelectionOnly = selectionText.Length > 0 && selectionText.Length != fullText.Length;
                int cursorPoint = selection.ActivePoint.AbsoluteCharOffset;

                string textToFormat = formatSelectionOnly ? selectionText : fullText;
                bool errorsFound = false;
                string formattedText = formattingManager.Format(textToFormat, ref errorsFound);

                bool abortFormatting = false;
                if (errorsFound)
                    abortFormatting = MessageBox.Show(generalResourceManager.GetString("ParseErrorWarningMessage"), generalResourceManager.GetString("ParseErrorWarningMessageTitle"), MessageBoxButtons.YesNo) != DialogResult.Yes;

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
                        ReplaceAllCodeInDocument(document, formattedText);
                        ((TextSelection)(document.Selection)).MoveToAbsoluteOffset(newPosition, false);
                    }
                }
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

        public static void DocumentEvents_DocumentSaved(Document Document)
        {
            TSqlStandardFormatter formatter = (TSqlStandardFormatter)Utils.GetFormattingManager(Properties.Settings.Default).Formatter;
            if (SavingDocument || !formatter.Options.FormatOnSave)
                return;

            string fileExtension = System.IO.Path.GetExtension(Document.FullName);
            bool isSqlFile = fileExtension.ToUpper().Equals(".SQL");

            if (isSqlFile)
            {
                FormatDocument(Document);

                SavingDocument = true;
                Document.Save(Document.FullName);
                SavingDocument = false;
            }
        }
    }
}
