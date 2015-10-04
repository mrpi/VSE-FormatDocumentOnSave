using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using EnvDTE;
using Microsoft.VisualStudio.TextManager.Interop;

namespace Elders.VSE_FormatDocumentOnSave
{
    internal class DocumentFormatter
    {
        private readonly IVsTextManager txtMngr;
        private readonly DTE dte;
        private readonly bool clangFormatPluginIsInstalled;

        public DocumentFormatter(IVsTextManager txtMngr, DTE dte)
        {
            this.txtMngr = txtMngr;
            this.dte = dte;
            this.clangFormatPluginIsInstalled = CheckIfClangFormatPluginIsInstalled();
        }

        private bool IsCppFile(string fileName)
        {
            return fileName.EndsWith(".cpp", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".hpp", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".cxx", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".hxx", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".c++", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".h++", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".cc", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".hh", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".c", true, CultureInfo.InvariantCulture)
            || fileName.EndsWith(".h", true, CultureInfo.InvariantCulture);
        }

        private bool CheckIfClangFormatPluginIsInstalled()
        {
            try
            {
                if (dte.Commands.Item("Tools.ClangFormat") != null)
                    return true;
            }
            catch (Exception)
            {
            }
            return false;
        }

        public void FormatCurrentActiveDocument()
        {
            try
            {
                var fileName = dte.ActiveDocument.ProjectItem.Name;
                bool isCppFile = IsCppFile(fileName);
                bool isCsHtmlFile = fileName.EndsWith(".cshtml", true, CultureInfo.InvariantCulture);
                bool isCSharpFile = fileName.EndsWith(".cs", true, CultureInfo.InvariantCulture);

                if (dte.ActiveWindow.Kind == "Document")
                {
                    if (clangFormatPluginIsInstalled && isCppFile)
                        FormatClang();
                    else if (isCsHtmlFile)
                        FormatCSHTML();
                    else if (isCSharpFile || isCppFile)
                        dte.ExecuteCommand("Edit.FormatDocument", string.Empty);
                }
            }
            catch (Exception) { }
        }

        private void FormatClang()
        {
            IVsTextView textViewCurrent;
            txtMngr.GetActiveView(1, null, out textViewCurrent);    // Gets the TextView (TextEditor) for the current active document
            int a, b, c, verticalScrollPosition;
            textViewCurrent.GetScrollInfo(1, out a, out b, out c, out verticalScrollPosition);

            dynamic selection = dte.ActiveDocument.Selection;
            int line = selection.CurrentLine;
            int lineLength = selection.ActivePoint.LineLength;
            int col = selection.CurrentColumn;

            dte.ExecuteCommand("Edit.SelectAll", string.Empty);
            dte.ExecuteCommand("Tools.ClangFormat", string.Empty);

            if (!selection.IsEmpty)
                selection.Cancel();

            selection.GoToLine(line);
            int offset = col - (lineLength - selection.ActivePoint.LineLength);
            selection.MoveToLineAndOffset(line, offset, false);

            textViewCurrent.SetScrollPosition(1, verticalScrollPosition);
        }

        private void FormatCSHTML()
        {
            IVsTextView textViewCurrent;
            txtMngr.GetActiveView(1, null, out textViewCurrent);    // Gets the TextView (TextEditor) for the current active document
            int a, b, c, verticalScrollPosition;
            textViewCurrent.GetScrollInfo(1, out a, out b, out c, out verticalScrollPosition);

            dynamic selection = dte.ActiveDocument.Selection;
            int line = selection.CurrentLine;
            int lineLength = selection.ActivePoint.LineLength;
            int col = selection.CurrentColumn;

            dte.ExecuteCommand("Edit.FormatDocument", string.Empty);

            if (!selection.IsEmpty)
                selection.Cancel();

            selection.GoToLine(line);
            int offset = col - (lineLength - selection.ActivePoint.LineLength);
            selection.MoveToLineAndOffset(line, offset, false);

            textViewCurrent.SetScrollPosition(1, verticalScrollPosition);
        }

        public void FormatDocuments(IEnumerable<Document> documents)
        {
            var currentDoc = dte.ActiveDocument;
            foreach (var doc in documents)
            {
                doc.Activate();
                FormatCurrentActiveDocument();
            }
            currentDoc.Activate();
        }
        public void FormatNonSavedDocuments()
        {
            FormatDocuments(GetNonSavedDocuments());
        }

        IEnumerable<Document> GetNonSavedDocuments()
        {
            return dte.Documents.OfType<Document>().Where(document => !document.Saved);
        }
    }
}