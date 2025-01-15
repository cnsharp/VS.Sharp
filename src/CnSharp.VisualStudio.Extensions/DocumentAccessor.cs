using EnvDTE;

namespace CnSharp.VisualStudio.Extensions
{
    public class DocumentAccessor : IDocumentAccessor
    {
        private static readonly _DTE Dte = Host.Instance.DTE;
        public string GetActiveDocumentSelectionText()
        {
            if (Dte.ActiveDocument == null)
                return null;
            return Dte.ActiveDocument.GetSelectionText();
        }

        public string GetActiveDocumentText()
        {
            return Dte.ActiveDocument.GetText();
        }

        public void NewDocument(string text)
        {
            var doc = (TextDocument)Dte.ActiveDocument.Object(null);
            doc.EndPoint.CreateEditPoint().Insert(text);
            doc.DTE.ActiveDocument.Activate();
        }

        public bool ReplaceSelection(string text)
        {
            if (Dte.ActiveDocument == null)
                return false;
            Dte.ActiveDocument.ReplaceSelection(text);
            return true;
        }

        public bool Insert(string text)
        {
            if (Dte.ActiveDocument == null)
                return false;
            Dte.ActiveDocument.Insert(text);
            return true;
        }

        public void Overwrite(string text)
        {
            if (Dte.ActiveDocument == null)
                return;
            Dte.ActiveDocument.Overwrite(text);
        }
    }
}
