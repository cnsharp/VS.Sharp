namespace CnSharp.VisualStudio.Extensions
{
    public interface IDocumentAccessor
    {
        string GetActiveDocumentSelectionText();

        string GetActiveDocumentText();

        void NewDocument(string text);

        bool ReplaceSelection(string text);

        bool Insert(string text);

        void Overwrite(string text);
    }
}
