using Microsoft.VisualStudio.Shell;

namespace CnSharp.VisualStudio.Extensions
{
    public interface IToolWindowAccessor
    {
        T NewWindow<T>(WindowInfo win) where T : ToolWindowPane;

        T FindWindows<T>(bool create, string uri = null) where T : ToolWindowPane;
    }
}