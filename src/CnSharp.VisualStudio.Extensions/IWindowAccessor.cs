using System;
using EnvDTE;
using EnvDTE80;

namespace CnSharp.VisualStudio.Extensions
{
    [Obsolete]
    public interface IWindowAccessor
    {
        Windows2 Windows { get; set; }

        Window FindWindowByGuid(string guid);

        Window FindWindowByCaption(string caption);

        Window FindWindow(Func<Window, bool> expression);

        void WriteToOutputWindow(string paneName, string message);
    }
}
