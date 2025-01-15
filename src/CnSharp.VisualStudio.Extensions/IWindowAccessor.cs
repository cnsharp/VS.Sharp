using System;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;

namespace CnSharp.VisualStudio.Extensions
{
    public interface IWindowAccessor
    {
        Windows2 Windows { get; set; }

        [Obsolete]
        AddIn AddIn { get; set; }

        Window FindWindowByGuid(string guid);

        Window FindWindowByCaption(string caption);

        Window FindWindow(Func<Window, bool> expression);

        object AddWindow(WindowInfo window);

        void WriteToOutputWindow(string paneName, string message);

        T NewWindow<T>(WindowInfo win) where T : UserControl;
    }
}
