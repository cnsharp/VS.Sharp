using System;
using System.Collections.Generic;
using System.Linq;
using EnvDTE;
using EnvDTE80;

namespace CnSharp.VisualStudio.Extensions
{
    public class WindowAccessor : IWindowAccessor
    {
        public Windows2 Windows { get; set; }


        public  Window FindWindowByGuid(string guid)
        {
            for (int i = 1; i <= Windows.Count; i++)
            {
                if (Windows.Item(i).ObjectKind == guid)
                    return Windows.Item(i);
            }

            return null;
        }

        public  Window FindWindowByCaption(string caption)
        {
            for (int i = 1; i <= Windows.Count; i++)
            {
                if (Windows.Item(i).Caption == caption)
                    return Windows.Item(i);
            }

            return null;
        }

        public  Window FindWindow(Func<Window, bool> expression)
        {
            var windowList = new List<Window>();

            for (int i = 1; i <= Windows.Count; i++)
            {
                windowList.Add(Windows.Item(i));
            }

            return windowList.FirstOrDefault(expression);
        }

        /// <summary>
        ///     output message to Output Window
        /// </summary>
        /// <param name="paneName"></param>
        /// <param name="message"></param>
        public  void WriteToOutputWindow(string paneName, string message)
        {

            var outputWindow = (OutputWindow) Windows.Item("Output").Object;

            OutputWindowPane pane;
            try
            {
                pane = outputWindow.OutputWindowPanes.Item(paneName);
            }
            catch
            {
                pane = outputWindow.OutputWindowPanes.Add(paneName);
            }


            outputWindow.Parent.AutoHides = false;
            outputWindow.Parent.Activate();
            pane.Activate();

            // Add a line of text to the new pane.
            pane.OutputString(message + Environment.NewLine);
        }
    }
}