using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;

namespace CnSharp.VisualStudio.Extensions
{
    public class WindowAccessor : IWindowAccessor
    {
        public Windows2 Windows { get; set; }

        [Obsolete]
        public AddIn AddIn { get; set; }

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

        [Obsolete]
        public  object AddWindow(WindowInfo window)
        {
            var myWindow = Windows;
            string assemblypath = window.AssemblyLocation;
            AddIn addinobj = AddIn;


            Object obj = null;
            Window newWinobj = myWindow.CreateToolWindow2(addinobj, assemblypath,
                window.ClassName, window.Caption, window.Guid, ref obj);

            newWinobj.Visible = true;
            if (window.Picture != null)
                newWinobj.SetTabPicture(window.Picture);
            //以下两句实现文档式停靠
            if (window.IsDocument)
            {
                newWinobj.IsFloating = false;
                newWinobj.Linkable = false;
            }
            else
            {
                if (window.Width > 0)
                    newWinobj.Width = window.Width;
                if (window.Height > 0)
                    newWinobj.Height = window.Height;

                if (window.LinkToWindow != null)
                    myWindow.CreateLinkedWindowFrame(window.LinkToWindow, newWinobj,
                        vsLinkedWindowType.vsLinkedWindowTypeTabbed);
            }


            return obj;
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
            pane.OutputString(message + "\r\n");
        }

        [Obsolete]
        public  T NewWindow<T>(WindowInfo win) where T : UserControl
        {
            Window window = FindWindowByCaption(win.Caption);
            T view = null;
            if (window != null)
            {
                window.Activate();
                view = window.Object as T;
            }
            if (view == null)
            {
                win.ClassName = typeof (T).FullName;
                object obj = AddWindow(win);
                view = obj as T;
            }
            return view;
        }
    }
}