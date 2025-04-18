﻿using System;
using System.Collections.Generic;
using CnSharp.VisualStudio.Extensions.Commands;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CnSharp.VisualStudio.Extensions
{
    public class Host
    {
        private static Host _host;
        private bool _closingLast;
        private DocumentEvents _documentEvents;

        private _DTE _dte;
        private DTEEvents _dteEvents;
        private SolutionEvents _solutionEvents;


        static Host()
        {
            Plugins = new List<Plugin>();
        }

        public static Host Instance => _host ?? (_host = new Host());

        //public Package Package { get; set; }
        //public AsyncPackage AsyncPackage { get; set; }

        public T GetService<T>() where T : class
        {
            return AsyncPackage.GetGlobalService(typeof(T)) as T;
        }

        public IToolWindowAccessor ToolWindowAccessor { get; set; }

        //public Assembly Assembly { get; set; }

        public string Location { get; set; }

        public string FileName { get; set; }

        public string Version { get; set; }

        public _DTE DTE
        {
            get => _dte;
            set
            {
                _dte = value;
                if (_dte != null)
                {
                    _documentEvents = _dte.Events.DocumentEvents;
                    _documentEvents.DocumentOpened += DocumentEvents_DocumentOpened;
                    _documentEvents.DocumentClosing += DocumentEvents_DocumentClosing;

                    _solutionEvents = _dte.Events.SolutionEvents;
                    _solutionEvents.Opened += SolutionEvents_Opened;
                    _solutionEvents.AfterClosing += SolutionEvents_AfterClosing;

                    _dteEvents = _dte.Events.DTEEvents;
                    _dteEvents.OnStartupComplete += DteEvents_OnStartupComplete;
                    _dteEvents.OnBeginShutdown += DteEvents_OnBeginShutdown;
                }
            }
        }

        public DTE2 Dte2 => (DTE2) DTE;

        public Windows2 Windows => Dte2.Windows as Windows2;

        public Solution Solution => DTE.Solution;

        public Solution2 Solution2 =>  Solution as Solution2;

        public IVsStatusbar Statusbar { get; set; }

        //public IVsOutputWindow OutputWindow { get; set; }


        public Action SolutionOpenedAction { get; set; }

        public Action AfterSolutionClosingAction { get; set; }

        public Action StartupCompleteAction { get; set; }

        public Action BeginShutdownAction { get; set; }

        public Action<Document> DocumentOpenedAction { get; set; }

        public Action<Document> DocumentClosingAction { get; set; }

        //public ResourceManager ResourceManager { get; set; }
        public static List<Plugin> Plugins { get; set; }

        public void ShowOutputWindow()
        {
            var outputWindow = (OutputWindow)Windows.Item("Output").Object;
            outputWindow.Parent.AutoHides = false;
            outputWindow.Parent.Activate();
        }

        /// <summary>
        ///     output message to Output Window
        /// </summary>
        /// <param name="paneName"></param>
        /// <param name="message"></param>
        public void WriteToOutputWindow(string paneName, string message)
        {

            var outputWindow = (OutputWindow)Windows.Item("Output").Object;

            OutputWindowPane pane;
            try
            {
                pane = outputWindow.OutputWindowPanes.Item(paneName);
            }
            catch
            {
                pane = outputWindow.OutputWindowPanes.Add(paneName);
            }
            pane.Activate();
            // Add a line of text to the new pane.
            pane.OutputString(message + Environment.NewLine);
        }

        private void DteEvents_OnStartupComplete()
        {
            StartupCompleteAction?.Invoke();
        }

        private void DteEvents_OnBeginShutdown()
        {
            BeginShutdownAction?.Invoke();
        }

        private void SolutionEvents_AfterClosing()
        {
            foreach (var plugin in Plugins)
            {
                plugin.CommandManager.ApplyDependencies(DependentItems.SolutionProject, false);
            }
            AfterSolutionClosingAction?.Invoke();
        }

        private void SolutionEvents_Opened()
        {
            foreach (var plugin in Plugins)
            {
                plugin.CommandManager.ApplyDependencies(DependentItems.SolutionProject, true);
            }
            SolutionOpenedAction?.Invoke();
        }

        private void DocumentEvents_DocumentClosing(Document document)
        {
            if (_dte.Documents.Count == 1) //the last one
            {
                _closingLast = true;
                foreach (var plugin in Plugins)
                {
                    plugin.CommandManager.ApplyDependencies(DependentItems.Document, false);
                }
                _closingLast = false;
            }
            DocumentClosingAction?.Invoke(document);
        }

        private void DocumentEvents_DocumentOpened(Document document)
        {
            foreach (var plugin in Plugins)
            {
                plugin.CommandManager.ApplyDependencies(DependentItems.Document, true);
            }
            DocumentOpenedAction?.Invoke(document);
        }

        public bool IsDependencySatisfied(DependentItems dependentItems)
        {
            var dependencies = dependentItems.ToString()
                .Split(new[] {" ,"}, StringSplitOptions.RemoveEmptyEntries);
            foreach (var dependency in dependencies)
            {
                var item = (DependentItems) Enum.Parse(typeof (DependentItems), dependency);
                if (!CheckDependency(item))
                    return false;
            }
            return true;
        }

        private bool CheckDependency(DependentItems dependentItems)
        {
            switch (dependentItems)
            {
                case DependentItems.Document:
                    return _dte.Documents.Count > 0 && !_closingLast;
                case DependentItems.SolutionProject:
                    return _dte.Solution != null && _dte.Solution.Projects.Count > 0;
            }
            return true;
        }
    }
}