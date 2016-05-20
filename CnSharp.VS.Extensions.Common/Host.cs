using System;
using System.Collections.Generic;
using CnSharp.VisualStudio.Extensions.Commands;
using CnSharp.VisualStudio.Extensions.SourceControl;
using EnvDTE;
using EnvDTE80;

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

        public static Host Instance
        {
            get { return _host ?? (_host = new Host()); }
        }

        //public Assembly Assembly { get; set; }

        public string Location { get; set; }

        public string FileName { get; set; }

        public string Version { get; set; }

        public _DTE DTE
        {
            get { return _dte; }
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
                    _dteEvents.OnBeginShutdown += _dteEvents_OnBeginShutdown;
                }
            }
        }

        public DTE2 Dte2
        {
            get { return (DTE2) DTE; }
        }

        public ISourceControl SourceControl { get; set; }

        public AddIn AddIn { get; set; }


        //public ResourceManager ResourceManager { get; set; }
        public static List<Plugin> Plugins { get; set; }

        private void _dteEvents_OnBeginShutdown()
        {
        }

        private void SolutionEvents_AfterClosing()
        {
            SourceControl = null;
            foreach (var plugin in Plugins)
            {
                plugin.CommandManager.ApplyDependencies(DependentItems.SolutionProject, false);
            }
        }

        private void SolutionEvents_Opened()
        {
            SourceControl = SourceControlManager.GetSolutionSourceControl(DTE.Solution);
            foreach (var plugin in Plugins)
            {
                plugin.CommandManager.ApplyDependencies(DependentItems.SolutionProject, true);
            }
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
        }

        private void DocumentEvents_DocumentOpened(Document document)
        {
            foreach (var plugin in Plugins)
            {
                plugin.CommandManager.ApplyDependencies(DependentItems.Document, true);
            }
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

    public static class EnumExt
    {
        /// <summary>
        ///     Check to see if a flags enumeration has a specific flag set.
        /// </summary>
        /// <param name="variable">Flags enumeration to check</param>
        /// <param name="value">Flag to check for</param>
        /// <returns></returns>
        public static bool HasFlag(this Enum variable, Enum value)
        {
            if (variable == null)
                return false;

            if (value == null)
                throw new ArgumentNullException("value");

            // Not as good as the .NET 4 version of this function, but should be good enough
            if (!Enum.IsDefined(variable.GetType(), value))
            {
                throw new ArgumentException(string.Format(
                    "Enumeration type mismatch.  The flag is of type '{0}', was expecting '{1}'.",
                    value.GetType(), variable.GetType()));
            }

            var num = Convert.ToUInt64(value);
            return ((Convert.ToUInt64(variable) & num) == num);
        }
    }
}