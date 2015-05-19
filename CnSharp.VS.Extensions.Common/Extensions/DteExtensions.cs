using System;
using System.Collections.Generic;
using EnvDTE;
using EnvDTE80;

namespace CnSharp.VisualStudio.Extensions
{
    public static class DteExtensions
    {
        public static Project GetStartupProject(this DTE2 dte)
        {
            var projects = (Array)dte.Solution.SolutionBuild.StartupProjects;
            if (projects != null && projects.Length >= 1)
            {
                return projects.GetValue(0) as Project;
            }
            return null;
        }

        public static IEnumerable<Project> GetSolutionProjects(this DTE2 dte)
        {
            var projects =  dte.Solution.Projects;
            foreach (var project in projects)
            {
                var p = project as Project;
                var fileName = string.Empty;
                try
                {
                    fileName = p.FileName; //some kind of projects get FileName throw a exception
                }
                catch
                {
                    
                }
               
                if (p != null && !string.IsNullOrEmpty(fileName))
                    yield return p;
            }
        }

        public static Project GetActiveProejct(this DTE2 dte)
        {
            var projects = (Array)dte.ActiveSolutionProjects;
            if (projects != null && projects.Length >= 1)
            {
                return projects.GetValue(0) as Project;
            }
            return null;
        }


        public static void OutputMessage(this DTE2 dte,string commandName, string message)
        {
            //get output window
            OutputWindow ow = dte.ToolWindows.OutputWindow;
            //create own pane type
            OutputWindowPane outputPane = null;
            foreach (OutputWindowPane pane in ow.OutputWindowPanes)
            {
                if (pane.Name == commandName)
                {
                    outputPane = pane;
                    break;
                }
            }
            if (outputPane == null)
                outputPane = ow.OutputWindowPanes.Add(commandName);
            //output message
            outputPane.OutputString(message);
            outputPane.Activate();
        }
    }
}