﻿using System;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;

namespace CnSharp.VisualStudio.Extensions
{
    public static class SolutionExtensions
    {
        public static ProjectItem FindSolutionItemByName(this Solution sln, string name, bool recursive)
        {
            ProjectItem projectItem = null;
            foreach (Project project in sln.Projects)
            {
                projectItem = project.FindProjectItem(name, recursive);

                if (projectItem != null)
                {
                    break;
                }
            }
            return projectItem;
        }

        public const string SolutionItemsFolderName = "Solution Items";

        public static Project GetOrAddSolutionItemsFolder(this Solution2 sln)
        {
            foreach (Project project in sln.Projects)
            {
                if(project.Name == SolutionItemsFolderName)
                {
                    return project;
                }
            }
            return sln.AddSolutionFolder(SolutionItemsFolderName);
        }

        public static ProjectItem AddSolutionItem(this Solution2 sln,string file)
        {
            var folder = sln.GetOrAddSolutionItemsFolder();
            foreach (ProjectItem projectItem in folder.ProjectItems)
            {
                var projItemGuid = new Guid(projectItem.Kind);
                var fileName = projItemGuid == VSConstants.GUID_ItemType_PhysicalFile
                    ? projectItem.FileNames[0]
                    : projectItem.FileNames[1];
                if (fileName == file) return projectItem;
            }
            return folder.ProjectItems.AddFromFile(file);
        }

        public static bool Build(this Solution2 sln, string configName = null)
        {
            var solutionBuild = (SolutionBuild2)sln.SolutionBuild;
            if (!string.IsNullOrWhiteSpace(configName))
            {
                solutionBuild.SolutionConfigurations.Item(configName).Activate();
            }
            solutionBuild.Build(true);
            return solutionBuild.LastBuildInfo == 0;
        }

        public static bool BuildRelease(this Solution2 sln)
        {
            return Build(sln, "Release");
        }

        public static bool BuildDebug(this Solution2 sln)
        {
            return Build(sln, "Debug");
        }
    }
}