using EnvDTE;
using EnvDTE80;

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
                if (projectItem.FileNames[0] == file) return projectItem;
            }
            return folder.ProjectItems.AddFromFile(file);
        }
    }
}
