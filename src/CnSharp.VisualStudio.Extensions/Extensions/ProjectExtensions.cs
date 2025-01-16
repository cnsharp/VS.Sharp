using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using CnSharp.VisualStudio.Extensions.Projects;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSLangProj;
using VSLangProj80;
using Constants = EnvDTE.Constants;

namespace CnSharp.VisualStudio.Extensions
{
    /// <summary>
    ///     extensions of <see cref="Project" />
    /// </summary>
    /// <remarks>
    ///     http://www.codeproject.com/Articles/36219/Exploring-EnvDTE
    /// </remarks>
    public static class ProjectExtensions
    {
        /// <summary>
        ///     get root namespace of project
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static string GetRootNameSpace(this Project project)
        {
            return project.Properties.Item("RootNamespace").Value.ToString();
        }

        /// <summary>
        ///     get project directory full path
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static string GetDirectory(this Project project)
        {
            return project.Properties.Item("FullPath").Value.ToString();
        }

        /// <summary>
        ///     get reference projects of project
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static IEnumerable<Project> GetReferenceProjects(this Project project)
        {
            var vsProject = (VSProject2)project.Object;
            return from Reference3 r in vsProject.References where r.SourceProject != null select r.SourceProject;
        }


        /// <summary>
        /// </summary>
        /// <param name="project"></param>
        /// <param name="referenceIdentity"></param>
        /// <param name="browseUrl">
        ///     is either the File Path or the Strong Name
        ///     e.g (System.Configuration, Version=2.0.0.0, Culture=neutral, PublicKeyToken=B03F5F7F11D50A3A)
        /// </param>
        public static void AddReference(this Project project, string referenceIdentity, string browseUrl)
        {
            var path = "";

            if (!browseUrl.StartsWith(referenceIdentity))
                //it is a path
                path = browseUrl;


            var vsProject = project.Object as VSProject;
            if (vsProject != null)
            {
                Reference reference = null;
                try
                {
                    reference = vsProject.References.Find(referenceIdentity);
                }
                catch (Exception ex)
                {
                    //it failed to find one, so it must not exist. 
                    //But it decided to error for the fun of it. :)
                }

                if (reference == null)
                {
                    if (path == "")
                        vsProject.References.Add(browseUrl);
                    else
                        vsProject.References.Add(path);
                }
                else
                {
                    throw new Exception("Reference already exists.");
                }
            }
        }


        /// <summary>
        /// </summary>
        /// <param name="project"></param>
        /// <param name="referenceIdentity"></param>
        /// <param name="browseUrl">
        ///     is either the File Path or the Strong Name
        ///     e.g (System.Configuration, Version=2.0.0.0, Culture=neutral, PublicKeyToken=B03F5F7F11D50A3A)
        /// </param>
        public static void RemoveReference(this Project project, string referenceIdentity)
        {
            var vsProject = project.Object as VSProject;
            if (vsProject != null)
            {
                Reference reference = null;
                try
                {
                    reference = vsProject.References.Find(referenceIdentity);
                }
                catch (Exception ex)
                {
                }

                if (reference != null) reference.Remove();
            }
        }


        public static List<KeyValuePair<string, string>> GetReferences(this Project project)
        {
            var list =
                new List<KeyValuePair<string, string>>();
            var vsproject = project.Object as VSProject;
            if (vsproject != null)
            {
                foreach (Reference reference in vsproject.References)
                    if (reference.StrongName)
                        //System.Configuration, Version=2.0.0.0,
                        //Culture=neutral, PublicKeyToken=B03F5F7F11D50A3A
                        list.Add(new KeyValuePair<string, string>(reference.Identity,
                            string.Format("{0}, Version={1}, Culture={2}, PublicKeyToken={3}", reference.Identity,
                                reference.Version,
                                string.IsNullOrEmpty(reference.Culture) ? "neutral" : reference.Culture,
                                reference.PublicKeyToken)));
                    else
                        list.Add(new KeyValuePair<string, string>(
                            reference.Identity, reference.Path));
                return list;
            }

            throw new NotSupportedException("Currently, system is only set up to do references for normal projects.");
        }


        public static void AddFromFile(this Project project, List<string> paths, string file)
        {
            var projectItems = project.ProjectItems;
            foreach (var p in paths) projectItems = projectItems.Item(p).ProjectItems;
            projectItems.AddFromFile(file);
        }


        public static void AddFromFileAsLink(this Project project, List<string> paths, string file, string linkName)
        {
            var projectItems = project.ProjectItems;
            foreach (var p in paths) projectItems = projectItems.Item(p).ProjectItems;
            var npi = projectItems.AddFromFile(file);
            project.SetItemAttribute(npi, "Link", linkName);
        }

        public static void SetItemAttribute(this Project project, List<string> paths, string file, string key,
            string value)
        {
            var projectItems = project.ProjectItems;
            foreach (var p in paths) projectItems = projectItems.Item(p).ProjectItems;
            project.SetItemAttribute(projectItems.Item(file), key, value);
        }

        public static void SetItemAttribute(this Project project, ProjectItem projectItem, string key, string value)
        {
            var uniqueName = project.UniqueName;
            var solution = (IVsSolution)Package.GetGlobalService(typeof(SVsSolution));
            IVsHierarchy hierarchy;
            solution.GetProjectOfUniqueName(uniqueName, out hierarchy);
            var buildPropertyStorage = hierarchy as IVsBuildPropertyStorage;
            var fullPath = (string)projectItem.Properties.Item("FullPath").Value;
            uint itemId;
            hierarchy.ParseCanonicalName(fullPath, out itemId);
            buildPropertyStorage.SetItemAttribute(itemId, key, value);
        }


        public static void AddFolder(this Project project, List<string> paths, string newFolder)
        {
            var projectItems = project.ProjectItems;
            foreach (var p in paths) projectItems = projectItems.Item(p).ProjectItems;
            projectItems.AddFolder(newFolder);
        }

        //path is a list of folders from the root of the project.
        public static void DeleteFileOrFolder(this Project project, List<string> paths, string item)
        {
            var projectItems = project.ProjectItems;
            foreach (var p in paths) projectItems = projectItems.Item(p).ProjectItems;
            projectItems.Item(item).Delete();
        }

        public static string GetPropertyValue(this Project project, string key)
        {
            try
            {
                return project.Properties.Item(key).Value.ToString();
            }
            catch
            {
                return null;
            }
        }

        public static bool IsNetCoreProject(this Project project)
        {
            return project.GetPropertyValue("TargetFrameworkMoniker").StartsWith(".NETCore");
        }

        public static bool IsNetStandardProject(this Project project)
        {
            return project.GetPropertyValue("TargetFrameworkMoniker").StartsWith(".NETStandard");
        }

        public static bool IsNetFrameworkProject(this Project project)
        {
            return project.GetPropertyValue("TargetFrameworkMoniker").StartsWith(".NETFramework");
        }

        public static bool IsSdkBased(this Project project)
        {
            var doc = new XmlDocument();
            doc.Load(project.FileName);
            return doc.DocumentElement.HasAttribute("Sdk");
        }

        public static string GetFileName(this Project project)
        {
            return project.GetPropertyValue("FileName");
        }

        public static string GetCodeFileExtension(this Project project)
        {
            return Path.GetExtension(project.GetFileName()).Replace("proj", string.Empty);
        }


        public static void SaveCommonAssemblyInfo(this Project project, CommonAssemblyInfo assemblyInfo)
        {
            var file = Path.Combine(Path.GetDirectoryName(project.DTE.Solution.FileName),
                typeof(CommonAssemblyInfo) + project.GetCodeFileExtension());
            var manager = AssemblyInfoFileManagerFactory.Get(project);
            manager.Save(assemblyInfo, file);
        }


        public static ProjectItem FindProjectItem(this Project project, string name, bool recursive)
        {
            ProjectItem projectItem = null;

            if (project.Kind != Constants.vsProjectKindSolutionItems)
            {
                if (project.ProjectItems != null && project.ProjectItems.Count > 0)
                    projectItem = FindItemByName(project.ProjectItems, name, recursive);
            }
            else
            {
                // if solution folder, one of its ProjectItems might be a real project
                foreach (ProjectItem item in project.ProjectItems)
                {
                    var realProject = item.Object as Project;

                    if (realProject != null)
                    {
                        projectItem = FindProjectItem(realProject, name, recursive);

                        if (projectItem != null) break;
                    }
                }
            }

            return projectItem;
        }

        public static ProjectItem FindItemByName(this ProjectItems items, string name, bool recursive)
        {
            foreach (ProjectItem item in items)
            {
                if (item.Name == name)
                    return item;
                if (recursive)
                    return FindItemByName(item.ProjectItems, name, true);
            }

            return null;
        }

        public static IEnumerable<ProjectItem> FindItemByName(this ProjectItems items, Func<string, bool> matchFunc)
        {
            return items.Cast<ProjectItem>().Where(item => matchFunc(item.Name));
        }

        public static string GetCommonAssemblyInfoFilePath(this Project project)
        {
            var manager = AssemblyInfoFileManagerFactory.Get(project);
            var folderName = manager.FolderName;
            var folder = Path.Combine(project.GetDirectory(), folderName);
            if (!Directory.Exists(folder)) return null;
            var folderItem = project.ProjectItems.FindItemByName(folderName, false);
            var sub =
                folderItem.ProjectItems.FindItemByName(n => Regex.IsMatch(n, ".+AssemblyInfo.+"))?.FirstOrDefault();
            return sub?.FileNames[0];
        }

        public static void LinkCommonAssemblyInfoFile(this Project project, string file)
        {
            var manager = AssemblyInfoFileManagerFactory.Get(project);
            var folderName = manager.FolderName;
            var folder = Path.Combine(project.GetDirectory(), folderName);
            ProjectItem folderItem;
            if (!Directory.Exists(folder))
                folderItem = project.ProjectItems.AddFolder(folderName);
            else
                folderItem = project.ProjectItems.FindItemByName(folderName, false);
            var fileName = Path.GetFileName(file);
            var linkItem = folderItem.ProjectItems.FindItemByName(fileName, false);
            linkItem?.Delete();
            folderItem.ProjectItems.AddFromFile(file);
        }

        public static PackageProjectProperties GetPackageProjectProperties(this Project project)
        {
            var ppp = new PackageProjectProperties();
            var properties = typeof(PackageProjectProperties).GetProperties().ToList();
            var doc = new XmlDocument();
            doc.Load(project.FileName);

            foreach (var p in properties)
                try
                {
                    var node = doc.SelectSingleNode("/Project/PropertyGroup/" + p.Name);
                    if (node != null)
                    {
                        if (p.PropertyType == typeof(bool))
                            p.SetValue(ppp, Convert.ToBoolean(node.InnerText), null);
                        else
                            p.SetValue(ppp, node.InnerText, null);
                    }
                }
                catch (Exception ex)
                {
                    // ignored
                }

            return ppp;
        }

        public static void SavePackageProjectProperties(this Project project, PackageProjectProperties ppp,
            params string[] skipProperties)
        {
            var properties = typeof(PackageProjectProperties).GetProperties().ToList();
            if (skipProperties != null)
                properties = properties.Where(p => !skipProperties.Contains(p.Name)).ToList();

            var doc = new XmlDocument();
            doc.Load(project.FileName);
            foreach (var p in properties)
                try
                {
                    var v = p.GetValue(ppp, null);
                    var val = p.PropertyType == typeof(bool)
                        ? v != null && Convert.ToBoolean(v)
                        : v ?? string.Empty;
                    var node = doc.SelectSingleNode("/Project/PropertyGroup/" + p.Name);
                    if (node != null)
                    {
                        node.InnerText = val.ToString();
                    }
                    else
                    {
                        var propertyGroup = doc.SelectSingleNode("/Project/PropertyGroup");
                        if (propertyGroup != null)
                        {
                            var newNode = doc.CreateElement(p.Name);
                            newNode.InnerText = val.ToString();
                            propertyGroup.AppendChild(newNode);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // ignored
                }

            doc.Save(project.FileName);
        }
    }
}