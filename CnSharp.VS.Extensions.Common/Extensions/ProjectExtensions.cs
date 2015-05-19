using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EnvDTE;
using VSLangProj;
using VSLangProj80;
using VsWebSite;

namespace CnSharp.VisualStudio.Extensions
{
    /// <summary>
    ///     extensions of <see cref="Project"/>
    /// </summary>
    /// <remarks>
    ///     http://www.codeproject.com/Articles/36219/Exploring-EnvDTE
    /// </remarks>
    public static class ProjectExtensions
    {
        /// <summary>
        /// get root namespace of project
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static string GetRootNameSpace(this Project project)
        {
            return project.Properties.Item("RootNamespace").Value.ToString();
        }

        /// <summary>
        /// get project directory full path
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static string GetDirectory(this Project project)
        {
            return project.Properties.Item("FullPath").Value.ToString();
        }

        /// <summary>
        /// get reference projects of project
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static IEnumerable<Project> GetReferenceProjects(this Project project)
        {
            var vsProject = ((VSProject2) (project.Object));
            return from Reference3 r in vsProject.References where r.SourceProject != null select (r.SourceProject);
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
            string path = "";

            if (!browseUrl.StartsWith(referenceIdentity))
            {
                //it is a path
                path = browseUrl;
            }


            VSProject vsProject = project.Object as VSProject;
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
            else
            {
                VSWebSite vsWebSite = project.Object as VSWebSite;
                if (vsWebSite != null)
                {
                    AssemblyReference reference = null;
                    try
                    {
                        foreach (AssemblyReference r in vsWebSite.References)
                        {
                            if (r.Name == referenceIdentity)
                            {
                                reference = r;
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        //it failed to find one, so it must not exist. 
                        //But it decided to error for the fun of it. :)
                    }
                    if (reference == null)
                    {
                        if (path == "")
                            vsWebSite.References.AddFromGAC(browseUrl);
                        else
                            vsWebSite.References.AddFromFile(path);
                    }
                    else
                    {
                        throw new Exception("Reference already exists.");
                    }
                }
                else
                {
                    throw new Exception("Currently, system is only set up to do references for normal projects.");
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


            VSProject vsProject = project.Object as VSProject;
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
                if (reference != null)
                {
                        reference.Remove();
                }
            }
            else
            {
                VSWebSite vsWebSite = project.Object as VSWebSite;
                if (vsWebSite != null)
                {
                    AssemblyReference reference = null;
                    try
                    {
                        foreach (AssemblyReference r in vsWebSite.References)
                        {
                            if (r.Name == referenceIdentity)
                            {
                                reference = r;
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        //it failed to find one, so it must not exist. 
                        //But it decided to error for the fun of it. :)
                    }
                    if (reference != null)
                    {
                       reference.Remove();
                    }
                }
                else
                {
                    throw new Exception("Currently, system is only set up  to do references for normal projects.");
                }
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
                {
                    if (reference.StrongName)
                        //System.Configuration, Version=2.0.0.0,
                        //Culture=neutral, PublicKeyToken=B03F5F7F11D50A3A
                        list.Add(new KeyValuePair<string, string>(reference.Identity,
                            string.Format("{0}, Version={1}, Culture={2}, PublicKeyToken={3}", reference.Identity,
                                reference.Version,
                                (string.IsNullOrEmpty(reference.Culture) ? "neutral" : reference.Culture),
                                reference.PublicKeyToken)));
                    else
                        list.Add(new KeyValuePair<string, string>(
                            reference.Identity, reference.Path));
                }
                return list;
            }
            var vswebsite = project.Object as VSWebSite;
            if (vswebsite != null)
            {
                foreach (AssemblyReference reference in vswebsite.References)
                {
                    string value = "";
                    if (reference.FullPath != "")
                    {
                        var f = new FileInfo(reference.FullPath + ".refresh");
                        if (f.Exists)
                        {
                            using (FileStream stream = f.OpenRead())
                            {
                                using (var r = new StreamReader(stream))
                                {
                                    value = r.ReadToEnd().Trim();
                                }
                            }
                        }
                    }
                    if (value == "")
                    {
                        list.Add(new KeyValuePair<string, string>(reference.Name,
                            reference.StrongName));
                    }
                    else
                    {
                        list.Add(new KeyValuePair<string, string>(reference.Name, value));
                    }
                }
                return list;
            }
            throw new NotSupportedException("Currently, system is only set up to do references for normal projects.");
        }

        public static void AddFromFile(this Project project, List<string> path, string file)
        {
            ProjectItems pi = project.ProjectItems;
            for (int i = 0; i < path.Count; i++)
            {
                pi = pi.Item(path[i]).ProjectItems;
            }
            pi.AddFromFile(file);
        }


        public static void AddFolder(this Project project, string newFolder, List<string> path)
        {
            ProjectItems pi = project.ProjectItems;
            for (int i = 0; i < path.Count; i++)
            {
                pi = pi.Item(path[i]).ProjectItems;
            }
            pi.AddFolder(newFolder);
        }

        //path is a list of folders from the root of the project.
        public static void DeleteFileOrFolder(this Project project, List<string> path, string item)
        {
            ProjectItems pi = project.ProjectItems;
            for (int i = 0; i < path.Count; i++)
            {
                pi = pi.Item(path[i]).ProjectItems;
            }
            pi.Item(item).Delete();
        }
    }
}