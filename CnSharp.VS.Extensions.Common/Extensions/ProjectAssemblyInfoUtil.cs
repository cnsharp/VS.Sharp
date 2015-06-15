using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using CnSharp.VisualStudio.Extensions.Projects;
using EnvDTE;

namespace CnSharp.VisualStudio.Extensions
{
    public static class ProjectAssemblyInfoUtil
    {
        public static ProjectAssemblyInfo GetProjectAssemblyInfo(this Project project)
        {
            string assemblyInfo = GetAssemblyInfo(project);
            if (string.IsNullOrEmpty(assemblyInfo))
                throw new FileNotFoundException("Assembly info file of project '" + project.FullName + "' missed.");
            assemblyInfo = Regex.Replace(assemblyInfo, "//.*", "");
            string fileVersion =
                Regex.Match(assemblyInfo, @"[^/]\[assembly:\s*AssemblyFileVersion\(""(?<content>.+)""\)").Groups["content"].Value;
            string version = Regex.Match(assemblyInfo, @"[^/]\[assembly:\s*AssemblyVersion\(""(?<content>.+)""\)").Groups["content"].Value;
            string productName =
                Regex.Match(assemblyInfo, @"[^/]\[assembly:\s*AssemblyProduct\(""(?<content>.+)""\)").Groups["content"].Value;
            string companyName =
                Regex.Match(assemblyInfo, @"[^/]\[assembly:\s*AssemblyCompany\(""(?<content>.+)""\)").Groups["content"].Value;
            string title =
               Regex.Match(assemblyInfo, @"[^/]\[assembly:\s*AssemblyTitle\(""(?<content>.+)""\)").Groups["content"].Value;
            string description =
              Regex.Match(assemblyInfo, @"[^/]\[assembly:\s*AssemblyDescription\(""(?<content>.+)""\)").Groups["content"].Value;
            string copyright =
             Regex.Match(assemblyInfo, @"[^/]\[assembly:\s*AssemblyCopyright\(""(?<content>.+)""\)").Groups["content"].Value;
#if DEBUG
            if (version.Contains("*"))
                version = fileVersion;
#endif
            return new ProjectAssemblyInfo
            {
                Project = project,
                FileVersion = fileVersion,
                Version = version,
                ProductName = productName,
                Company = companyName,
                Title = title,
                Description = description,
                Copyright = copyright
            };
        }

        public static void ModifyAssemblyInfo(this Project project,ProjectAssemblyInfo assemblyInfo)
        {
            var file = project.GetAssemblyInfoFileName();

            var assemblyText = project.GetAssemblyInfo();

            var sc = Host.Instance.SourceControl;
            if (sc != null)
            {
               var ok =  sc.CheckOut(Path.GetDirectoryName(Host.Instance.DTE.Solution.FullName), file);
                if (ok < 0)
                {
                    throw new ApplicationException(string.Format("Check out file {0} failed,you may connect source control server at first.",file));
                }
            }

            using (var sw = new StreamWriter(file,false, Encoding.UTF8))
            {
               
                assemblyText = Regex.Replace(assemblyText, @"[^/]\[assembly:\s*AssemblyFileVersion\("".+""\)",
                              string.Format("\n[assembly: AssemblyFileVersion(\"{0}\")", assemblyInfo.FileVersion));
                assemblyText = Regex.Replace(assemblyText, @"\n\[assembly:\s*AssemblyVersion\("".+""\)",
                              string.Format("\n[assembly: AssemblyVersion(\"{0}\")", assemblyInfo.Version));
                assemblyText = Regex.Replace(assemblyText, @"[^/]\[assembly:\s*AssemblyProduct\("".+""\)",
                              string.Format("\n[assembly: AssemblyProduct(\"{0}\")", assemblyInfo.ProductName));
                assemblyText = Regex.Replace(assemblyText, @"[^/]\[assembly:\s*AssemblyCompany\("".*""\)",
                              string.Format("\n[assembly: AssemblyCompany(\"{0}\")", assemblyInfo.Company));
                assemblyText = Regex.Replace(assemblyText, @"[^/]\[assembly:\s*AssemblyTitle\("".+""\)",
                              string.Format("\n[assembly: AssemblyTitle(\"{0}\")", assemblyInfo.Title));
                assemblyText = Regex.Replace(assemblyText, @"[^/]\[assembly:\s*AssemblyDescription\("".+""\)",
                              string.Format("\n[assembly: AssemblyDescription(\"{0}\")", assemblyInfo.Description));
                assemblyText = Regex.Replace(assemblyText, @"[^/]\[assembly:\s*AssemblyCopyright\("".+""\)",
                             string.Format("\n[assembly: AssemblyCopyright(\"{0}\")", assemblyInfo.Copyright));
                sw.Write(assemblyText);
            }
        }

        public static string GetAssemblyInfoFileName(this Project project)
        {
            string prjDir = Path.GetDirectoryName(project.FileName);
            string assemblyInfoFile = prjDir + "\\Properties\\AssemblyInfo.cs";
            if (!File.Exists(assemblyInfoFile))
            {
                assemblyInfoFile = prjDir + "\\AssemblyInfo.cs";
            }
            if (!File.Exists(assemblyInfoFile))
            {
                return null;
            }
            return assemblyInfoFile;
        }

        public static string GetAssemblyInfo(this Project project)
        {
            string assemblyInfoFile = GetAssemblyInfoFileName(project);
            if (string.IsNullOrEmpty(assemblyInfoFile))
                return null;

            using (var sr = new StreamReader(assemblyInfoFile, Encoding.Default))
            {
                return sr.ReadToEnd();
            }
        }
    }
}