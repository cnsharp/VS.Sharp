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
            string fileVersion =
                Regex.Match(assemblyInfo, @"AssemblyFileVersion\(""(?<content>.+)""\)").Groups["content"].Value;
            string version = Regex.Match(assemblyInfo, @"AssemblyVersion\(""(?<content>.+)""\)").Groups["content"].Value;
            string productName =
                Regex.Match(assemblyInfo, @"AssemblyProduct\(""(?<content>.+)""\)").Groups["content"].Value;
            string companyName =
                Regex.Match(assemblyInfo, @"AssemblyCompany\(""(?<content>.+)""\)").Groups["content"].Value;
            string title =
               Regex.Match(assemblyInfo, @"AssemblyTitle\(""(?<content>.+)""\)").Groups["content"].Value;
            string description =
              Regex.Match(assemblyInfo, @"AssemblyDescription\(""(?<content>.+)""\)").Groups["content"].Value;
            string copyright =
             Regex.Match(assemblyInfo, @"AssemblyCopyright\(""(?<content>.+)""\)").Groups["content"].Value;
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
           
            using (var sw = new StreamWriter(file,false, Encoding.Default))
            {
              
                assemblyText = Regex.Replace(assemblyText, @"AssemblyFileVersion\("".+""\)",
                              string.Format("AssemblyFileVersion(\"{0}\")", assemblyInfo.FileVersion));
                assemblyText = Regex.Replace(assemblyText, @"AssemblyVersion\("".+""\)",
                              string.Format("AssemblyVersion(\"{0}\")", assemblyInfo.Version));
                assemblyText = Regex.Replace(assemblyText, @"AssemblyProduct\("".+""\)",
                              string.Format("AssemblyProduct(\"{0}\")", assemblyInfo.ProductName));
                assemblyText = Regex.Replace(assemblyText, @"AssemblyCompany\("".+""\)",
                              string.Format("AssemblyCompany(\"{0}\")", assemblyInfo.Company));
                assemblyText = Regex.Replace(assemblyText, @"AssemblyTitle\("".+""\)",
                              string.Format("AssemblyTitle(\"{0}\")", assemblyInfo.Title));
                assemblyText = Regex.Replace(assemblyText, @"AssemblyDescription\("".+""\)",
                              string.Format("AssemblyDescription(\"{0}\")", assemblyInfo.Description));
                assemblyText = Regex.Replace(assemblyText, @"AssemblyCopyright\("".+""\)",
                             string.Format("AssemblyCopyright(\"{0}\")", assemblyInfo.Copyright));
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