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
                throw new FileNotFoundException("Assembly info file 'AssemblyInfo.cs' not found in this project.");
            assemblyInfo = Regex.Replace(assemblyInfo, "//.*", "");
            string fileVersion = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyFileVersion");
             
            string version = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyVersion"); 
            string productName = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyProduct");
            string companyName = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyCompany");
            string title = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyTitle");
            string description = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyDescription");
            string copyright = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyCopyright");
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

        private static string GetAssemblyAnnotationValue(string assemblyInfo, string attributeName)
        {
            return
                Regex.Match(assemblyInfo, $"[^/]\\[assembly:\\s*?{attributeName}\\(\"(?<content>.+)\"\\)").Groups["content"].Value;
        }

        public static void ModifyAssemblyInfo(this Project project,ProjectAssemblyInfo assemblyInfo)
        {
            var file = project.GetAssemblyInfoFileName();

            var assemblyText = project.GetAssemblyInfo();

            var sc = Host.Instance.SourceControl;
            sc?.CheckOut(Path.GetDirectoryName(Host.Instance.DTE.Solution.FullName), file);

            using (var sw = new StreamWriter(file,false, Encoding.Unicode))
            {
                assemblyText = ReplaceAssemblyAnnotation(assemblyText, "AssemblyFileVersion", assemblyInfo.FileVersion);
                assemblyText = ReplaceAssemblyAnnotation(assemblyText, "AssemblyVersion", assemblyInfo.Version);
                assemblyText = ReplaceAssemblyAnnotation(assemblyText, "AssemblyProduct", assemblyInfo.ProductName);
                assemblyText = ReplaceAssemblyAnnotation(assemblyText, "AssemblyCompany", assemblyInfo.Company);
                assemblyText = ReplaceAssemblyAnnotation(assemblyText, "AssemblyTitle", assemblyInfo.Title);
                assemblyText = ReplaceAssemblyAnnotation(assemblyText, "AssemblyDescription", assemblyInfo.Description);
                assemblyText = ReplaceAssemblyAnnotation(assemblyText, "AssemblyCopyright", assemblyInfo.Copyright);
                sw.Write(assemblyText);
            }
        }

        private static string ReplaceAssemblyAnnotation(string assemblyText, string attributeName, string value)
        {
            var text = Regex.Replace(assemblyText, $"[^/]\\[assembly:\\s*?{attributeName}\\(\".*?\"\\)\\]",
                $"[assembly: {attributeName}(\"{value}\")]");
            return text.Replace("\r[", "\r\n[");//这里有个坑
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