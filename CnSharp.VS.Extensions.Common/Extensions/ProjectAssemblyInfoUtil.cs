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
            string assemblyInfoFile = GetAssemblyInfoFileName(project);
            var manager = AssemblyInfoFileManagerFactory.Get(assemblyInfoFile);
            var info = manager.Read(assemblyInfoFile);
            info.Project = project;
            return info;
        }

   

        public static void ModifyAssemblyInfo(this Project project,ProjectAssemblyInfo assemblyInfo)
        {
            var assemblyInfoFile = project.GetAssemblyInfoFileName();
            var manager = AssemblyInfoFileManagerFactory.Get(assemblyInfoFile);
            manager.Save(assemblyInfo,assemblyInfoFile);
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
                assemblyInfoFile = prjDir + "\\My Project\\AssemblyInfo.vb";
            }
            if (!File.Exists(assemblyInfoFile))
            {
                assemblyInfoFile = prjDir + "\\AssemblyInfo.vb";
            }
            if (!File.Exists(assemblyInfoFile))
            {
                throw new FileNotFoundException("AssemblyInfo file not found in this project.");
            }
            return assemblyInfoFile;
        }

    }


    public class AssemblyInfoFileManagerFactory
    {
        public static AssemblyInfoFileManager Get(string file)
        {
            var ext = Path.GetExtension(file).TrimStart('.').ToLower();
            return ext == "cs" ? (AssemblyInfoFileManager) new AssemblyInfoCsManager() : new AssemblyInfoVbManager();
        }
    }

    public abstract class AssemblyInfoFileManager
    {
        protected string ReadRegexPattern;
        protected string WriteRegexPattern;

        public  ProjectAssemblyInfo Read(string file)
        {
            var assemblyInfo = ReadFile(file);
            if (string.IsNullOrEmpty(assemblyInfo))
                throw new FileLoadException("AssemblyInfo file content is empty.");
            assemblyInfo = Regex.Replace(assemblyInfo, "['|//].*", "");
            string fileVersion = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyFileVersion");

            string version = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyVersion");
            string productName = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyProduct");
            string companyName = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyCompany");
            string title = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyTitle");
            string description = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyDescription");
            string copyright = GetAssemblyAnnotationValue(assemblyInfo, "AssemblyCopyright");
            return new ProjectAssemblyInfo
            {
                FileVersion = fileVersion,
                Version = version,
                ProductName = productName,
                Company = companyName,
                Title = title,
                Description = description,
                Copyright = copyright
            };
        }


        protected string ReadFile(string file)
        {
            using (var sr = new StreamReader(file, Encoding.Default))
            {
                return sr.ReadToEnd();
            }
        }


        protected  string GetAssemblyAnnotationValue(string assemblyInfo, string attributeName)
        {
            return
                Regex.Match(assemblyInfo, string.Format(ReadRegexPattern, attributeName)).Groups["content"].Value;
        }

        public virtual void Save(ProjectAssemblyInfo assemblyInfo, string file)
        {
            var assemblyText = ReadFile(file);

            var sc = Host.Instance.SourceControl;
            sc?.CheckOut(Path.GetDirectoryName(Host.Instance.DTE.Solution.FullName), file);

            using (var sw = new StreamWriter(file, false, Encoding.Unicode))
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

       protected string ReplaceAssemblyAnnotation(string assemblyText, string attributeName, string value)
        {
            //var text = Regex.Replace(assemblyText, $"[^/]\\[assembly:\\s*?{attributeName}\\(\".*?\"\\)\\]",
            //    $"[assembly: {attributeName}(\"{value}\")]\n");
            //return text.Replace("\r[", "\r\n[");//这里有个坑

            var text = assemblyText;
            var m = Regex.Match(text, string.Format(WriteRegexPattern,attributeName));
            if (m.Success)
            {
                var newValue = Regex.Replace(m.Value, "\\(\".*?\"\\)", $"(\"{value}\")");
                text = text.Replace(m.Value, newValue);
            }
            return text;
        }
    }

    public class AssemblyInfoCsManager : AssemblyInfoFileManager
    {
        public AssemblyInfoCsManager()
        {
            ReadRegexPattern = "[^/]\\[assembly:\\s*?{0}\\(\"(?<content>.+)\"\\)";
            WriteRegexPattern = "[^/]\\[assembly:\\s*?{0}\\(\".*?\"\\)\\]";
        }
    }

    public class AssemblyInfoVbManager : AssemblyInfoFileManager
    {
        public AssemblyInfoVbManager()
        {
            ReadRegexPattern = "[^`]\\<Assembly:\\s*?{0}\\(\"(?<content>.+)\"\\)";
            WriteRegexPattern = "[^`]\\<Assembly:\\s*?{0}\\(\".*?\"\\)\\>";
        }
    }
}