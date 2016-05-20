using System.IO;
using System.Reflection;
using EnvDTE;

namespace CnSharp.VisualStudio.Extensions.SourceControl
{
    public class SourceControlManager
    {
        private static readonly string[] SupportTypes = new[]
        {
            "CnSharp.VisualStudio.SourceControl.Tfs.TfsSourceControl,CnSharp.VisualStudio.SourceControl.Tfs"
        };
        public static ISourceControl GetSolutionSourceControl(Solution solution)
        {
            if (solution == null)
                return null;
            var fileName = solution.FileName;
            var scFileName = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName) + ".vssscc");
            if (File.Exists(scFileName))
                return LoadInstance(SupportTypes[0]) as ISourceControl;
            return null;
        }

        private static  object LoadInstance(string typeName)
        {
            if (typeName.Contains(","))
            {
                var arr = typeName.Split(',');
                if (arr.Length < 2)
                    return null;
                var assemblyName = arr[1];
                var assembly = Assembly.Load(assemblyName);
                return assembly.CreateInstance(arr[0]);
            }
            return typeof(SourceControlManager).Assembly.CreateInstance(typeName);
        }
    }
}
