using System.IO;
using System.Linq;
using System.Xml.Linq;
using CnSharp.VisualStudio.Extensions.Projects;
using EnvDTE;
using NuGet;

namespace CnSharp.VisualStudio.Extensions
{
    public static class NuGetExtensions
    {
        public static ManifestMetadata GetManifestMetadata(this Project project)
        {
            var metadata = new ManifestMetadata();
            if (project.IsNetFrameworkProject())
            {
                var ai = project.GetProjectAssemblyInfo();
                //from nuspec
                var nuspecFile = Path.Combine(project.GetDirectory(), NuGetDomain.NuSpecFileName);
                if (!File.Exists(nuspecFile))
                    return ai.ToManifestMetadata();
                return metadata.LoadFromNuspecFile(nuspecFile).CopyFromAssemblyInfo(ai);
            }
            var props = typeof(ManifestMetadata).GetProperties();
            foreach (var prop in props)
            {
                var key = prop.Name;
                var v = project.GetPropertyValue(key) ?? project.GetPropertyValue("Package"+key);
                if (v != null)
                    prop.SetValue(metadata,v,null);
            }
            return metadata;
        }

        public static ManifestMetadata LoadFromNuspecFile(this ManifestMetadata metadata,string file)
        {
            var doc = new XDocument(file);
            var elements = doc.Element("package").Element("metadata").Elements();
            var props = typeof(ManifestMetadata).GetProperties();
            foreach (var prop in props)
            {
                var key = prop.Name.ToLower();
                if (prop.PropertyType == typeof(string))
                {
                    var v = elements.FirstOrDefault(m => m.Name == key)?.Value;
                    if (v != null)
                        prop.SetValue(metadata, v, null);
                }
            }
           
            return metadata;
        }

        public static bool IsEmptyOrPlaceHolder(this string value)
        {
            return string.IsNullOrWhiteSpace(value) || value.StartsWith("$");
        }

        public static ManifestMetadata ToManifestMetadata(this ProjectAssemblyInfo ai)
        {
            var metadata = new ManifestMetadata
            {
                Id = ai.Title,
                Owners = ai.Company,
                Title = ai.Title,
                Description = ai.Description,
                Authors = ai.Company,
                Copyright = ai.Copyright
            };
            return metadata;
        }

        public static ManifestMetadata CopyFromAssemblyInfo(this ManifestMetadata metadata,ProjectAssemblyInfo ai)
        {
            if (metadata.Id.IsEmptyOrPlaceHolder())
                metadata.Id = ai.Title;
            if (metadata.Title.IsEmptyOrPlaceHolder())
                metadata.Title = ai.Title;
            if (metadata.Owners.IsEmptyOrPlaceHolder())
                metadata.Owners = ai.Company;
            if (metadata.Description.IsEmptyOrPlaceHolder())
                metadata.Description = ai.Description;
            if (metadata.Authors.IsEmptyOrPlaceHolder())
                metadata.Authors = ai.Company;
            if (metadata.Copyright.IsEmptyOrPlaceHolder())
                metadata.Copyright = ai.Copyright;
            return metadata;
        }
    }

    public class NuGetDomain
    {
        public const string NuSpecFileName = "package.nuspec";
    }

}
