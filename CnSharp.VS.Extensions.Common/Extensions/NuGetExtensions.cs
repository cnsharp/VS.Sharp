using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using CnSharp.VisualStudio.Extensions.Projects;
using CnSharp.VisualStudio.Extensions.Util;
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
            metadata = project.GetPackageProjectProperties().ToManifestMetadata();
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

        public static void UpdateNuspec(this Project project, ManifestMetadata metadata)
        {
            var nuspecFile = Path.Combine(project.GetDirectory(), NuGetDomain.NuSpecFileName);
            if (!File.Exists(nuspecFile))
            {
                return;
            }
            else
            {
                var doc = new XmlDocument();
                doc.Load(nuspecFile);
                metadata.SyncToXmlDocument(doc);
            }
        }

        public static void SyncToXmlDocument(this  ManifestMetadata metadata, XmlDocument doc)
        {
            var metadataNode = doc.SelectSingleNode("package/metadata");
            if (metadataNode == null)
                return;
            metadataNode.SetXmlNode("id", metadata.Id);
            metadataNode.SetXmlNode("title", metadata.Title);
            metadataNode.SetXmlNode("description", metadata.Description);
            metadataNode.SetXmlNode("owners", metadata.Owners);
            metadataNode.SetXmlNode("authors", metadata.Authors);
            metadataNode.SetXmlNode("version", metadata.Version);
            metadataNode.SetXmlNode("releaseNotes", metadata.ReleaseNotes);
            UpdateDependencies( metadata,doc);
        }

        private static void SetXmlNode(this XmlNode metadataNode, string key, string value)
        {
            var idNode = metadataNode.SelectSingleNode(key);
            if (idNode != null)
                idNode.InnerText = value;
        }

        public static void UpdateDependencies(this ManifestMetadata metadata, XmlDocument doc)
        {
            var metadataNode = doc.SelectSingleNode("package/metadata");
            if (metadataNode == null)
                return;

            if (metadata.DependencySets != null)
            {
                var depNode = metadataNode.SelectSingleNode("dependencies");
                if (depNode == null)
                {
                    var node = doc.CreateElement("dependencies");
                    metadataNode.AppendChild(node);
                    depNode = node;
                }

                depNode.RemoveAll();
                var tempNode = doc.CreateElement("temp");
                tempNode.InnerXml = XmlSerializerHelper.GetXmlStringFromObject(metadata.DependencySets);
                depNode.InnerXml = tempNode.ChildNodes[0].InnerXml;
            }
        }



        public static bool IsEmptyOrPlaceHolder(this string value)
        {
            return string.IsNullOrWhiteSpace(value) || value.StartsWith("$");
        }

        public static ManifestMetadata ToManifestMetadata(this ProjectAssemblyInfo pai)
        {
            var metadata = new ManifestMetadata
            {
                Id = pai.Title,
                Owners = pai.Company,
                Title = pai.Title,
                Description = pai.Description,
                Authors = pai.Company,
                Copyright = pai.Copyright
            };
            return metadata;
        }

        public static ManifestMetadata CopyFromAssemblyInfo(this ManifestMetadata metadata,ProjectAssemblyInfo pai)
        {
            if (metadata.Id.IsEmptyOrPlaceHolder())
                metadata.Id = pai.Title;
            if (metadata.Title.IsEmptyOrPlaceHolder())
                metadata.Title = pai.Title;
            if (metadata.Owners.IsEmptyOrPlaceHolder())
                metadata.Owners = pai.Company;
            if (metadata.Description.IsEmptyOrPlaceHolder())
                metadata.Description = pai.Description;
            if (metadata.Authors.IsEmptyOrPlaceHolder())
                metadata.Authors = pai.Company;
            if (metadata.Copyright.IsEmptyOrPlaceHolder())
                metadata.Copyright = pai.Copyright;
            return metadata;
        }

        public static ManifestMetadata ToManifestMetadata(this PackageProjectProperties ppp)
        {
            return new ManifestMetadata
            {
                Id = ppp.PackageId,
                Authors = ppp.Authors,
                Copyright = ppp.Copyright,
                Owners = ppp.Company,
                Description = ppp.Description,
                IconUrl = ppp.PackageIconUrl,
                Language = ppp.NeutralLanguage,
                LicenseUrl = ppp.PackageLicenseUrl,
                ReleaseNotes = ppp.PackageReleaseNotes,
                RequireLicenseAcceptance = ppp.PackageRequireLicenseAcceptance,
                ProjectUrl = ppp.PackageProjectUrl,
                Tags = ppp.PackageTags,
                Version = ppp.Version
            };
        }

        public static void SyncToPackageProjectProperties(this ManifestMetadata metadata, PackageProjectProperties ppp)
        {
            ppp.PackageId = metadata.Id;
            ppp.Authors = metadata.Authors;
            ppp.Copyright = metadata.Copyright;
            ppp.Company = metadata.Owners;
            ppp.Description = metadata.Description;
            ppp.PackageIconUrl = metadata.IconUrl;
            ppp.NeutralLanguage = metadata.Language;
            ppp.PackageLicenseUrl = metadata.LicenseUrl;
            ppp.PackageReleaseNotes = metadata.ReleaseNotes;
            ppp.PackageRequireLicenseAcceptance = metadata.RequireLicenseAcceptance;
            ppp.PackageProjectUrl = metadata.ProjectUrl;
            ppp.PackageTags = metadata.Tags;
            ppp.Version = metadata.Version;
        }
    }

    public class NuGetDomain
    {
        public const string NuSpecFileName = "package.nuspec";
    }

}
