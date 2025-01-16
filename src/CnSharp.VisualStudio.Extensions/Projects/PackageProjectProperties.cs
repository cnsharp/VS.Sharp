using System;

namespace CnSharp.VisualStudio.Extensions.Projects
{
    /// <summary>
    /// Package properties of project,see <PropertyGroup></PropertyGroup> section in *proj file.
    /// </summary>
    public class PackageProjectProperties
    {
        public bool GeneratePackageOnBuild { get; set; }
        public bool IncludeSymbols { get; set; } 
        public bool PackageRequireLicenseAcceptance { get; set; }
        public string AssemblyVersion { get; set; }
        public string Authors { get; set; }
        public string Company { get; set; }
        public string Copyright { get; set; }
        public string Description { get; set; }
        public string FileVersion { get; set; }
        public string Name { get; set; }
        public string NeutralLanguage { get; set; }
        public string PackageIcon { get; set; }
        [Obsolete]
        public string PackageIconUrl { get; set; }
        public string PackageId { get; set; }
        public string PackageLicenseExpression { get; set; }
        public string PackageLicenseUrl { get; set; }
        public string PackageProjectUrl { get; set; }
        public string PackageReadmeFile { get; set; }
        public string PackageReleaseNotes { get; set; }
        public string PackageTags { get; set; }
        public string Product { get; set; }
        public string RepositoryType { get; set; }
        public string RepositoryUrl { get; set; }
        public string Version { get; set; }
    }
}