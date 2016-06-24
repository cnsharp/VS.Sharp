using System;
using System.Xml.Serialization;

namespace CnSharp.VisualStudio.Extensions.Projects
{
    [Serializable]
    public class ProjectAssemblyInfo : IComparable<ProjectAssemblyInfo>
    {
        [XmlIgnore]
        public EnvDTE.Project Project { get; set; }

        public string ProjectName
        {
            get { return Project == null ? string.Empty : Project.Name; }
        }
        public string Version { get; set; }
        public string FileVersion { get; set; }
        public string Copyright { get; set; }
        public string ProductName { get; set; }
        public string Company { get; set; }
        public string Title { get; set; }

        public string Description { get; set; }


        public int CompareTo(ProjectAssemblyInfo other)
        {
            if (String.CompareOrdinal(Version, other.Version) != 0)
                return -1;
            if (String.CompareOrdinal(FileVersion, other.FileVersion) != 0)
                return -1;
            if (String.CompareOrdinal(Copyright, other.Copyright) != 0)
                return -1;
            if (String.CompareOrdinal(ProductName, other.ProductName) != 0)
                return -1;
            if (String.CompareOrdinal(Company, other.Company) != 0)
                return -1;
            if (String.CompareOrdinal(Title, other.Title) != 0)
                return -1;
            if (String.CompareOrdinal(Description, other.Description) != 0)
                return -1;
            return 0;
        }

        public ProjectAssemblyInfo Copy()
        {
            return new ProjectAssemblyInfo
            {
                Project = Project,
                Company = Company,
                Copyright = Copyright,
                FileVersion = FileVersion,
                ProductName = ProjectName,
                Title = Title,
                Version = Version
            };
        }

        public void Save()
        {
            if(Project == null)
                throw new InvalidOperationException("No project binding.");
            Project.ModifyAssemblyInfo(this);
        }
    }
}