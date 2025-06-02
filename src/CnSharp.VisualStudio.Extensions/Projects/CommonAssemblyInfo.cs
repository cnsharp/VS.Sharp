using System;

namespace CnSharp.VisualStudio.Extensions.Projects
{
    /// <summary>
    /// Common properties of assembly information.
    /// convention:The properties of this class are named to match the properties of the AssemblyInfo.cs file but remove the prefix 'Assembly'.
    /// e.g. AssemblyVersion becomes Version, AssemblyCompany becomes Company, etc.
    /// </summary>
    [Serializable]
    public class CommonAssemblyInfo : IComparable<CommonAssemblyInfo>
    {
        public string Version { get; set; }
        public string Copyright { get; set; }
        public string Product { get; set; }
        public string Company { get; set; }
        public string Trademark { get; set; }

        public string InformationalVersion { get; set; }

        #region Implementation of IComparable<in CommonAssemblyInfo>

        public int CompareTo(CommonAssemblyInfo other)
        {
            if (string.CompareOrdinal(Version, other.Version) != 0)
                return -1;
            if (string.CompareOrdinal(Copyright, other.Copyright) != 0)
                return -1;
            if (string.CompareOrdinal(Product, other.Product) != 0)
                return -1;
            if (string.CompareOrdinal(Company, other.Company) != 0)
                return -1;
            if (string.CompareOrdinal(Trademark, other.Trademark) != 0)
                return -1;
            if (string.CompareOrdinal(InformationalVersion, other.InformationalVersion) != 0)
                return -1;
            return 0;
        }

        #endregion
    }
}