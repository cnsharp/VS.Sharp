using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CnSharp.VisualStudio.Extensions.Projects
{
    public class ProjectProperties : PackageProjectProperties
    {
        public string OutputType { get; set; }
        public string PlatformTarget { get; set; }
        public string TargetFramework { get; set; }
        public bool UseWindowsForms { get; set; }
    }
}
