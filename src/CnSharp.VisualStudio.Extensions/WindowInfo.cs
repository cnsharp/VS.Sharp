using System;
using EnvDTE;

namespace CnSharp.VisualStudio.Extensions
{
    public class WindowInfo
    {
        public string Caption { get; set; }
        [Obsolete]
        public bool IsDocument { get; set; }
        public string AssemblyLocation { get; set; }
        public string ClassName { get; set; }
        [Obsolete]
        public int Width { get; set; }
        [Obsolete]
        public int Height { get; set; }
        [Obsolete]
        public string Guid { get; set; }
        public object Picture { get; set; }
        [Obsolete]
        public Window LinkToWindow { get; set; }
        public string Id { get; set; }
    }
}
