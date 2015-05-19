using EnvDTE;

namespace CnSharp.VisualStudio.Extensions
{
    public class WindowInfo
    {
        public WindowInfo()
        {
            Guid = "{" + System.Guid.NewGuid() + "}";
        }
        public string Caption { get; set; }
        public bool IsDocument { get; set; }
        public string AssemblyLocation { get; set; }
        public string ClassName { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string Guid { get; set; }
        public object Picture { get; set; }
        public Window LinkToWindow { get; set; }
    }
}
