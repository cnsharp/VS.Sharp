using System;
using System.IO;
using System.Reflection;
using System.Resources;
using CnSharp.VisualStudio.Extensions.Commands;

namespace CnSharp.VisualStudio.Extensions
{
    public class Plugin
    {
        private Assembly _assembly;

        public Assembly Assembly
        {
            get { return _assembly; }
            set
            {
                _assembly = value;
                Location = _assembly == null ? null : Path.GetDirectoryName(new UriBuilder(_assembly.CodeBase).Path);
            }
        }

        public string Location { get; private set; }

        public CommandConfig CommandConfig { get; set; }

        public CommandManager CommandManager { get; set; }

        public ResourceManager ResourceManager
        {
            get
            {
                return  (CommandConfig != null && Assembly != null) ?
                    new ResourceManager(CommandConfig.ResourceManager, Assembly) : 
                    null;
            }
        }
    }
}