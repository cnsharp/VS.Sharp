using System;
using System.IO;
using System.Reflection;
using System.Resources;
using CnSharp.VisualStudio.Extensions.Commands;

namespace CnSharp.VisualStudio.Extensions
{
    public class Plugin
    {
        public Guid Id { get; set; }

        public string Name { get; set; }
        
        private Assembly _assembly;

        public Assembly Assembly
        {
            get => _assembly;
            set
            {
                _assembly = value;
                Location = _assembly == null ? null : Path.GetDirectoryName(new UriBuilder(_assembly.CodeBase).Path);
            }
        }

        public string Location { get; set; }

        public CommandConfig CommandConfig { get; set; }

        public CommandManager CommandManager { get; set; }

        public ResourceManager ResourceManager =>
            CommandConfig != null && !string.IsNullOrWhiteSpace(CommandConfig.ResourceManager) && Assembly != null ?
                new ResourceManager(CommandConfig.ResourceManager, Assembly) : 
                null;
    }
}