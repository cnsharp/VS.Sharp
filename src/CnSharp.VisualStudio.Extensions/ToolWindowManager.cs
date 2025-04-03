using System;
using System.Collections.Generic;

namespace CnSharp.VisualStudio.Extensions
{
    public class ToolWindowManager
    {
        private static readonly Dictionary<string, int> WindowBaseIds = new Dictionary<string, int>();
        private static readonly Dictionary<string, int> WindowIdUsed = new Dictionary<string, int>();
        private static readonly Dictionary<string, int> WindowIds = new Dictionary<string, int>();

        public static void RegisterWindowBaseId(Type type, int baseId)
        {
            WindowBaseIds[type.FullName] = baseId;
        }

        public static int GetWindowId(Type type, string uri)
        {
            var key = type.FullName + "/" + uri;
            if (WindowIds.ContainsKey(key))
            {
                return WindowIds[key];
            }

            if (WindowIdUsed.ContainsKey(type.FullName))
            {
                var id = WindowIdUsed[type.FullName] + 1;
                WindowIdUsed[type.FullName] = id;
                WindowIds[key] = id;
                return id;
            }
            var baseId = WindowBaseIds[type.FullName];
            WindowIdUsed[type.FullName] = baseId;
            WindowIds[key] = baseId;
            return baseId;
        }
    }
}
