using System;
using System.IO;

namespace CnSharp.VisualStudio.Extensions.Util
{
    /// <summary>
    /// Provides utility methods for handling file and directory paths.
    /// </summary>
    public static class PathUtil
    {
        /// <summary>
        /// Converts a relative path to an absolute path based on the specified base directory.
        /// </summary>
        /// <param name="path">The path to convert.</param>
        /// <param name="baseDir">The base directory to resolve the path against.</param>
        /// <returns>The absolute path.</returns>
        public static string ToAbsolutePath(string path, string baseDir)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Contains(":\\"))
            {
                return path;
            }
            Directory.SetCurrentDirectory(Path.GetDirectoryName(baseDir));
            return Path.GetFullPath(path);
        }

        /// <summary>
        /// Gets the relative path from the base directory to the specified path.
        /// </summary>
        /// <param name="path">The target path.</param>
        /// <param name="baseDir">The base directory.</param>
        /// <returns>The relative path from the base directory to the target path.</returns>
        public static string GetRelativePath(string path, string baseDir)
        {
            var uri = new Uri(path);
            var baseUri = new Uri(baseDir);
            var relativePath = baseUri.MakeRelativeUri(uri);
            return relativePath.ToString();
        }
    }
}
                