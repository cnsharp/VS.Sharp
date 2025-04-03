using System.Drawing;
using System.IO;
using System.Resources;
using stdole;

namespace CnSharp.VisualStudio.Extensions.Resources
{
    public static class ResourceHelper
    {
        public static Bitmap LoadBitmap(this ResourceManager rm, string name)
        {
            var map = (Bitmap)rm.GetObject(name);
            return map;
        }

        public static Bitmap LoadBitmapFromBytes(this ResourceManager rm, string name)
        {
            var bytes = (byte[])rm.GetObject(name);
            return ConvertBytesToBitmap(bytes);
        }

        public static Bitmap ConvertBytesToBitmap(byte[] bytes)
        {
            using (var ms = new MemoryStream(bytes))
            {
                return new Bitmap(ms);
            }
        }

        public static StdPicture LoadPicture(this ResourceManager rm, string name)
        {
            var map = rm.LoadBitmap(name);
            return (StdPicture)ImageConverter.ImageToIPicture(map);
        }

        public static StdPicture LoadPictureFromBytes(this ResourceManager rm, string name)
        {
            var map = rm.LoadBitmapFromBytes(name);
            return (StdPicture)ImageConverter.ImageToIPicture(map);
        }

    }
}