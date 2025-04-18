using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace CnSharp.VisualStudio.Extensions.Util
{
    public class XmlSerializerHelper
    {
        public static T Copy<T>(T t)
        {
            return LoadObjectFromXmlString<T>(GetXmlStringFromObject(t));
        }

        public static string GetXmlStringFromObject<T>(T obj)
        {
            using (var stream = new MemoryStream())
            {
                using (var writer = new XmlTextWriter(stream, Encoding.UTF8))
                {
                    writer.Indentation = 3;
                    writer.IndentChar = ' ';
                    writer.Formatting = Formatting.Indented;
                    new XmlSerializer(obj.GetType()).Serialize(writer, obj);
                    var xml = Encoding.UTF8.GetString(stream.ToArray());
                    var index = xml.IndexOf("?>");
                    if (index > 0)
                    {
                        xml = xml.Substring(index + 2);
                    }
                   return xml.Trim();
                }
            }
        }

        public static T LoadObjectFromXml<T>(string filename)
        {
            using (var reader = new StreamReader(filename, Encoding.UTF8))
            {
                var serializer = new XmlSerializer(typeof (T));
                return (T) serializer.Deserialize(reader);
            }
        }

        public static T LoadObjectFromXmlString<T>(string xml)
        {
            var serializer = new XmlSerializer(typeof (T));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                return (T) serializer.Deserialize(stream);
            }
        }

        public static void SaveObjectToXml<T>(T obj, string filename)
        {
            using (var writer = new StreamWriter(filename, false, Encoding.UTF8))
            {
                var serializer = new XmlSerializer(obj.GetType());
                serializer.Serialize(writer, obj);
            }
        }
    }
}