using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace LocalPDF.Services
{
    internal static class JsonFiles
    {
        internal static T Read<T>(string path)
        {
            using (var stream = File.OpenRead(path)) return (T)new DataContractJsonSerializer(typeof(T)).ReadObject(stream);
        }

        internal static void Write<T>(string path, T value)
        {
            using (var stream = File.Create(path)) new DataContractJsonSerializer(typeof(T)).WriteObject(stream, value);
        }
    }
}
