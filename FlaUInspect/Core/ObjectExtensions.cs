using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlaUInspect.Core
{
    public static class ObjectExtensions
    {
        /// <summary>Serializes the object to a JSON string.</summary>
        /// <returns>A JSON string representation of the object.</returns>
        public static string ToJson(this object value)
        {
            var settings = new JsonSerializerSettings
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver(),
                Converters = new List<JsonConverter> { new StringEnumConverter() },
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore//,
                                                                    //MaxDepth=1

            };
            return JsonConvert.SerializeObject(value, settings);
        }
        public static string ToStringNz(this string value, string DefaultReturn)
        {
            value += "";
            return value.Length == 0 ? DefaultReturn : value;
        }

        public static string ReplaceInvalidChars(this string filename)
        {
            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));
        }

    }
}
