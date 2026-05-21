using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace HNRoadFormatConverter.Commons
{
    public  static  class JsonHelper
    {

       
     
   
        public static string Readjsonstring (string jsonPath, string key)
        {
            
         
            using (System.IO.StreamReader file = System.IO.File.OpenText(jsonPath))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject o = (JObject)JToken.ReadFrom(reader);
                    var value = o[key].ToString();
                    return value;
                }
            }
        }
        /// <summary>
        /// 把对象转换为JSON字符串
        /// </summary>
        /// <param name="o">对象</param>
        /// <returns>JSON字符串</returns>
        public static string ToJSON(this object o)
        {
            if (o == null)
            {
                return null;
            }
            return JsonConvert.SerializeObject(o);
        }
        /// <summary>
        /// 把Json文本转为实体
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="input"></param>
        /// <returns></returns>
        public static T FromJSON<T>(this string input)
        {
            try
            {
                T deserializedProduct = (T)JsonConvert.DeserializeObject(input, typeof(T));
                return deserializedProduct;
            }
            catch (Exception)
            {
                return default(T);
            }
        }
    }
}
