using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace a4PrintTag
{
    public class JsonTransform
    {
        /// <summary>
        /// 将一个或者一list的对象，序列化为json串
        /// </summary>
        /// <param name="listobj"></param>
        /// <returns></returns>
        public static string ObjectToJson(object value)
        {
            JsonSerializer serializer = new JsonSerializer();
            StringWriter sw = new StringWriter();

            serializer.Serialize(new JsonTextWriter(sw), value);
            return sw.GetStringBuilder().ToString();
        }
        /// <summary>
        /// 将一个DataTable转化为json串
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string DataTableToJson(DataTable dt)
        {
            JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
            ArrayList arrayList = new ArrayList();
            foreach (DataRow dataRow in dt.Rows)
            {
                Dictionary<string, object> dictionary = new Dictionary<string, object>();
                foreach (DataColumn dataColumn in dt.Columns)
                {
                    dictionary.Add(dataColumn.ColumnName, dataRow[dataColumn.ColumnName]);
                }
                arrayList.Add(dictionary);
            }
            return javaScriptSerializer.Serialize(arrayList);
        }
        /// <summary>
        /// 将json转换为对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static T JsonToAnything<T>(string data)
        {
            System.Web.Script.Serialization.JavaScriptSerializer json = new System.Web.Script.Serialization.JavaScriptSerializer();
            return json.Deserialize<T>(data);
        }
    }
}
