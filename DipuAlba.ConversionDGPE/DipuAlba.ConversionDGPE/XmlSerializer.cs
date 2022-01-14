using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DipuAlba.ConversionDGPE
{
    public sealed class XmlSerializer
    {

        /// <summary>
        /// Convert a class state into XML
        /// </summary>
        /// <typeparam name="T">The type of object</typeparam>
        /// <param name="obj">The object to serilize</param>
        /// <param name="sConfigFilePath">The path to the XML</param>
        public static void Serialize<T>(T obj, string sConfigFilePath)
        {
            var XmlBuddy = new System.Xml.Serialization.XmlSerializer(typeof(T));
            var MySettings = new System.Xml.XmlWriterSettings();
            MySettings.Indent = true;
            MySettings.CloseOutput = true;
            var MyWriter = System.Xml.XmlWriter.Create(sConfigFilePath, MySettings);
            XmlBuddy.Serialize(MyWriter, obj);
            MyWriter.Flush();
            MyWriter.Close();
        }



        /// <summary>
        /// Restore a class state from XML
        /// </summary>
        /// <typeparam name="T">The type of object</typeparam>
        /// <param name="xml">the path to the XML</param>
        /// <returns>The object to return</returns>
        public static T Deserialize<T>(string xml)
        {
            var XmlBuddy = new System.Xml.Serialization.XmlSerializer(typeof(T));
            var fs = new FileStream(xml, FileMode.Open);
            var reader = new System.Xml.XmlTextReader(fs);
            if (XmlBuddy.CanDeserialize(reader))
            {
                T tempObject = (T)XmlBuddy.Deserialize(reader);
                reader.Close();
                return tempObject;
            }
            else
            {
                return default;
            }
        }
    }
}
