using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace SalaryReport
{
    class XmlParser
    {
        public string GetCurrency (XmlDocument doc)
        {
            XmlElement xRoot = doc.DocumentElement;
            var nodeList = doc.GetElementsByTagName("Currency");
            string result = null;
            foreach (XmlNode el in nodeList)
            {
                if (el.ChildNodes[3].InnerText == "Доллар США")
                {
                    result = el.ChildNodes[4].InnerText;
                    break;
                }
            }

            return result;

        }
    }
}
