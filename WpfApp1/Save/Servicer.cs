using System;
using System.Net;
using System.Threading.Tasks;
using System.Xml;

namespace SalaryReport.Save
{
    public class Servicer
    {
        public XmlDocument GetXmlCurencyData(string url)
        {
            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(response.GetResponseStream());
                return xmlDocument;
            }
            catch (Exception e)
            {
                return null;
            }
        }
    }
}