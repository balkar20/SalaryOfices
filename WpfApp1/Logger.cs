using System;
using System.IO;
using System.Text;

namespace salary3Offices
{
    public static class Logger
    {
        public static StringBuilder LogString = new StringBuilder();

        public static void Out(string str)
        {
            LogString.Append(str).Append(Environment.NewLine);
        }

		public static string Save(string folder)
		{
			string fileName =
				new StringBuilder(folder).Append("\\log_").Append(DateTime.Now.ToString("dMMMyyyy_HHmmss")).Append(".txt").ToString();
			using (StreamWriter outfile = new StreamWriter(fileName))
            {
                outfile.Write(LogString.ToString());
            }

		    LogString.Clear();
			return fileName;
		}
    }
}
