using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel ;

namespace salary3Offices
{
    class OpEvantArgs
    {
        public string Message { get; set; }
    }
    class Helper
    {
        public static string from = string.Empty;
        public static string fromsign = string.Empty;
        public static string smtphost = string.Empty;
        public static Dictionary<string, string> to = new Dictionary<string, string>();
        //public static int i = 0;

        public static NetworkCredential login;
        public static int port = 0;

        public static string currency = "";
        public static string currencyZP = "";
        public static string curencyHolliday = "";
        public static string curencyHolliday2 = "";
        public static string curencyHolliday3 = "";

        public static string dateOfZpString = "";
        public static string dateOfAvansString = "";
        public static string dateOfHollydayString = "";
        public static string dateOfHollydayString2 = "";
        public static string dateOfHollydayString3 = "";


        public static DateTime curentDate = DateTime.Now;

        static int cy = curentDate.Year;
        static int cm = curentDate.Month;
        static int cd = curentDate.Day;

        static int cmAvanse = cm - 1 > 0 ? cm - 1 : 12;
        static int cyAvance = cmAvanse != 12 ? cy : cy - 1;


        public static DateTime dateAvance = new DateTime(cyAvance, cmAvanse, 25);
        public static DateTime dateZP = new DateTime(cy, cm, 10);
        private static int count;

        public static ObservableCollection<string> dates = new ObservableCollection<string> { dateOfZpString };

        public static string patternDate = @"^(((0[1-9]|[12]\d|3[01])\.(0[13578]|1[02])\.((19|[2-9]\d)\d{2}))|((0[1-9]|[12]\d|30)\.(0[13456789]|1[012])\.((19|[2-9]\d)\d{2}))|((0[1-9]|1\d|2[0-8])\.02\.((19|[2-9]\d)\d{2}))|(29\.02\.((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))))$";
        public static bool artOrVega;

        public delegate void MyHandler(string Messsage);
        public static event MyHandler Op;

        private static Application xlApp;
        private static Workbook xlWorkBook;
        private static Worksheet xlWorkSheet;
        private static object misValue = System.Reflection.Missing.Value;
        private static Application xlAppNew;
        private static Workbook xlWorkBookNew;
        private static Worksheet xlWorkSheetNew;
        public static string pathToCopyExcel;

        public static string sent = "\\Sent\\";

        public static string ConvertXslToCsv(string settingfile, string receiptFile, string emailtext)
        {
            string folder = Path.GetDirectoryName(receiptFile);

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(receiptFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            string end = string.Empty;
            if (xlWorkSheet != null)
            {
                int rowStart = 0;
                int rowEnd = 0;
                int pos = 0;
                Microsoft.Office.Interop.Excel.Range last = xlWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                Microsoft.Office.Interop.Excel.Range excelCell = xlWorkSheet.get_Range("A1", last);
                var cellValue = (Object[,])excelCell.Value;
                string rstart = string.Empty;
                string rend = string.Empty;
                count = to.Count;
                foreach (object s in (Array)cellValue)
                {
                    pos++;
                    if (s != null && s.ToString().Contains("Расчетный листок"))
                    {
                        if (rowStart == 0)
                        {
                            rowStart = 1 + pos / 5;
                        }
                        else
                        {
                            rowEnd = pos / 5;

                            rstart = "A" + rowStart;
                            rend = "E" + rowEnd;

                            CopyRange(rstart, rend, rowStart, rowEnd, folder, emailtext);

                            rowStart = rowEnd + 1;
                        }
                    }
                }
                CopyRange("A" + (rowEnd + 1), "E" + excelCell.Rows.Count + 1, rowEnd, excelCell.Rows.Count, folder, emailtext);
            }
            xlWorkBook.Close();
            var pathToLogs = Environment.CurrentDirectory + @"..\..\..\logs";
            return Logger.Save(pathToLogs);
        }

        private static void CopyRange(string rstart, string rend, int rowStart, int rowEnd, string folder, string emailtext)
        {
            xlAppNew = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBookNew = xlApp.Workbooks.Add(misValue);
            xlWorkSheetNew = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range excelCellFrom = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.get_Range(rstart, rend);
            Microsoft.Office.Interop.Excel.Range excelCellNew = (Microsoft.Office.Interop.Excel.Range)xlWorkSheetNew.get_Range("A1", "E" + (rowEnd - rowStart));
            excelCellFrom.Copy(misValue);

            excelCellNew.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths,
                Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            excelCellNew.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAllUsingSourceTheme,
                Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            string emplNameVega;
            string emplPeriosVega;

            string vegaNamePattern = @"(оклад:)\s[0-9]{1,3}(?:[.,][0-9]{1,3})?\z";
            string vegaPeriosPattern = @"^(Расчетный листок  за )";

            Regex rgx = new Regex(vegaNamePattern);
            Regex rgx2 = new Regex(vegaPeriosPattern);

            //Находит ФИО 
            var emplname = (string)(excelCellNew.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Value;

            //Находит Период
            var emplperios = (string)(excelCellNew.Cells[2, 2] as Microsoft.Office.Interop.Excel.Range).Value;

            if (artOrVega == false)
            {
                emplNameVega = (string)(excelCellNew.Cells[2, 1] as Microsoft.Office.Interop.Excel.Range).Value;
                emplPeriosVega = (string)(excelCellNew.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range).Value;

                string newName = rgx.Replace(emplNameVega, "");
                string newPerios = rgx2.Replace(emplPeriosVega, "");

                //Находит ФИО 
                emplname = newName;

                //Находит Период
                emplperios = newPerios;
            }

            //Конфиг - файл + send -- Добавляет папку send
            bool folderExists = Directory.Exists(folder + sent);
            if (!folderExists)
                Directory.CreateDirectory(folder + sent);

            //Указываем имя файла который будет отправлен
            string fileToSent =
                new StringBuilder(pathToCopyExcel).Append(emplname).Append(".xls").ToString();
            //Сохраняем в файл sent

            try
            {
                xlWorkBookNew.SaveAs(fileToSent,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBookNew.Close(true, misValue, misValue);
            }
            catch (Exception e)
            {
                Logger.Out(String.Format("Ошибка при сохранении Excel для {0}: {1}", emplname, e.Message));
                Op(String.Format("Ошибка при сохранении Excel для {0}: {1}", emplname, e.Message));
            }
            

            bool errFlag = true;
            int counterFail = 60;
            while (errFlag)
            {
                try
                {
                    SendNew(fileToSent, from, emplname, emplperios, emailtext);
                    errFlag = false;
                }
                catch (Exception e)
                {
                    if (counterFail >= 0)
                    {
                        Logger.Out("Упс! Пробуем отправить еще раз!");
                        Op("Упс! Пробуем отправить еще раз!");
                        counterFail--;
                        Thread.Sleep(300);
                    }
                    else
                    {
                        Logger.Out(String.Format("Ошибка при попытке отправить письмо для {0}: {1}", emplname, e.Message));
                        Op(String.Format("Ошибка при попытке отправить письмо для {0}: {1}", emplname, e.Message));
                        errFlag = false;
                    }
                }
            }
        }

        public static void SendNew(string filename, string from, string employeeFullName, string period, string emailtext)
        {
            SmtpClient smtp = null;
            Attachment attachment = null;

            if (smtphost.Equals(string.Empty))
            {
                Logger.Out(String.Format("SMTP host не задан. Письмо не будет отправлено."));
                return;
            }
            if (!to.ContainsKey(employeeFullName))
            {
                Logger.Out(String.Format("Емейл не найден {0}", employeeFullName));
            }
            else
            {
                MailMessage message = new MailMessage(from, to[employeeFullName]);
                attachment = new Attachment(filename);

                message.Attachments.Add(attachment);

                StringBuilder mailBody = new StringBuilder();
                mailBody.AppendFormat("Уважаемый(ая) {0},", employeeFullName);
                mailBody.Append(Environment.NewLine).Append(Environment.NewLine);


                mailBody.AppendFormat("Ваш расчетный листок {0} находится во вложении." + "\n\n", period);
                if (currency != "")
                {
                    mailBody.AppendFormat("Курс доллара США на {0}: {1} BYN" + "\n\r", dateOfAvansString, currency);
                }
                if (currencyZP != "")
                {
                    mailBody.AppendFormat("Курс доллара США на {0}: {1} BYN" + "\n\r", dateOfZpString, currencyZP);
                }
                if (curencyHolliday != "")
                {
                    mailBody.AppendFormat("Курс доллара США на {0}: {1} BYN" + "\n\r", dateOfHollydayString, curencyHolliday);
                }
                if (curencyHolliday2 != "")
                {
                    mailBody.AppendFormat("Курс доллара США на {0}: {1} BYN" + "\n\r", dateOfHollydayString2, curencyHolliday2);
                }
                if (curencyHolliday3 != "")
                {
                    mailBody.AppendFormat("Курс доллара США на {0}: {1} BYN" + "\n\r", dateOfHollydayString3, curencyHolliday3);
                }
                mailBody.Append(emailtext);
                mailBody.Append(Environment.NewLine).Append(Environment.NewLine).Append("Kind regards,").Append(Environment.NewLine);
                mailBody.Append(fromsign);
                message.Body = mailBody.ToString();
                message.Subject = String.Format("Расчетный листок {0}", period);
                smtp = new SmtpClient(smtphost);
                smtp.Port = port;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = login;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                Logger.Out(String.Format(to[employeeFullName]));
                    smtp.Send(message);
                

                Logger.Out(String.Format("Расчетный листок для {0} был отправлен на адрес {1}", employeeFullName,
                                         to[employeeFullName]));
                Op(String.Format("Расчетный листок для {0} был отправлен на адрес {1}", employeeFullName,
                    to[employeeFullName]));
            }

        }

        public static void ReadSettings(string settingfile)
        {
            to.Clear();
            using (StreamReader reader = new StreamReader(settingfile, Encoding.Default))
            {
                string setting;

                while ((setting = reader.ReadLine()) != null)
                {
                    if (setting.StartsWith("FromEmail:"))
                    {
                        @from = setting.Replace("FromEmail:", string.Empty);
                    }
                    else if (setting.StartsWith("FromSignature:"))
                    {
                        fromsign = setting.Replace("FromSignature:", string.Empty);
                    }
                    else if (setting.StartsWith("SMTP:"))
                    {
                        smtphost = setting.Replace("SMTP:", string.Empty);
                    }
                    else
                    {
                        string patternForTwo = @"([а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+)(:)([a-zA-Z0-9._]+@artezio.com)";
                        string pattern = @"([а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+)(:)([a-zA-Z]+\.[a-zA-Z]+@artezio.com)";
                        string pattern1 = @"([а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+)(:)([a-zA-Z0-9._]+@gmail.com)";
                        string pattern2 = @"([а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+)(:)([a-zA-Z0-9._]+@yandex.ru)";
                        string pattern3 = @"([а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+\s+[а-яА-ЯёЁ]+)(:)([a-zA-Z0-9._]+@mail.ru)";
                        string pattern4 = @"([a-zA-Z0-9._]+)(:)([a-zA-Z0-9._]+[@][a-zA-z0-9._]+)";

                        Regex rgx = new Regex(pattern, RegexOptions.CultureInvariant);
                        MatchCollection matches = rgx.Matches(setting);
                        Regex rgx1 = new Regex(pattern1, RegexOptions.CultureInvariant);
                        MatchCollection matches1 = rgx1.Matches(setting);
                        Regex rgx2 = new Regex(pattern2, RegexOptions.IgnoreCase);
                        MatchCollection matches2 = rgx2.Matches(setting);
                        Regex rgx3 = new Regex(pattern3, RegexOptions.IgnoreCase);
                        MatchCollection matches3 = rgx3.Matches(setting);
                        Regex rgx4 = new Regex(pattern4, RegexOptions.IgnoreCase);
                        MatchCollection matches4 = rgx4.Matches(setting);
                        Regex rgxFor2 = new Regex(patternForTwo, RegexOptions.IgnoreCase);
                        MatchCollection matchesFor2 = rgxFor2.Matches(setting);
                        if (matches.Count > 0)
                        {
                            GroupCollection groups = matches[0].Groups;
                            to.Add(groups[1].ToString(), groups[3].ToString());
                            continue;
                        }
                        if (matchesFor2.Count > 0)
                        {
                            GroupCollection groups2 = matchesFor2[0].Groups;
                            to.Add(groups2[1].ToString(), groups2[3].ToString());
                            continue;
                        }
                        if (matches1.Count > 0)
                        {
                            GroupCollection groups = matches1[0].Groups;
                            to.Add(groups[1].ToString(), groups[3].ToString());
                            continue;
                        }
                        if (matches2.Count > 0)
                        {
                            GroupCollection groups = matches2[0].Groups;
                            to.Add(groups[1].ToString(), groups[3].ToString());
                            continue;
                        }
                        if (matches3.Count > 0)
                        {
                            GroupCollection groups = matches3[0].Groups;
                            to.Add(groups[1].ToString(), groups[3].ToString());
                            continue;
                        }
                        if (matches4.Count > 0)
                        {
                            GroupCollection groups = matches4[0].Groups;
                            to.Add(groups[1].ToString(), groups[3].ToString());
                        }
                    }
                }
            }
        }
    }
}
