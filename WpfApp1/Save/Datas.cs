using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SalaryReport.Save
{
    [Serializable]
    public class Datas
    {
        public string Login { get; set; }
        public string Password { get; set; }
        public string PathToCopy { get; set; }
        public string Currency { get; set; }
        public string CurrencyZP { get; set; }
        public string CurrencyHoliday { get; set; }
        public string CurrencyHoliday2 { get; set; }
        public string CurrencyHoliday3 { get; set; }
        public string DateHoliday { get; set; }
        public string DateHoliday2 { get; set; }
        public string DateHoliday3 { get; set; }
        public string DateZp { get; set; }
        public string DateAvans { get; set; }
        public string FileFolder { get; set; }
        public string SettingsFolder { get; set; }
        public string EmailText { get; set; }
    }
}
