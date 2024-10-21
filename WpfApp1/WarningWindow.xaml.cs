using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using salary3Offices;

namespace SalaryReport
{
    /// <summary>
    /// Interaction logic for WarningWindow.xaml
    /// </summary>
    public partial class WarningWindow : Window
    {
        

        public WarningWindow()
        {
            InitializeComponent();
        }

        private void btnYes_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        private void btnNo_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton presed = (RadioButton)sender;
            string sendFirm = presed.Content.ToString();

            if(sendFirm == "Artezio")
            {
                Helper.artOrVega = true;
            }
            else if(sendFirm == "VegaSoft")
            {
                Helper.artOrVega = false;
            }
            else
            {
                Helper.artOrVega = true;
            }
        }
    }
}
