using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для Options.xaml
    /// </summary>
    public partial class Options : Window
    {
        OptionMain OptM;
        public Options()
        {
            
            InitializeComponent();
            //frame.Source = new Uri("OptionMain.xaml",UriKind.Relative);
            OptM = new OptionMain(this as Options);
            frame.Content = OptM;
            //MessageBox.Show(OpM.ToString()+"\n"+frame.ToString());
        }

        public void ChangeOptions()
        {
            buttonOK.IsEnabled = true;
        }

        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
            if (OptM.change)
                OptM.submitChanges();
            this.DialogResult = true;
            Properties.Settings.Default.Save();
        }

        private void buttonDefault_Click(object sender, RoutedEventArgs e)
        {
            OptM.defaultOptions();
        }
    }
}
