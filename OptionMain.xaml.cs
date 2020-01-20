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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для OptionMain.xaml
    /// </summary>
    public partial class OptionMain : Page
    {
        public OptionMain()
        {
            InitializeComponent();
        }

        private void buttonBrowseMainFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "РАСП";
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Все (.*)|*.*|Книга Excel (.xlsx)|*.xlsx|Книга Excel 97-2003 (.xls)|*.xls";
            if(dlg.ShowDialog()==true)
            {
                textBoxSettingPath.Text = dlg.FileName;
            }
        }
    }
}
