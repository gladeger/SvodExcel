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
    /// Логика взаимодействия для OpenFileTable.xaml
    /// </summary>
    public partial class OpenFileTable : Window
    {
        public OpenFileTable()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Height = 400;
            textBoxFileName.Text = "";
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                //string dataString = (string)e.Data.GetData(DataFormats.StringFormat);
                string[] dataString = (string[])e.Data.GetData(DataFormats.FileDrop);
                
                MessageBox.Show(dataString[0]);
            }
             
        }

        private void buttonBrowseMainFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            //dlg.FileName = "РАСП";
            dlg.Filter = "Книга Excel (.xlsx)|*.xlsx|Книга Excel 97-2003 (.xls)|*.xls|Все (.*)|*.*";
            dlg.DefaultExt = ".xlsx";
            dlg.Multiselect = true;
            if (dlg.ShowDialog() == true)
            {
                textBoxFileName.Text = "";
                for (int i=0;i<dlg.FileNames.Length;i++)
                {
                    textBoxFileName.Text += dlg.FileNames[i]+";";
                }
                
            }
        }
    }
}
