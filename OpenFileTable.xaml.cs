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
using System.IO;

namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для OpenFileTable.xaml
    /// </summary>
    public partial class OpenFileTable : Window
    {
        List<string> InputFileName = new List<string>();
        BitmapImage BitmapOpenFile =new BitmapImage(new Uri("OpenFile.png", UriKind.Relative));
        BitmapImage BitmapOpenFileDisable = new BitmapImage(new Uri("OpenFile_disable.png", UriKind.Relative));
        public OpenFileTable()
        {
            InitializeComponent();
            InputFileName.Clear();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textBoxFileName.Text = "";
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
                AddFilesToOpen(dlg.FileNames);
            }
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                //string dataString = (string)e.Data.GetData(DataFormats.StringFormat);
                string[] dataString = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddFilesToOpen(dataString);
            }
        }

        private void AddFilesToOpen(string[] FileNames, bool Recursia)
        {
            textBoxFileName.Text = "";
            for (int i = 0; i < FileNames.Length; i++)
            {
                if (File.Exists(FileNames[i]))
                {
                    if (File.GetAttributes(FileNames[i]).HasFlag(FileAttributes.Directory))
                    {
                        if(!Recursia)
                        {
                            //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        }
                    }
                    else
                    {
                        InputFileName.Add(FileNames[i]);
                        StackPanel stk = new StackPanel();
                        stk.Orientation = Orientation.Horizontal;
                        Image img = new Image();
                        img.Width = 20;
                        img.Height = 20;
                        img.Source = BitmapOpenFile;
                        TextBlock tbl = new TextBlock();
                        tbl.Text = FileNames[i].Substring(FileNames[i].LastIndexOf('\\') + 1);
                        ToolTip ttp = new ToolTip();
                        tbl.ToolTip = ttp;
                        ttp.Content = FileNames[i];
                        stk.Children.Add(img);
                        stk.Children.Add(tbl);
                        listBoxInputFiles.Items.Add(stk);
                        //textBoxFileName.Text += FileNames[i] + "|";
                    }
                }
            }
        }
        private void AddFilesToOpen(string[] FileNames)
        {
            AddFilesToOpen(FileNames, false);
        }

        private void buttonOpenFile_Click(object sender, RoutedEventArgs e)
        {
            string[] dataString = textBoxFileName.Text.Split('|');
            AddFilesToOpen(dataString);
        }

        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
        }
    }
}
