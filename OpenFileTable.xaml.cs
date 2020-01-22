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
    /// 
    public partial class OpenFileTable : Window
    {
        List<string> InputFileName = new List<string>();
        BitmapImage BitmapOpenFile =new BitmapImage(new Uri("OpenFile.png", UriKind.Relative));
        BitmapImage BitmapOpenFileDisable = new BitmapImage(new Uri("OpenFile_disable.png", UriKind.Relative));
        List<InputDataFile> IDFs = new List<InputDataFile>();
        List<int> IDFsIndex = new List<int>();
        InputDataFile IDF = new InputDataFile();
        public OpenFileTable(string[] dataString=null)
        {
            InitializeComponent();
            InputFileName.Clear();
            IDFsIndex.Clear();
            if(dataString!=null)
            {
                if(dataString.Length>0)
                    AddFilesToOpen(dataString);
            }
            dataGridExport.ItemsSource = IDF.InputDataFileRows;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textBoxFileName.Text = "";
            if(dataGridExport.Columns.Count>0)
            {
                dataGridExport.Columns[0].Header = "Дата проведения";
                dataGridExport.Columns[1].Header = "Время проведения";
                dataGridExport.Columns[2].Header = "Преподаватель";
                dataGridExport.Columns[3].Header = "Номер группы";
                dataGridExport.Columns[4].Header = "Категория слушателей";
                dataGridExport.Columns[5].Header = "Место проведения";
            }
            dataGridExport.UpdateLayout();
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
                if(listBoxInputFiles.Items.Count>0)
                {
                    listBoxInputFiles.SelectedIndex= listBoxInputFiles.Items.Count - 1;
                }
            }
        }

        private void AddFilesToOpen(string[] FileNames, bool Recursia=false)
        {
            textBoxFileName.Text = "";
            for (int i = 0; i < FileNames.Length; i++)
            {
                if (File.Exists(FileNames[i]))
                {
                    string buf = FileNames[i].Substring(FileNames[i].LastIndexOf('.') + 1);
                    if (buf == "xlsx" || buf == "xls")
                    {
                        if(InputFileName.IndexOf(FileNames[i])<0)
                        {
                            InputFileName.Add(FileNames[i]);
                            StackPanel stk = new StackPanel();
                            stk.Orientation = Orientation.Horizontal;
                            Image img = new Image();
                            img.Width = 20;
                            img.Height = 20;
                            img.Margin = new Thickness(0, 0, 5, 0);
                            if (ReadFile(FileNames[i]))
                                img.Source = BitmapOpenFile;
                            else
                                img.Source = BitmapOpenFileDisable;
                            ToolTip ttpi = new ToolTip();
                            ttpi.Content = "Подходит для экспорта данных";
                            img.ToolTip = ttpi;
                            TextBlock tbl = new TextBlock();
                            tbl.Text = FileNames[i].Substring(FileNames[i].LastIndexOf('\\') + 1);
                            ToolTip ttpt = new ToolTip();
                            tbl.ToolTip = ttpt;
                            ttpt.Content = FileNames[i];
                            stk.Children.Add(img);
                            stk.Children.Add(tbl);
                            listBoxInputFiles.Items.Add(stk);
                            //textBoxFileName.Text += FileNames[i] + "|";
                        }
                    }
                }
                else
                {
                    if (Directory.Exists(FileNames[i]))
                    {
                        if (!Recursia)
                        {
                            string[] SubfileNames = Directory.GetFiles(FileNames[i]);
                            AddFilesToOpen(SubfileNames, true);
                        }
                    }
                }                    
            }
        }

        private void buttonOpenFile_Click(object sender, RoutedEventArgs e)
        {
            string[] dataString = textBoxFileName.Text.Split('|');
            AddFilesToOpen(dataString);
        }

        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(InputFileName[listBoxInputFiles.SelectedIndex]);
        }

        private void buttonDeleteFile_Click(object sender, RoutedEventArgs e)
        {
            if(listBoxInputFiles.SelectedItem!=null)
            {
                int Ind = listBoxInputFiles.SelectedIndex;
                InputFileName.RemoveAt(Ind);
                IDFs.RemoveAt(Ind);
                listBoxInputFiles.Items.RemoveAt(Ind);
                if (listBoxInputFiles.Items.Count > Ind)
                {
                    listBoxInputFiles.SelectedIndex = Ind;
                }
                else
                {
                    if(listBoxInputFiles.Items.Count > 0)
                    {
                        listBoxInputFiles.SelectedIndex = listBoxInputFiles.Items.Count - 1;
                    }
                    else
                        dataGridExport.ItemsSource = null;
                }
            }
            listBoxInputFiles.UpdateLayout();
            dataGridExport.UpdateLayout();
        }

        private bool ReadFile(string FileName)
        {
            InputDataFile tempIDF = new InputDataFile();
            if (tempIDF.OpenFile(FileName))
            {
                IDFs.Add(new InputDataFile(FileName));
                //IDFs.Last().OpenFile(FileName);
                return true;
            }                
            else
            {
                IDFs.Add(null);
                return false;
            }
                
                
        }
        private void listBoxInputFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int Ind = listBoxInputFiles.SelectedIndex;
            if(Ind>=0)
            {
                if (IDFs[Ind] != null)
                    dataGridExport.ItemsSource = IDFs[Ind].InputDataFileRows;
                else
                    dataGridExport.ItemsSource = null;
            }
            else
            {
                dataGridExport.ItemsSource = null;
            }
            dataGridExport.UpdateLayout();
        }
    }
}
