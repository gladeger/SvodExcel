﻿using System;
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
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для OpenFileTable.xaml
    /// </summary>
    /// 
    public class NoneTeacher
    {
        public string Name { get; set; }
        public NoneTeacher(){}
    }

    public struct ErrorRowInIDFsAndLinkToOtherRow
    {
        public int ErrorFile;
        public int ErrorRow;
        public int LinkFile;
        public int LinkRow;
    }
public struct ErrorRowInIDFsAndLinkToRowOnSvod
    {
        public int ErrorFile;
        public int ErrorRow;
        public int LinkRow;
    }
    public partial class OpenFileTable : Window
    {
        List<string> InputFileName = new List<string>();
        BitmapImage BitmapOpenFile = new BitmapImage(new Uri(@"Images\OpenFile.png", UriKind.Relative));
        BitmapImage BitmapOpenFileDisable = new BitmapImage(new Uri(@"Images\OpenFile_disable.png", UriKind.Relative));
        List<InputDataFile> IDFs = new List<InputDataFile>();
        List<int> IDFsIndex = new List<int>();
        InputDataFile IDF = new InputDataFile();
        List<string> TimeTemplate = new List<string>();
        List<string> TeacherTemplate = new List<string>();
        public List<string> NoneTeacherTemplate = new List<string>();
        public List<NoneTeacher> NTT = new List<NoneTeacher>();
        public List<NoneTeacher> YTT = new List<NoneTeacher>();
        public List<DataTableRow> SvodData =new List<DataTableRow>();
        ulong countAllRecords = 0;
        private bool ClickToAddRow = true;
        public ListViewEditWindow LVEW = new ListViewEditWindow();
        public ListViewEditWindow LVEWY = new ListViewEditWindow();
        public List<ErrorRowInIDFsAndLinkToOtherRow> ERLR = new List<ErrorRowInIDFsAndLinkToOtherRow>();
        public List<ErrorRowInIDFsAndLinkToRowOnSvod> ERLS = new List<ErrorRowInIDFsAndLinkToRowOnSvod>();
        public OpenFileTable(string[] dataString = null, Window OwnerWindow = null)
        {
            InitializeComponent();
            StartListTimes();
            StartListTeacher();
            InputFileName.Clear();
            IDFsIndex.Clear();
            if (OwnerWindow != null)
                Owner = OwnerWindow;
            if (dataString != null)
            {
                if (dataString.Length > 0)
                {
                    AddFilesToOpen(dataString);
                    countAllRecords = 0;
                    for (int i = 0; i < IDFs.Count; i++)
                    {
                        //if(IDFs[i].InputDataFileRows!=null)
                        countAllRecords += (ulong)IDFs[i].InputDataFileRows.Count;
                    }
                    StatusStringCountRecordAllFile.Content = countAllRecords.ToString();
                }
            }
            dataGridExport.ItemsSource = IDF.InputDataFileRows;
            countAllRecords = 0;
            ((INotifyCollectionChanged)listBoxInputFiles.Items).CollectionChanged += listBoxInputFilesItemsChanges;
            //buttonFindDublicates.Visibility = Visibility.Collapsed;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textBoxFileName.Text = "";
            if (dataGridExport.Columns.Count > 0)
            {
                dataGridExport.Columns[0].Header = "Дата проведения";
                dataGridExport.Columns[1].Header = "Время проведения";
                dataGridExport.Columns[2].Header = "Преподаватель";
                dataGridExport.Columns[3].Header = "Номер группы";
                dataGridExport.Columns[4].Header = "Категория слушателей";
                dataGridExport.Columns[5].Header = "Место проведения";
            }
            dataGridExport.UpdateLayout();

            if (listBoxInputFiles.Items.Count > 0)
            {
                listBoxInputFiles.SelectedIndex = 0;
            }

            LVEW.Title = "Список игнорируемых преподавателей";
            LVEW.Owner = this;

            NTT.Clear();
            for (int i = 0; i < NoneTeacherTemplate.Count; i++)
            {
                NoneTeacher bNTT = new NoneTeacher();
                bNTT.Name = NoneTeacherTemplate[i];
                NTT.Add(bNTT);
            }


            Binding bind = new Binding();
            bind.Path = new PropertyPath(".");
            bind.Source = NTT;
            //bind.XPath = ".";
            bind.Mode = BindingMode.TwoWay;
            
            
            LVEW.dataGrid.ItemsSource = NTT;
            
            
            //labelTech.Content = NoneTeacherTemplate[0];

            LVEW.dataGrid.SetBinding(ItemsControl.ItemsSourceProperty, bind);
            CollectionViewSource.GetDefaultView(LVEW.dataGrid.ItemsSource).Refresh();
            LVEW.dataGrid.CanUserAddRows = true;
            LVEW.dataGrid.Columns[0].Header = "ФИО";
            LVEW.textBlockInfo.Text = "\tПреподаватели из этого списка игнорируются при добавлении новых записей из файлов, но только в том случае, если данного преподавателя еще нет в общем файле расписания.";

            /*string bufstr = "";
            for(int i=0;i<TimeTemplate.Count;i++)
            {
                bufstr += TimeTemplate[i] + "\n";
            }
            MessageBox.Show(bufstr);*/


            LVEWY.Title = "Список известных преподавателей";
            LVEWY.Owner = this;

            YTT.Clear();
            for (int i = 0; i < TeacherTemplate.Count; i++)
            {
                NoneTeacher bYTT = new NoneTeacher();
                bYTT.Name = TeacherTemplate[i];
                YTT.Add(bYTT);
            }


            Binding bindY = new Binding();
            bindY.Path = new PropertyPath(".");
            bindY.Source = YTT;
            //bind.XPath = ".";
            bindY.Mode = BindingMode.OneWay;


            LVEWY.dataGrid.ItemsSource = YTT;

            LVEWY.dataGrid.SetBinding(ItemsControl.ItemsSourceProperty, bindY);
            CollectionViewSource.GetDefaultView(LVEWY.dataGrid.ItemsSource).Refresh();
            LVEWY.dataGrid.CanUserAddRows = true;
            LVEWY.dataGrid.Columns[0].Header = "ФИО";
            LVEWY.textBlockInfo.Text = "\tПреподаватели из этого списка \bизвестны\b системе, часть из них взята из общего файла расписания, часть из добавляемых вами записей";
            LVEWY.buttonSingleInputHot.IsEnabled = false;
            LVEWY.dataGrid.IsReadOnly = true;
        }
        private void Window_Closed(object sender, EventArgs e)
        {
            string path = @".\ListNoneTeacher.dat";
            File.WriteAllLines(path, NoneTeacherTemplate);
            path = @".\ListTeacher.dat";
            File.WriteAllLines(path, TeacherTemplate);
        }

        private void StartListTimes()
        {
            string pathT = @".\ListTime.dat";

            if (!File.Exists(pathT))
            {
                File.WriteAllText(pathT, "10:00-16:40");
                File.AppendAllText(pathT, "\n" + "12:00-18:40");
            }

            TimeTemplate.Clear();
            TimeTemplate = File.ReadAllLines(pathT).ToList<string>();
            for (int i = 0; i < TimeTemplate.Count; i++)
            {
                if (TimeTemplate[i][0] == '0')
                {
                    TimeTemplate[i] = TimeTemplate[i].Substring(1);
                }
                TimeTemplate[i] = TimeTemplate[i].Replace(':', '.');
            }
        }
        private void StartListTeacher()
        {

            string path = @".\ListTeacher.dat";

            if (!File.Exists(path))
            {
                File.WriteAllText(path, "");
                File.WriteAllText(path, "\n" + "Moodle");
                File.AppendAllText(path, "\n" + "Пронина Л.Н.");
                File.AppendAllText(path, "\n" + "Григорьева А.И.");
            }
            TeacherTemplate.Clear();
            TeacherTemplate = File.ReadAllLines(path).ToList<string>();

            //MessageBox.Show(TeacherTemplate[TeacherTemplate.Count - 1]);

            YTT.Clear();
            for (int i = 0; i < TeacherTemplate.Count; i++)
            {
                NoneTeacher bYTT = new NoneTeacher();
                bYTT.Name = TeacherTemplate[i];
                YTT.Add(bYTT);
            }
            if (LVEWY.dataGrid != null)
                if (LVEWY.dataGrid.ItemsSource != null)
                {
                    CollectionViewSource.GetDefaultView(LVEWY.dataGrid.ItemsSource).Refresh();
                    LVEWY.dataGrid.UpdateLayout();
                }                   


            path = @".\ListNoneTeacher.dat";
            NoneTeacherTemplate.Clear();
            if (File.Exists(path))
            {
                NoneTeacherTemplate = File.ReadAllLines(path).ToList<string>();
            }
            NTT.Clear();
            for (int i = 0; i < NoneTeacherTemplate.Count; i++)
            {
                NoneTeacher bNTT = new NoneTeacher();
                bNTT.Name = NoneTeacherTemplate[i];
                NTT.Add(bNTT);
            }
            if(LVEW.dataGrid!=null)
                if(LVEW.dataGrid.ItemsSource!=null)
                {
                    CollectionViewSource.GetDefaultView(LVEW.dataGrid.ItemsSource).Refresh();
                    LVEW.dataGrid.UpdateLayout();
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
                AddFilesToOpen(dlg.FileNames);
                if (listBoxInputFiles.Items.Count > 0)
                {
                    listBoxInputFiles.SelectedIndex = listBoxInputFiles.Items.Count - 1;
                }
            }
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
                objBlur.Radius = 4;
                this.Effect = objBlur;
                UpdateLayout();

                //string dataString = (string)e.Data.GetData(DataFormats.StringFormat);
                string[] dataString = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddFilesToOpen(dataString);
                if (listBoxInputFiles.Items.Count > 0)
                {
                    listBoxInputFiles.SelectedIndex = listBoxInputFiles.Items.Count - 1;
                }

                this.Effect = null;
                UpdateLayout();
            }
        }

        public void AddFilesToOpen(string[] FileNames, bool Recursia = false)
        {
            textBoxFileName.Text = "";
            for (int i = 0; i < FileNames.Length; i++)
            {
                if (File.Exists(FileNames[i]))
                {
                    string buf = FileNames[i].Substring(FileNames[i].LastIndexOf('.') + 1);
                    if (buf == "xlsx" || buf == "xls")
                    {
                        if (InputFileName.IndexOf(FileNames[i]) < 0)
                        {
                            InputFileName.Add(FileNames[i]);
                            StackPanel stk = new StackPanel();
                            stk.Orientation = Orientation.Horizontal;
                            Image img = new Image();
                            img.Width = 20;
                            img.Height = 20;
                            img.Margin = new Thickness(0, 0, 5, 0);
                            ToolTip ttpi = new ToolTip();
                            ttpi.Content = "Не подходит для экспорта данных, ошибка в форме записи или невозможно прочесть данные";
                            img.Source = BitmapOpenFileDisable;
                            if (ReadFile(FileNames[i]))
                            {
                                ttpi.Content = "Подходит для экспорта данных";
                                img.Source = BitmapOpenFile;
                            }
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

        private void DeleteFilesToOpen(string[] FileNames)
        {
            for (int i = 0; i < FileNames.Length; i++)
                DeleteFilesToOpen(InputFileName.IndexOf(FileNames[i]));
        }
        private void DeleteFilesToOpen(string FileName)
        {
            DeleteFilesToOpen(InputFileName.IndexOf(FileName));
        }
        private void DeleteFilesToOpen(int FileIndex)
        {
            InputFileName.RemoveAt(FileIndex);
            IDFs.RemoveAt(FileIndex);
            listBoxInputFiles.Items.RemoveAt(FileIndex);
            if (listBoxInputFiles.Items.Count > FileIndex)
            {
                listBoxInputFiles.SelectedIndex = FileIndex;
            }
            else
            {
                if (listBoxInputFiles.Items.Count > 0)
                {
                    listBoxInputFiles.SelectedIndex = listBoxInputFiles.Items.Count - 1;
                }
                else
                    dataGridExport.ItemsSource = null;
            }
            listBoxInputFiles.UpdateLayout();
            dataGridExport.UpdateLayout();
        }

        private void buttonOpenFile_Click(object sender, RoutedEventArgs e)
        {
            string[] dataString = textBoxFileName.Text.Split('|');
            AddFilesToOpen(dataString);
            if (listBoxInputFiles.Items.Count > 0)
            {
                listBoxInputFiles.SelectedIndex = listBoxInputFiles.Items.Count - 1;
            }
        }

        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
            if (checkBoxFindDublicate.IsChecked.Value)
            {
                FindDublicateRecord();
            }

            MainWindow home = this.Owner as MainWindow;
            for (int i = 0; i < IDFs.Count; i++)
            {
                for (int j = 0; j < IDFs[i].InputDataFileRows.Count; j++)
                {
                    home.AddNewItem(IDFs[i].InputDataFileRows[j]);
                }
            }
            home.StartListTeacher();
            Close();
        }
        private void buttonFindDublicates_Click(object sender, RoutedEventArgs e)
        {
            FindDublicateRecord();
        }

        public void UpdateView()
        {
            string pathB = Properties.Settings.Default.PathToGlobal + "\\" + Properties.Settings.Default.GlobalMarker;
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению, на данный момент обновление невозможно - другой пользователь обновляет общий файл.\nПопробуйте еще раз чуть позже");
            }
            else
            {
                string pathA = Properties.Settings.Default.PathToGlobalData;
                string pathC = Directory.GetCurrentDirectory() + "\\" + "View_" + Properties.Settings.Default.GlobalData;
                if (File.Exists(pathC))
                {
                    FileInfo localdata = new FileInfo(pathC);
                    FileInfo globaldata = new FileInfo(pathA);
                    if (globaldata.LastWriteTime.ToLocalTime() > localdata.LastWriteTime.ToLocalTime())
                    {
                        localdata.IsReadOnly = false;
                        File.Delete(pathC);
                        File.Copy(pathA, pathC);
                        localdata.IsReadOnly = true;
                    }
                }
                else
                {
                    File.Copy(pathA, pathC);
                    FileInfo localdata = new FileInfo(pathC);
                    localdata.IsReadOnly = true;

                }

                int i;
                String connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathC + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                switch (pathC.Substring(pathC.LastIndexOf('.')))
                {
                    case ".xls":
                        connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathC + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                        break;
                    case ".xlsx":
                        connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathC + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
                        break;
                    default:
                        MessageBox.Show("Ошибка неизвестного формата файла " + pathC.Substring(pathC.LastIndexOf('.')), "Ошибка расширения", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                        break;
                }
                OleDbConnection con = new OleDbConnection(connection);
                DataTable dtExcelSchema;
                con.Open();
                dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                //con.Close();
                DataSet ds = new DataSet();

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                String Command = "Select * from [" + SheetName + "A15:H]";
                //String Command = "Select * from [Лист1$A15:H]";

                OleDbCommand cmd = new OleDbCommand(Command, con);
                OleDbDataAdapter db = new OleDbDataAdapter(cmd);
                DataTable dt_input = new DataTable();
                db.Fill(dt_input);

                for (i = 0; i < dt_input.Rows.Count; i++)
                {
                    if (dt_input.Rows[i].ItemArray.GetValue(2).ToString().Length == 0 && dt_input.Rows[i].ItemArray.GetValue(3).ToString().Length == 0)
                    {
                        dt_input.Rows[i].Delete();
                        //i -= 1;
                    }
                }
                dt_input.AcceptChanges();
                SvodData.Clear();
                for (i = 0; i < dt_input.Rows.Count; i++)
                {
                    SvodData.Add(new DataTableRow(
                            dt_input.Rows[i].ItemArray.GetValue(1).ToString().Length > 0 ? (dt_input.Rows[i].ItemArray.GetValue(1).ToString().LastIndexOf(" ") > 0 ? dt_input.Rows[i].ItemArray.GetValue(1).ToString().Substring(0, (dt_input.Rows[i].ItemArray.GetValue(1).ToString().IndexOf(" "))) : dt_input.Rows[i].ItemArray.GetValue(1).ToString()) : "",
                            dt_input.Rows[i].ItemArray.GetValue(2).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(2).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(3).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(3).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(4).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(4).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(5).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(5).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(6).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(6).ToString() : ""
                            ));
                }
                //dataGridViewFast.ItemsSource = dt_input.AsDataView();
                cmd.Dispose();
                con.Close();
                con.Dispose();
                //exApp.Quit();
            }
        }


        private void FindDublicateRecord()
        {
            UpdateView();
            for(int i=0;i<IDFs.Count;i++)
            {
                for (int j = 0; j < IDFs[i].InputDataFileRows.Count; j++)
                {
                    for(int l=0;l<SvodData.Count;l++)
                    {
                        if (IDFs[i].InputDataFileRows[j].Intersection(SvodData[l]))
                        {
                            /*if(
                                MessageBox.Show(
                                    "Пересекаются записи \n" +
                                    "из файла\n" +
                                    InputFileName[i] +
                                    ":\n" +
                                    IDFs[i].InputDataFileRows[j].Date + " " + IDFs[i].InputDataFileRows[j].Time + " " + IDFs[i].InputDataFileRows[j].Teacher + " " +
                                    "\n и \n" +
                                    "из общего сводного файла:\n" +
                                    SvodData[l].Date + " " + SvodData[l].Time + " " + SvodData[l].Teacher + " " +
                                    "\n" +
                                    "ОК - продолжить проверку,\nОтмена - прекратить проверку."
                                    , "Обнаружены наложения занятий", MessageBoxButton.OKCancel, MessageBoxImage.Warning)
                                ==MessageBoxResult.Cancel
                                )
                            {
                                return;
                            }
                            */
                            ERLS.Add(new ErrorRowInIDFsAndLinkToRowOnSvod
                                {
                                    ErrorFile = i,
                                    ErrorRow = j,
                                    LinkRow = l
                                }
                            );
                        }
                    }
                }
            }
            for (int i = 0; i < IDFs.Count; i++)
            {
                for (int j = 0; j < IDFs[i].InputDataFileRows.Count; j++)
                {
                    for (int k = i; k < IDFs.Count; k++)
                    {
                        int m = 0;
                        if (k == i)
                            m = j + 1;
                        else
                            m = 0;
                        for (; m < IDFs[k].InputDataFileRows.Count; m++)
                        {
                            if (IDFs[i].InputDataFileRows[j].Intersection(IDFs[k].InputDataFileRows[m]))
                            {
                                /*
                                if(
                                    MessageBox.Show(
                                        "Пересекаются записи \n" +
                                        "из файла\n" +
                                        InputFileName[i] +
                                        ":\n" +
                                        IDFs[i].InputDataFileRows[j].Date + " " + IDFs[i].InputDataFileRows[j].Time + " " + IDFs[i].InputDataFileRows[j].Teacher + " " +
                                        "\n и \n" +
                                        "из файла\n" +
                                        InputFileName[k] +
                                        ":\n" +
                                        IDFs[k].InputDataFileRows[m].Date + " " + IDFs[k].InputDataFileRows[m].Time + " " + IDFs[k].InputDataFileRows[m].Teacher + " " +
                                        "\n" +
                                        "ОК - продолжить проверку,\nОтмена - прекратить проверку."
                                        , "Обнаружены наложения занятий", MessageBoxButton.OKCancel, MessageBoxImage.Warning)
                                    ==MessageBoxResult.Cancel
                                    )
                                {
                                    return;
                                }
                                */
                                ERLR.Add(new ErrorRowInIDFsAndLinkToOtherRow
                                    {
                                        ErrorFile = i,
                                        ErrorRow = j,
                                        LinkFile = k,
                                        LinkRow = m
                                    }
                                );
                            }
                            else
                            { 
                                
                            }
                        }
                    }
                }
            }
            dataGridExport.UpdateLayout();
        }
        private void buttonDeleteFile_Click(object sender, RoutedEventArgs e)
        {
            if (listBoxInputFiles.SelectedItem != null)
            {
                int Ind = listBoxInputFiles.SelectedIndex;
                DeleteFilesToOpen(Ind);
            }
        }


        private bool ReadFile(string FileName)
        {
            InputDataFile tempIDF = new InputDataFile();
            if (tempIDF.OpenFile(FileName))
            {
                IDFs.Add(new InputDataFile(FileName));
                int AllNoneCreatedRecords = 0;
                for (int i = 0; i < IDFs[IDFs.Count - 1].InputDataFileRows.Count; i++)
                {
                    if (TimeTemplate.IndexOf(IDFs[IDFs.Count - 1].InputDataFileRows[i].Time) < 0)
                    {
                        IDFs[IDFs.Count - 1].InputDataFileRows.RemoveAt(i);
                        i -= 1;
                    }
                    else
                    {
                        if (TeacherTemplate.IndexOf(IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher) < 0)
                        {
                            if (NoneTeacherTemplate.IndexOf(IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher) < 0)
                            {
                                if (MessageBox.Show("Обнаруженный не записанный ранее в общий файл расписания преподаватель:\n" +
                                    IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher +
                                    "\nВы хотите добавить его в общий файл?\n(Если ответите \"Нет\", то записи с этим преподавателем будут пропущены.)"
                                    , "Найден незарегистрированный преподаватель", MessageBoxButton.YesNo, MessageBoxImage.Warning
                                    ) == MessageBoxResult.Yes
                                    )
                                {
                                    TeacherTemplate.Add(IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher);
                                    NoneTeacher bYTT = new NoneTeacher();
                                    bYTT.Name = TeacherTemplate[TeacherTemplate.Count - 1];
                                    YTT.Add(bYTT);                                
                                    if (LVEWY.dataGrid != null)
                                        if (LVEWY.dataGrid.ItemsSource != null)
                                        {
                                            CollectionViewSource.GetDefaultView(LVEWY.dataGrid.ItemsSource).Refresh();
                                            LVEWY.dataGrid.UpdateLayout();
                                        }
                                    string path = @".\ListTeacher.dat";
                                    if (File.Exists(path))
                                    {
                                        File.AppendAllText(path, "\n" + IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher);
                                    }
                                    else
                                    {
                                        File.WriteAllText(path, IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher);
                                    }
                                }
                                else
                                {
                                    string path = @".\ListNoneTeacher.dat";
                                    if (File.Exists(path))
                                    {
                                        File.AppendAllText(path, IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher + "\n");
                                    }
                                    else
                                    {
                                        File.WriteAllText(path, IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher + "\n");
                                    }
                                    NoneTeacherTemplate.Add(IDFs[IDFs.Count - 1].InputDataFileRows[i].Teacher);
                                    
                                        NoneTeacher bNTT = new NoneTeacher();
                                        bNTT.Name = NoneTeacherTemplate[NoneTeacherTemplate.Count-1];
                                        NTT.Add(bNTT);
                                    CollectionViewSource.GetDefaultView(LVEW.dataGrid.ItemsSource).Refresh();

                                    IDFs[IDFs.Count - 1].InputDataFileRows.RemoveAt(i);
                                    i -= 1;
                                    AllNoneCreatedRecords += 1;
                                }
                            }
                            else
                            {
                                IDFs[IDFs.Count - 1].InputDataFileRows.RemoveAt(i);
                                i -= 1;
                                AllNoneCreatedRecords += 1;
                            }
                        }

                    }
                }
                if(AllNoneCreatedRecords >0)
                {
                    MessageBox.Show("Из файла \n" + FileName + "\nне удалось добавить из-за ошибок ввода или ФИО преподавателя " + AllNoneCreatedRecords.ToString() + " запис(ь/и/ей)", "Неудачные записи", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (IDFs[IDFs.Count - 1].InputDataFileRows.Count <= 0)
                {
                    IDFs[IDFs.Count - 1].InputDataFileRows.Clear();
                    return false;
                }
                return true;
            }
            else
            {
                IDFs.Add(new InputDataFile());
                return false;
            }


        }
        private void listBoxInputFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int Ind = listBoxInputFiles.SelectedIndex;
            if (Ind >= 0)
            {
                if (IDFs[Ind] != null)
                {
                    dataGridExport.ItemsSource = IDFs[Ind].InputDataFileRows;
                    StatusStringCountRecordFile.Content = IDFs[Ind].InputDataFileRows.Count.ToString();
                }
                else
                {
                    dataGridExport.ItemsSource = null;
                    StatusStringCountRecordFile.Content = "0";
                }
            }
            else
            {
                dataGridExport.ItemsSource = null;
                StatusStringCountRecordFile.Content = "0";
            }
            //StatusStringCountRecordFile.Content = (dataGridExport.ItemsSource as List<DataTableRow>).Count.ToString();
            if (dataGridExport.Columns.Count > 0)
            {
                dataGridExport.Columns[0].Header = "Дата проведения";
                dataGridExport.Columns[1].Header = "Время проведения";
                dataGridExport.Columns[2].Header = "Преподаватель";
                dataGridExport.Columns[3].Header = "Номер группы";
                dataGridExport.Columns[4].Header = "Категория слушателей";
                dataGridExport.Columns[5].Header = "Место проведения";
            }
            dataGridExport.UpdateLayout();

            if (dataGridExport.SelectedIndex >= 0)
            {
                buttonEditInputHot.IsEnabled = true;
                buttonDeleteHot.IsEnabled = true;
            }
            else
            {
                buttonDeleteHot.IsEnabled = false;
                buttonEditInputHot.IsEnabled = false;
            }
        }

        private void listBoxInputFilesItemsChanges(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (listBoxInputFiles.Items.Count == 0)
            {
                InputFileName.Clear();
                IDFsIndex.Clear();
            }
            countAllRecords = 0;
            for (int i = 0; i < IDFs.Count; i++)
            {
                //if(IDFs[i].InputDataFileRows!=null)
                countAllRecords += (ulong)IDFs[i].InputDataFileRows.Count;
            }
            StatusStringCountRecordAllFile.Content = countAllRecords.ToString();
            if (dataGridExport.SelectedIndex >= 0)
            {
                buttonEditInputHot.IsEnabled = true;
                buttonDeleteHot.IsEnabled = true;
            }
            else
            {
                buttonDeleteHot.IsEnabled = false;
                buttonEditInputHot.IsEnabled = false;
            }
        }

        private void buttonUpdateTimeTemplates_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            UpdateLayout();
            /*
            SingleInput SItemp = new SingleInput();
            SItemp.Owner = this;
            SItemp.exApp = (Owner as MainWindow).exApp;
            SItemp.UpdateListTimes();
            SItemp.Owner = null;
            */
            DataWork.UpdateListTimes((Owner as MainWindow).exApp);
            StartListTimes();

            this.Effect = null;
            UpdateLayout();
        }

        private void buttonUpdateTeacherTemplates_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            UpdateLayout();
            /*
            SingleInput SItemp = new SingleInput();
            SItemp.Owner = this;
            SItemp.exApp = (Owner as MainWindow).exApp;
            SItemp.UpdateListTeacher();
            SItemp.Owner = null;
            */
            DataWork.UpdateTeachersList((Owner as MainWindow).exApp);
            StartListTeacher();
            this.Effect = null;
            UpdateLayout();
        }

        private void dataGridExport_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                if (dataGridExport.Items.Count > 0 && dataGridExport.SelectedIndex >= 0)
                    DeleteItem(dataGridExport.SelectedIndex);
            }
        }
        public void DeleteItem(int RowIndex)
        {
            if (RowIndex >= 0 && RowIndex < dataGridExport.Items.Count)
            {
                int Ind = listBoxInputFiles.SelectedIndex;
                if (Ind >= 0)
                {
                    if (MessageBox.Show("Вы действительно хотите удалить из экспортируемых данных файла\n" + InputFileName[Ind] + "\n запись\n" + IDFs[Ind].InputDataFileRows[RowIndex].Date + " " + IDFs[Ind].InputDataFileRows[RowIndex].Time + " " + IDFs[Ind].InputDataFileRows[RowIndex].Teacher + "\n?", "Удаление элемента из экспорта", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                    {

                        if (IDFs[Ind] != null)
                        {
                            IDFs[Ind].InputDataFileRows.Remove(IDFs[Ind].InputDataFileRows[RowIndex]);
                        }
                        else
                        {
                            dataGridExport.ItemsSource = null;
                            StatusStringCountRecordFile.Content = "0";
                        }
                    }
                    CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
                    if (IDFs[Ind].InputDataFileRows.Count < 1)
                    {
                        DeleteFilesToOpen(Ind);
                        buttonDeleteHot.IsEnabled = false;
                        buttonEditInputHot.IsEnabled = false;
                    }
                }
            }
            else
                MessageBox.Show("Ошибка удаления элемента");
        }

        private void buttonDeleteHot_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridExport.Items.Count > 0 && dataGridExport.SelectedIndex >= 0)
                DeleteItem(dataGridExport.SelectedIndex);
        }

        private void DataGridCell_PreviewSelected(object sender, RoutedEventArgs e)
        {
            if (dataGridExport.SelectedIndex >= 0)
            {
                buttonEditInputHot.IsEnabled = true;
                buttonDeleteHot.IsEnabled = true;
            }
            else
            {
                buttonDeleteHot.IsEnabled = false;
                buttonEditInputHot.IsEnabled = false;
            }
        }

        private void DataGridCell_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ClickToAddRow = false;
            DataGridCell cell = sender as DataGridCell;
            ChangeDataGrid();
        }

        private void ChangeDataGrid()
        {
            int Ind = listBoxInputFiles.SelectedIndex;
            int RowIndex = dataGridExport.SelectedIndex;
            if (dataGridExport.SelectedIndex >= 0)
            {
                {
                    int SI = dataGridExport.SelectedIndex;
                    SingleInput f = new SingleInput();
                    f.Owner = this;
                    f.exApp = (Owner as MainWindow).exApp;
                    f.Top = this.Top + 50;
                    f.Left = this.Left + 50;
                    f.RowIndex = dataGridExport.SelectedIndex;
                    if (dataGridExport.CurrentColumn != null)
                    {
                        switch (dataGridExport.CurrentColumn.DisplayIndex)
                        {
                            case 0:
                                f.DatePicker_Date.Focus();
                                break;
                            case 1:
                                f.MaskedTextBoxStartTime.Focus();
                                break;
                            case 2:
                                f.comboBoxTeacher.Focus();
                                break;
                            case 3:
                                f.textboxGroup.Focus();
                                break;
                            case 4:
                                f.textBoxCategory.Focus();
                                break;
                            case 5:
                                f.textBoxPlace.Focus();
                                break;
                            default:
                                f.ButtonCancel.Focus();
                                break;
                        }
                    }
                    else
                    {
                        f.DatePicker_Date.Focus();
                    }

                    f.DatePicker_Date.Text = IDFs[Ind].InputDataFileRows[RowIndex].Date;
                    f.comboBoxTeacher.Text = IDFs[Ind].InputDataFileRows[RowIndex].Teacher;

                    f.MaskedTextBoxStartTime.Text = IDFs[Ind].InputDataFileRows[RowIndex].Time.Substring(0, 5).Replace('.', ':');
                    if (f.MaskedTextBoxStartTime.Text[0] == '_')
                    {
                        f.MaskedTextBoxStartTime.Text = "0" + IDFs[Ind].InputDataFileRows[RowIndex].Time.Substring(0, 4).Replace('.', ':');
                    }
                    f.MaskedTextBoxEndTime.Text = IDFs[Ind].InputDataFileRows[RowIndex].Time.Substring(IDFs[Ind].InputDataFileRows[RowIndex].Time.Length - 5, 5).Replace('.', ':');
                    if (f.MaskedTextBoxEndTime.Text[0] == '_')
                    {
                        f.MaskedTextBoxEndTime.Text = "0" + IDFs[Ind].InputDataFileRows[RowIndex].Time.Substring(IDFs[Ind].InputDataFileRows[RowIndex].Time.Length - 4, 4).Replace('.', ':');
                    }
                    f.comboBoxTeacher.SelectedIndex = f.comboBoxTeacher.Items.IndexOf(IDFs[Ind].InputDataFileRows[RowIndex].Teacher);
                    f.textboxGroup.Text = IDFs[Ind].InputDataFileRows[RowIndex].Group;
                    f.textBoxCategory.Text = IDFs[Ind].InputDataFileRows[RowIndex].Category;
                    f.textBoxPlace.Text = IDFs[Ind].InputDataFileRows[RowIndex].Place;
                    f.Title = "Редактирование записи \"" + IDFs[Ind].InputDataFileRows[RowIndex].Date + " " + IDFs[Ind].InputDataFileRows[RowIndex].Time + " " + IDFs[Ind].InputDataFileRows[RowIndex].Teacher + "\"";
                    f.ButtonWriteAndContinue.IsEnabled = false;
                    f.ButtonWriteAndContinue.Visibility = Visibility.Collapsed;
                    f.ButtonWriteAndStop.Content = "Внести изменения";
                    f.ButtonWriteAndStop.HorizontalAlignment = HorizontalAlignment.Left;
                    f.ButtonWriteAndStop.Margin = new Thickness(10, 0, 0, 10);
                    f.ShowDialog();
                }
            }
        }

        private void MenuItemEditInput_Click(object sender, RoutedEventArgs e)
        {
            ChangeDataGrid();
        }

        public void EditItem(int RowIndex, DataTableRow newDTR)
        {
            int Ind = listBoxInputFiles.SelectedIndex;
            if (Ind >= 0)
            {
                IDFs[Ind].InputDataFileRows[RowIndex] = newDTR;
            }
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
        }

        private void buttonUpdateNot_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите обнулить список игнорируемых преподавателей при добавлении новых записей?", "Сброс Списка игнорируемых преподавателей", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                string path = @".\ListNoneTeacher.dat";
                NoneTeacherTemplate.Clear();
                if (File.Exists(path))
                {
                    File.WriteAllText(path, "");
                }
                NTT.Clear();
                CollectionViewSource.GetDefaultView(LVEW.dataGrid.ItemsSource).Refresh();
            }
        }

        public void buttonListNotTeachers_Click(object sender, RoutedEventArgs e)
        {

            //LVEW.SetBinding();
            //labelTech.Content = LVEW.dataGrid.ItemsSource.GetEnumerator().ToString();
            //for (int i = 0; i < NoneTeacherTemplate.Count; i++) MessageBox.Show(NoneTeacherTemplate[i]);
            //MessageBox.Show(LVEW.dataGrid.Items[0].ToString());
            LVEW.Show();
            this.Activate();
            /*
             * LVEW.Close();

            LVEW = new ListViewEditWindow(NoneTeacherTemplate);
            Binding bind = new Binding();
            bind.Source = NoneTeacherTemplate;
            bind.Path = new PropertyPath(".");
            //bind.XPath = ".";
            bind.Mode = BindingMode.TwoWay;
            //LVEW.dataGrid.ItemsSource = NoneTeacherTemplate;
            //LVEW.dataGrid.SetBinding(ItemsControl.ItemsSourceProperty, bind);
            //LVEW.dataGrid.Columns[0].Header = "ФИО";
            LVEW.Show();
            */
        }
        
        public void SaveDataListViewEditWindow()
        {
            NoneTeacherTemplate.Clear();
            for (int i = 0; i < NTT.Count; i++)
            {
                NoneTeacherTemplate.Add(NTT[i].Name);
            }
        }

        private void buttonListTeacher_Click(object sender, RoutedEventArgs e)
        {
            LVEWY.Show();
        }

        private void dataGridExport_LayoutUpdated(object sender, EventArgs e)
        {
            if(ERLR.Count>0||ERLS.Count>0)
            {
                for (int i= 0;i< dataGridExport.Items.Count;i++)
                {
                    DataGridRow row = (DataGridRow)dataGridExport.ItemContainerGenerator.ContainerFromIndex(i);
                    SolidColorBrush brush = new SolidColorBrush(Colors.LawnGreen);
                    row.Background = brush;
                }
                for(int i=0;i<ERLS.Count;i++)
                {
                    if(ERLS[i].ErrorFile==listBoxInputFiles.SelectedIndex)
                    {

                    }
                }
                for (int i = 0; i < ERLR.Count; i++)
                {

                }
            }
        }
    }
}
