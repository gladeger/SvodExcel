using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.IO.Compression;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using ExcelLibrary;
using System.Collections.Specialized;


namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
        public bool AdminMode, ConnectMode, NoneSave;
        public List<DataTableRow> DTR = new List<DataTableRow>();
        public List<DataViewTableRow> vDTR = new List<DataViewTableRow>();
        public List<DataViewFastTableRow> vfDTR = new List<DataViewFastTableRow>();
        private bool ClickToAddRow = true;
        private struct markerActionData
        {
            public int IndexData { get; set; }
            public bool Action { get; set; }
        }
        private List<markerActionData> MAD = new List<markerActionData>();
        public MainWindow()
        {
            InitializeComponent();
            DTR.Clear();
            vDTR.Clear();
            vfDTR.Clear();
            MAD.Clear();

            AdminMode = false;
            ConnectMode = false;
            NoneSave = false;
            dataGridExport.ItemsSource = DTR;
            dataGridView.ItemsSource = vDTR;
            dataGridViewFast.ItemsSource = vfDTR;
            buttonAdminOff.IsEnabled = false;
            buttonAdminOff.Visibility = Visibility.Collapsed;
            MenuItemAdminOff.IsEnabled = false;
            MenuItemAdminOff.Visibility = Visibility.Collapsed;
            MenuItemOptions.IsEnabled = false;
            ViewEditTab.Visibility = Visibility.Collapsed;

            ((INotifyCollectionChanged)dataGridExport.Items).CollectionChanged += dataGridExportItemsChanges;

            AdminModeActive();//вкл/выкл режим админа
        }
        private void SvodExcel_Loaded(object sender, RoutedEventArgs e)
        {
            DTR.Clear();
            vDTR.Clear();
            vfDTR.Clear();

            // example data
            //AddNewItem(new DataTableRow("06.11.2019", "10:00-16:40", "Пронина Л.Н.", "","-----","-----"));
            //AddNewItem(new DataTableRow("07.11.2019", "12:00-18:40", "Радюхина Е.И.", "", "#######", "*?!~%$#"));
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            CollectionViewSource.GetDefaultView(dataGridViewFast.ItemsSource).Refresh();
            //----exmpla data

            ClearHang();
            buttonDebug.Visibility = Visibility.Collapsed;
        }
        private void SvodExcel_Closed(object sender, EventArgs e)
        {
            ClearHang();
            exApp.Quit();
            exApp = null;
            System.Windows.Application.Current.Shutdown();
            GC.Collect();
        }

        private void MenuItem_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            AboutBox1 f = new AboutBox1();
            f.ShowDialog();
        }

        private void openSingleInput()
        {
            if(tabControl.SelectedIndex==0)
            {
                SingleInput f = new SingleInput();
                f.Owner = this;
                f.exApp = exApp;
                f.Top = this.Top + 50;
                f.Left = this.Left + 50;
                f.RowIndex = -1;
                f.ShowDialog();
                //f.Show();

                CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            }
        }
        private void MenuItemSingleInput_Click(object sender, RoutedEventArgs e)
        {
            openSingleInput();
        }

        public void AddNewItem(DataTableRow newDTR)
        {
            DTR.Add(newDTR);
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            buttonExport.IsEnabled = true;
            buttonExportHot.IsEnabled = true;
            //buttonDeleteHot.IsEnabled = true;
        }


        private void DataGridCell_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ClickToAddRow = false;
            DataGridCell cell = sender as DataGridCell;
            ChangeDataGrid();
        }
        private void ChangeDataGrid()
        {
            SingleInput f;
            int SI;
            switch (tabControl.SelectedIndex)
            {
                case 0:
                    {
                        SI =DTR.IndexOf(dataGridExport.SelectedItem as DataTableRow);
                        f = new SingleInput();
                        f.Owner = this;
                        f.exApp = exApp;
                        f.Top = this.Top + 50;
                        f.Left = this.Left + 50;
                        f.RowIndex = SI;
                        if(dataGridExport.CurrentColumn!=null)
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
                        
                        f.DatePicker_Date.Text = DTR[SI].Date;
                        f.comboBoxTeacher.Text = DTR[SI].Teacher;

                        f.MaskedTextBoxStartTime.Text = DTR[SI].Time.Substring(0, 5).Replace('.', ':');
                        if (f.MaskedTextBoxStartTime.Text[0] == '_')
                        {
                            f.MaskedTextBoxStartTime.Text = "0" + DTR[SI].Time.Substring(0, 4).Replace('.', ':');
                        }
                        f.MaskedTextBoxEndTime.Text = DTR[SI].Time.Substring(DTR[SI].Time.Length - 5, 5).Replace('.', ':');
                        if (f.MaskedTextBoxEndTime.Text[0] == '_')
                        {
                            f.MaskedTextBoxEndTime.Text = "0" + DTR[SI].Time.Substring(DTR[SI].Time.Length - 4, 4).Replace('.', ':');
                        }
                        f.comboBoxTeacher.SelectedIndex = f.comboBoxTeacher.Items.IndexOf(DTR[SI].Teacher);
                        f.textboxGroup.Text = DTR[SI].Group;
                        f.textBoxCategory.Text = DTR[SI].Category;
                        f.textBoxPlace.Text = DTR[SI].Place;
                        f.Title = "Редактирование записи \"" + DTR[SI].Date + " " + DTR[SI].Time + " " + DTR[SI].Teacher + "\"";
                        f.ButtonWriteAndContinue.IsEnabled = false;
                        f.ButtonWriteAndContinue.Visibility = Visibility.Collapsed;
                        f.ButtonWriteAndStop.Content = "Внести изменения";
                        f.ButtonWriteAndStop.HorizontalAlignment = HorizontalAlignment.Left;
                        f.ButtonWriteAndStop.Margin = new Thickness(10, 0, 0, 10);
                        f.ShowDialog();
                    }                    
                    break;
                case 3:
                    SI = vDTR.IndexOf(dataGridViewEdit.SelectedItem as DataViewTableRow);
                    f = new SingleInput();
                    f.Owner = this;
                    f.exApp = exApp;
                    f.Top = this.Top + 50;
                    f.Left = this.Left + 50;
                    f.RowIndex = SI;
                    if (dataGridExport.CurrentColumn != null)
                    {
                        switch (dataGridViewEdit.CurrentColumn.DisplayIndex)
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

                    f.DatePicker_Date.Text = vDTR[SI].Date;
                    f.comboBoxTeacher.Text = vDTR[SI].Teacher;

                    f.MaskedTextBoxStartTime.Text = vDTR[SI].Time.Substring(0, 5).Replace('.', ':');
                    if (f.MaskedTextBoxStartTime.Text[0] == '_')
                    {
                        f.MaskedTextBoxStartTime.Text = "0" + vDTR[SI].Time.Substring(0, 4).Replace('.', ':');
                    }
                    f.MaskedTextBoxEndTime.Text = vDTR[SI].Time.Substring(vDTR[SI].Time.Length - 5, 5).Replace('.', ':');
                    if (f.MaskedTextBoxEndTime.Text[0] == '_')
                    {
                        f.MaskedTextBoxEndTime.Text = "0" + vDTR[SI].Time.Substring(vDTR[SI].Time.Length - 4, 4).Replace('.', ':');
                    }
                    f.comboBoxTeacher.SelectedIndex = f.comboBoxTeacher.Items.IndexOf(vDTR[SI].Teacher);
                    f.textboxGroup.Text = vDTR[SI].Group;
                    f.textBoxCategory.Text = vDTR[SI].Category;
                    f.textBoxPlace.Text = vDTR[SI].Place;
                    f.Title = "Редактирование записи \"" + vDTR[SI].Date + " " + vDTR[SI].Time + " " + vDTR[SI].Teacher + "\"";
                    f.ButtonWriteAndContinue.IsEnabled = false;
                    f.ButtonWriteAndContinue.Visibility = Visibility.Collapsed;
                    f.ButtonWriteAndStop.Content = "Внести изменения";
                    f.ButtonWriteAndStop.HorizontalAlignment = HorizontalAlignment.Left;
                    f.ButtonWriteAndStop.Margin = new Thickness(10, 0, 0, 10);
                    f.ShowDialog();
                    break;
                default:
                    break;
            }
            
        }
        public void EditItem(int RowIndex,DataTableRow newDTR)
        {
            DTR[RowIndex] = newDTR;
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
        }
        public void EditItemEdition(int RowIndex, DataViewTableRow newDTR)
        {
            NoneSave = true;
            int bufRowIndex = vDTR.IndexOf(dataGridViewEdit.Items[RowIndex] as DataViewTableRow);
            vDTR[bufRowIndex] = newDTR;
            markerActionData bufmad= new markerActionData();
            bufmad.IndexData = bufRowIndex;
            bufmad.Action = true;
            MAD.Add(bufmad);
            CollectionViewSource.GetDefaultView(dataGridViewEdit.ItemsSource).Refresh();
            
        }
        public void DeleteItem(int RowIndex)
        {
            if (RowIndex >= 0 && RowIndex < DTR.Count)
            {
                if(MessageBox.Show("Вы действительно хотите удалить из экспортируемых данных запись\n"+ DTR[RowIndex].Date + " " + DTR[RowIndex].Time + " " + DTR[RowIndex].Teacher+"\n?","Удаление элемента из экспорта",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No) ==MessageBoxResult.Yes)
                {
                    DTR.Remove(dataGridExport.Items[RowIndex] as DataTableRow);
                }
                CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
                if(DTR.Count<1)
                {
                    buttonDeleteHot.IsEnabled = false;
                    buttonEditInputHot.IsEnabled = false;
                    buttonExport.IsEnabled = false;
                    buttonExportHot.IsEnabled = false;
                }
            }
            else
                MessageBox.Show("Ошибка удаления элемента");
        }
        public void DeleteItemEdition(int RowIndex)
        {
            if (RowIndex >= 0 && RowIndex < vDTR.Count)
            {
                if (MessageBox.Show("Вы действительно хотите удалить из экспортируемых данных запись\n" + vDTR[RowIndex].Date + " " + vDTR[RowIndex].Time + " " + vDTR[RowIndex].Teacher + "\n?", "Удаление элемента из экспорта", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                {
                    NoneSave = true;
                    int bufRowIndex = vDTR.IndexOf(dataGridViewEdit.Items[RowIndex] as DataViewTableRow);
                    vDTR.RemoveAt(bufRowIndex);
                    markerActionData bufmad = new markerActionData();
                    bufmad.IndexData = bufRowIndex;
                    bufmad.Action = false;
                    MAD.Add(bufmad);
                    labelTech.Content += "\n" + bufRowIndex.ToString();
                }                    
                CollectionViewSource.GetDefaultView(dataGridViewEdit.ItemsSource).Refresh();
                if (vDTR.Count < 1)
                {
                    buttonDeleteHot.IsEnabled = false;
                    buttonEditInputHot.IsEnabled = false;
                    buttonExport.IsEnabled = false;
                    buttonExportHot.IsEnabled = false;
                }
            }
            else
                MessageBox.Show("Ошибка удаления элемента");
        }

        private void buttonDeleteHot_Click(object sender, RoutedEventArgs e)
        {
            switch (tabControl.SelectedIndex)
            {
                case 0:
                    DeleteItem(dataGridExport.SelectedIndex);
                    break;
                case 3:
                    DeleteItemEdition(dataGridViewEdit.SelectedIndex);
                    break;
                default:
                    break;
            }            
        }

        private void Export_Click()
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            UpdateLayout();
            if (MessageBox.Show("Вы действительно хотите добавить в общий файл все созданные ранее записи?\nВсего записей для экспорта: " + DTR.Count, "Экспот данных в общий файл", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                double This_TH2 = this.Top + this.Height / 2.0;
                double This_LW2 = this.Left + this.Width / 2.0;
                /*
                Thread newWindowThread = new Thread(new ThreadStart(() =>
                {
                    SvodExcel.ProgressBar PB = new SvodExcel.ProgressBar();
                    PB.Top = This_TH2 - PB.Height / 2.0;
                    PB.Left = This_LW2 - PB.Width / 2.0;
                    PB.Topmost = false;
                    PB.ShowDialog();
                    System.Windows.Threading.Dispatcher.Run();
                }));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                */
                ExportData();
                //newWindowThread.Abort();
            }
            this.Effect = null;
            UpdateLayout();
        }
        private void buttonExport_Click(object sender, RoutedEventArgs e)
        {
            Export_Click();  
        }
        private void buttonExportHot_Click(object sender, RoutedEventArgs e)
        {
            Export_Click();
        }

        public void ExportData()//вставляем в общий файл данные
        {           
            string pathB = Properties.Settings.Default.PathToGlobal+"\\"+Properties.Settings.Default.GlobalMarker;
            ClearHang();
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению, на данный момент экспорт невозможен - другой пользователь уже начал оновлять общий файл!\nПопробуйте еще раз чуть позже");
            }
            else
            {
                string timeforname= DateTime.Now.ToString().Replace(':', '_');
                string pathC = Directory.GetCurrentDirectory() + "\\" + Properties.Settings.Default.GlobalData;
                if (File.Exists(pathC))
                {
                    try { File.Delete(pathC); }
                    catch
                    {
                        MessageBox.Show("Ошибка обращения к локальной копии сводного документа.\nПерезапустите компьютер");
                        return;
                    }
                    
                }
                StreamWriter sw = File.CreateText(pathB);
                String host = System.Net.Dns.GetHostName();
                System.Net.IPAddress ip = System.Net.Dns.GetHostEntry(host).AddressList[0];
                sw.WriteLine(ip.ToString());
                sw.Close();
                string pathA = Properties.Settings.Default.PathToGlobalData;
                File.Copy(pathA, pathC);
                //var exApp = new Microsoft.Office.Interop.Excel.Application();
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                
                var exBook = exApp.Workbooks.Open(pathC);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                int BlinkEnd = 0;
                var lastcell = ExSheet.Cells.SpecialCells(Type: Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                if (ExSheet.Cells[lastcell.Row, 2].Value != null || ExSheet.Cells[lastcell.Row, 3].Value != null || ExSheet.Cells[lastcell.Row, 4].Value != null || ExSheet.Cells[lastcell.Row, 5].Value != null || ExSheet.Cells[lastcell.Row, 6].Value != null || ExSheet.Cells[lastcell.Row, 7].Value != null)
                    BlinkEnd = 1;
                for (int i= BlinkEnd; i<(DTR.Count+BlinkEnd); i++)
                {
                    ExSheet.Cells[lastcell.Row + i, 2] = DTR[i-BlinkEnd].Date;
                    ExSheet.Cells[lastcell.Row + i, 3] = DTR[i - BlinkEnd].Time;
                    ExSheet.Cells[lastcell.Row + i, 4] = DTR[i - BlinkEnd].Teacher;
                    ExSheet.Cells[lastcell.Row + i, 5] = DTR[i - BlinkEnd].Group;
                    ExSheet.Cells[lastcell.Row + i, 6] = DTR[i - BlinkEnd].Category;
                    ExSheet.Cells[lastcell.Row + i, 7] = DTR[i - BlinkEnd].Place;
                }    
                exBook.Close(true);            
               
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                File.Move(pathA, pathA.Substring(0, pathA.Length - 5) + " " + timeforname + ".xlsx");
                File.Copy(pathC, pathA);
                //File.Replace(pathC,pathA,pathA.Substring(0, pathA.Length-5)+ " "+DateTime.Now.ToString().Replace(':','_')+ ".xlsx");
                File.Delete(pathC);

                //*******************************************************************************************************************************************************************************************************************************
                {
                    string pathD = Directory.GetCurrentDirectory() + "\\" + timeforname + " " + System.Net.Dns.GetHostName() + " " + System.Security.Principal.WindowsIdentity.GetCurrent().Name.Substring(System.Security.Principal.WindowsIdentity.GetCurrent().Name.LastIndexOf('\\') + 1) + "." + Properties.Settings.Default.ExtensionFileNewData;
                    if (File.Exists(pathD))
                    {
                        File.Delete(pathD);
                    }
                    String connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathD + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                    switch (Properties.Settings.Default.ExtensionFileNewData)
                    {
                        case "xls":
                            connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathD + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                            break;
                        case "xlsx":
                            connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathD + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
                            break;
                        default:
                            MessageBox.Show("Ошибка неизвестного формата файла " + Properties.Settings.Default.ExtensionFileNewData, "Ошибка расширения", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                            break;
                    }
                    DataSet ds = new DataSet();
                    DataTable dt = new DataTable("Sheet_1");
                    ds.Tables.Add(dt);
                    dt.Columns.Add("Date", Type.GetType("System.String"));
                    dt.Columns.Add("Time", Type.GetType("System.String"));
                    dt.Columns.Add("Teacher", Type.GetType("System.String"));
                    dt.Columns.Add("Group", Type.GetType("System.String"));
                    dt.Columns.Add("Category", Type.GetType("System.String"));
                    dt.Columns.Add("Place", Type.GetType("System.String"));
                    int i;
                    for (i = 0; i < DTR.Count; i++)
                    {
                        dt.Rows.Add(DTR[i].Date, DTR[i].Time, DTR[i].Teacher, DTR[i].Group, DTR[i].Category, DTR[i].Place);
                    }
                    for (; i <= 100; i++)
                    {
                        dt.Rows.Add("", "", "", "", "", "");
                    }
                    DataSetHelper.CreateWorkbook(pathD, ds);
                    ds.Dispose();
                    dt.Dispose();
                    File.Copy(pathD, Properties.Settings.Default.PathToGlobal+"\\"+pathD.Substring(pathD.LastIndexOf('\\')+1));
                    File.Delete(pathD);
                }               
//*******************************************************************************************************************************************************************************************************************************
                File.Delete(pathB);
                DTR.Clear();
                CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            }
        }
        public void ClearHang()
        {
            string pathB = Properties.Settings.Default.PathToGlobal +"\\"+ Properties.Settings.Default.GlobalMarker;
            if (File.Exists(pathB))
            {
                FileInfo employed = new FileInfo(pathB);
                StreamReader sw = new StreamReader(pathB);
                String host = System.Net.Dns.GetHostName();
                System.Net.IPAddress ip = System.Net.Dns.GetHostEntry(host).AddressList[0];
                string busy_customer = "";
                busy_customer = sw.ReadLine();
                sw.Close();
                if ((busy_customer == ip.ToString())||(DateTime.Now.Subtract(employed.CreationTime.ToLocalTime()).TotalMinutes > Properties.Settings.Default.WaitingInLine))
                {
                    File.Delete(pathB);
                    int i = 0;
                    while (File.Exists(pathB)) {
                        if (i > 10000)
                            break;
                        i++;
                    }
                }
            }
        }

        private void buttonDebug_Click(object sender, RoutedEventArgs e)
        {
        }

        private void tabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            buttonDeleteHot.IsEnabled = false;
            buttonEditInputHot.IsEnabled = false;
            switch (tabControl.SelectedIndex)
            {
                case 0:
                    menu_Hot_Export.IsEnabled = true;
                    menu_Hot_Edit.IsEnabled = true;
                    menu_Hot_View.IsEnabled = false;
                    menu_Hot_ViewFast.IsEnabled = false;
                    MenuInputData.IsEnabled = true;

                    menu_Hot_Export.Visibility = Visibility.Visible;
                    menu_Hot_Edit.Visibility = Visibility.Visible;
                    menu_Hot_View.Visibility = Visibility.Collapsed;
                    menu_Hot_ViewFast.Visibility = Visibility.Collapsed;
                    break;
                case 1:
                    menu_Hot_Export.IsEnabled = false;
                    menu_Hot_Edit.IsEnabled = false;
                    menu_Hot_View.IsEnabled = true;
                    menu_Hot_ViewFast.IsEnabled = true;
                    MenuInputData.IsEnabled = false;

                    menu_Hot_Export.Visibility = Visibility.Collapsed;
                    menu_Hot_Edit.Visibility = Visibility.Collapsed;
                    menu_Hot_View.Visibility = Visibility.Visible;
                    menu_Hot_ViewFast.Visibility = Visibility.Visible;
                    break;
                case 2:
                    menu_Hot_Export.IsEnabled = false;
                    menu_Hot_Edit.IsEnabled = false;
                    menu_Hot_View.IsEnabled = true;
                    menu_Hot_ViewFast.IsEnabled = false;
                    MenuInputData.IsEnabled = false;

                    menu_Hot_Export.Visibility = Visibility.Collapsed;
                    menu_Hot_Edit.Visibility = Visibility.Collapsed;
                    menu_Hot_View.Visibility = Visibility.Visible;
                    menu_Hot_ViewFast.Visibility = Visibility.Collapsed;
                    break;
                case 3:
                    menu_Hot_Export.IsEnabled = false;
                    menu_Hot_Edit.IsEnabled = true;
                    menu_Hot_View.IsEnabled = false;
                    menu_Hot_ViewFast.IsEnabled = false;
                    MenuInputData.IsEnabled = false;

                    menu_Hot_Export.Visibility = Visibility.Collapsed;
                    menu_Hot_Edit.Visibility = Visibility.Visible;
                    menu_Hot_View.Visibility = Visibility.Collapsed;
                    menu_Hot_ViewFast.Visibility = Visibility.Collapsed;
                    break;
                default:
                    menu_Hot_Export.IsEnabled = false;
                    menu_Hot_Edit.IsEnabled = false;
                    menu_Hot_View.IsEnabled = false;
                    menu_Hot_ViewFast.IsEnabled = false;
                    MenuInputData.IsEnabled = false;

                    menu_Hot_Export.Visibility = Visibility.Visible;
                    menu_Hot_Edit.Visibility = Visibility.Visible;
                    menu_Hot_View.Visibility = Visibility.Visible;
                    menu_Hot_ViewFast.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void buttonView_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            UpdateLayout();
            if (MessageBox.Show("Вы действительно хотите скачать и посмотреть данные из общего файла?\nЭто может занять несколько минут.", "Просмотр общих данных",MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                double This_TH2 = this.Top + this.Height / 2.0;
                double This_LW2 = this.Left + this.Width / 2.0;
               /* 
                Thread newWindowThread = new Thread(new ThreadStart(() =>
                {
                    SvodExcel.ProgressBar PB = new SvodExcel.ProgressBar();
                    PB.Top = This_TH2 - PB.Height / 2.0;
                    PB.Left = This_LW2 - PB.Width / 2.0;
                    PB.Topmost = false;
                    PB.ShowDialog();
                    System.Windows.Threading.Dispatcher.Run();
                }));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                */
                switch(tabControl.SelectedIndex)
                {
                    case 1:
                        UpdateViewFast();
                        break;
                    case 2:
                        UpdateView();
                        break;
                    default:
                        break;
                }
                
                //newWindowThread.Abort();
            }           
            this.Effect = null;
            UpdateLayout();

        }
        
        public void UpdateViewFast()//Обновление быстрого просмотра сводной таблицы
        {
            string pathB = Properties.Settings.Default.PathToGlobal +"\\"+ Properties.Settings.Default.GlobalMarker;
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению, на данный момент обновление невозможно - другой пользователь обновляет общий файл.\nПопробуйте еще раз чуть позже");
            }
            else
            {
                string pathA = Properties.Settings.Default.PathToGlobalData;
                string pathFast = Directory.GetCurrentDirectory() + "\\" + Properties.Settings.Default.ViewFast;
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
                        CollectionViewSource.GetDefaultView(dataGridViewFast.ItemsSource).Refresh();
                        if (File.Exists(pathFast))
                        {
                            FileInfo localfastdata = new FileInfo(pathFast);
                            localfastdata.IsReadOnly = false;
                            File.Delete(pathFast);
                        }
                    }
                    else
                    {
                        if (dataGridViewFast.Items.Count > 0)
                            return;
                    }
                }
                else
                {
                    File.Copy(pathA, pathC);
                    FileInfo localdata = new FileInfo(pathC);
                    localdata.IsReadOnly = true;
                    if(File.Exists(pathFast))
                    {
                        FileInfo localfastdata = new FileInfo(pathFast);
                        localfastdata.IsReadOnly =false;
                        File.Delete(pathFast);
                    }
                }

                vfDTR.Clear();
                if (File.Exists(pathFast))
                {
                    String connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathFast + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                    switch (Properties.Settings.Default.ViewFast.Substring(Properties.Settings.Default.ViewFast.LastIndexOf('.')))
                    {
                        case ".xls":
                            connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathFast + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                            break;
                        case ".xlsx":
                            connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathFast + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
                            break;
                        default:
                            MessageBox.Show("Ошибка неизвестного формата файла "+ Properties.Settings.Default.ViewFast.Substring(Properties.Settings.Default.ViewFast.LastIndexOf('.')), "Ошибка расширения", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    String Command = "Select * from [" + SheetName + "]";
                    //String Command = "Select * from [Sheet_1$]";
                    //OleDbConnection con = new OleDbConnection(connection);
                    
                    //con.Open();
                    OleDbCommand cmd = new OleDbCommand(Command, con);
                    OleDbDataAdapter db = new OleDbDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    db.Fill(dt);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i].ItemArray.GetValue(0).ToString().Length==0)
                        {
                            dt.Rows[i].Delete();
                            
                            //i -= 1;
                        }

                    }
                    dt.AcceptChanges();
                    dataGridViewFast.ItemsSource = dt.AsDataView();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                }
                else
                {
                   
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
                    //OleDbConnection con = new OleDbConnection(connection);

                    //con.Open();
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
                    //dataGridViewFast.ItemsSource = dt_input.AsDataView();

                    string BufStringExcel;
                    List<string> TeacherList = new List<string>();
                    for (int j = 0; j < dt_input.Rows.Count; j++)
                    {
                        BufStringExcel = dt_input.Rows[j].ItemArray.GetValue(3).ToString();
                        if (TeacherList.IndexOf(BufStringExcel) < 0)
                        {
                            TeacherList.Add(BufStringExcel);
                            vfDTR.Add(new DataViewFastTableRow(
                            BufStringExcel
                           , dt_input.Rows[j].ItemArray.GetValue(7).ToString()
                           ));
                        }
                        CollectionViewSource.GetDefaultView(dataGridViewFast.ItemsSource).Refresh();
                    }
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();

                    ds = new DataSet();
                    DataTable dt = new DataTable("Sheet_1");
                    ds.Tables.Add(dt);
                    dt.Columns.Add("Teacher", Type.GetType("System.String"));
                    dt.Columns.Add("Result", Type.GetType("System.String"));
                    for (i=0;i<vfDTR.Count;i++)
                    {
                        dt.Rows.Add(vfDTR[i].Teacher, vfDTR[i].Result);
                    }
                    for(;i<=100;i++)
                    {
                        dt.Rows.Add("","");
                    }
                    DataSetHelper.CreateWorkbook(pathFast, ds);
                    ds.Dispose();
                    dt.Dispose();
                }
                //exApp.Quit();
                dataGridViewFast.Columns[0].Header = "Преподаватель";
                dataGridViewFast.Columns[1].Header = "Всего часов";
            }
            CollectionViewSource.GetDefaultView(dataGridViewFast.ItemsSource).Refresh();
        }


        public void UpdateView()
        {
            string pathB = Properties.Settings.Default.PathToGlobal +"\\"+ Properties.Settings.Default.GlobalMarker;
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению, на данный момент обновление невозможно - другой пользователь обновляет общий файл.\nПопробуйте еще раз чуть позже");
            }
            else
            {
                string pathA = Properties.Settings.Default.PathToGlobalData;
                string pathC = Directory.GetCurrentDirectory() + "\\" + "View_"+Properties.Settings.Default.GlobalData;
                if (File.Exists(pathC))
                {
                    FileInfo localdata = new FileInfo(pathC);
                    FileInfo globaldata = new FileInfo(pathA);
                    if(globaldata.LastWriteTime.ToLocalTime() > localdata.LastWriteTime.ToLocalTime())
                    {
                        localdata.IsReadOnly = false;
                        File.Delete(pathC);
                        File.Copy(pathA, pathC);
                        localdata.IsReadOnly = true;
                        CollectionViewSource.GetDefaultView(dataGridView.ItemsSource).Refresh();
                    }
                    else
                    {
                        if(dataGridView.Items.Count>0)
                        return;
                    }                    
                }
                else
                {
                    File.Copy(pathA, pathC);
                    FileInfo localdata = new FileInfo(pathC);
                    localdata.IsReadOnly = true;
                    
                }

                vDTR.Clear();
               
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
                for (i = 0; i < dt_input.Rows.Count; i++)
                {
                    vDTR.Add(new DataViewTableRow(
                            dt_input.Rows[i].ItemArray.GetValue(1).ToString().Length > 0 ? (dt_input.Rows[i].ItemArray.GetValue(1).ToString().LastIndexOf(" ") > 0 ? dt_input.Rows[i].ItemArray.GetValue(1).ToString().Substring(0, (dt_input.Rows[i].ItemArray.GetValue(1).ToString().IndexOf(" "))) : dt_input.Rows[i].ItemArray.GetValue(1).ToString()) : "",
                            dt_input.Rows[i].ItemArray.GetValue(2).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(2).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(3).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(3).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(4).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(4).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(5).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(5).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(6).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(6).ToString() : "",
                            dt_input.Rows[i].ItemArray.GetValue(7).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(7).ToString() : ""
                            ));
                }

                //dataGridViewFast.ItemsSource = dt_input.AsDataView();
                cmd.Dispose();
                con.Close();
                con.Dispose();
                //exApp.Quit();
                CollectionViewSource.GetDefaultView(dataGridView.ItemsSource).Refresh();
            }
        }

        private void dataGridView_Loaded(object sender, RoutedEventArgs e)
        {
            if (dataGridView.Columns.Count > 0)
            {
                dataGridView.Columns[0].Header = "Дата проведения";
                dataGridView.Columns[1].Header = "Время проведения";
                dataGridView.Columns[2].Header = "Преподаватель";
                dataGridView.Columns[3].Header = "Номер группы";
                dataGridView.Columns[4].Header = "Категория слушателей";
                dataGridView.Columns[5].Header = "Место проведения";
                dataGridView.Columns[6].Header = "Итого";
                dataGridView.Columns[0].MaxWidth = 100;
                dataGridView.Columns[1].MaxWidth = 100;
                dataGridView.Columns[2].MaxWidth = 200;
                dataGridView.Columns[3].MaxWidth = 100;
                dataGridView.Columns[6].MaxWidth = 60;
            }
            CollectionViewSource.GetDefaultView(dataGridView.ItemsSource).Refresh();

        }
        private void dataGridViewFast_Loaded(object sender, RoutedEventArgs e)
        {
            
            if (dataGridView.Columns.Count > 0)
            {
                dataGridViewFast.Columns[0].Header = "Преподаватель";
                dataGridViewFast.Columns[1].Header = "Всего часов";
            }
            CollectionViewSource.GetDefaultView(dataGridViewFast.ItemsSource).Refresh();
        }
        private void dataGridExport_Loaded(object sender, RoutedEventArgs e)
        {
            if(dataGridExport.Columns.Count>0)
            {
                dataGridExport.Columns[0].Header = "Дата проведения";
                dataGridExport.Columns[1].Header = "Время проведения";
                dataGridExport.Columns[2].Header = "Преподаватель";
                dataGridExport.Columns[3].Header = "Номер группы";
                dataGridExport.Columns[4].Header = "Категория слушателей";
                dataGridExport.Columns[5].Header = "Место проведения";
                dataGridExport.Columns[0].MaxWidth = 120;
                dataGridExport.Columns[1].MaxWidth = 120;
                dataGridExport.Columns[2].MaxWidth = 200;
                dataGridExport.Columns[3].MaxWidth = 120;
            }
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
        }

        private void buttonView_Download_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            UpdateLayout();
            SaveExcel();
            this.Effect = null;
            UpdateLayout();
        }

        public void SaveExcel()
        {
            string pathB = Properties.Settings.Default.PathToGlobal +"\\"+ Properties.Settings.Default.GlobalMarker;
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению, на данный момент скачивание невозможно - другой пользователь обновляет общий файл.\nПопробуйте еще раз чуть позже");
            }
            else
            {
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.FileName = "Общий файл расписания (только для просмотра)";
                dlg.DefaultExt = ".xlsx";
                dlg.Filter = "Книга Excel(.xlsx)|*.xlsx";

                Nullable<bool> result = dlg.ShowDialog();

                if (result == true)
                {
                    string pathSave = dlg.FileName;
                    string pathA = Properties.Settings.Default.PathToGlobalData;
                    FileInfo localdata;
                    FileInfo globaldata = new FileInfo(pathA);
                    if (File.Exists(pathSave))
                    {
                        localdata = new FileInfo(pathSave);
                        localdata.IsReadOnly = false;
                        File.Delete(pathSave);
                    }
                    File.Copy(pathA, pathSave);
                    localdata = new FileInfo(pathSave);
                    localdata.IsReadOnly = true;
                }
            }
        }

        private void SvodExcel_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if(dataGridExport.Items.Count>0)
            {
                if(MessageBox.Show("В подготовленных для экспорта остались не отправленные записи ("+ dataGridExport.Items.Count.ToString() + ").\nВы действительно хотите закрыть приложение не отправив новые записи в общий файл?","Закрытие программы с неотправленными данными",MessageBoxButton.OKCancel,MessageBoxImage.Exclamation,MessageBoxResult.Cancel)== MessageBoxResult.OK)
                {
                    e.Cancel = false;
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void dataGridViewFast_AddingNewItem(object sender, AddingNewItemEventArgs e)
        {
            //buttonSaveFast.IsEnabled = true;
        }

        private void dataGridViewFast_LayoutUpdated(object sender, EventArgs e)
        {
            if (dataGridViewFast.Items.Count > 0)
            { buttonSaveFast.IsEnabled = true;
                MenuItemSaveFast.IsEnabled = true;
                dataGridViewFast.Columns[0].Header = "Преподаватель";
                dataGridViewFast.Columns[1].Header = "Всего часов";
            }              
            else
            {
                MenuItemSaveFast.IsEnabled = false;
                buttonSaveFast.IsEnabled = false; }
              
        }

        private void dataGridExport_LayoutUpdated(object sender, EventArgs e)
        {
            if (dataGridExport.Items.Count > 0)
            {
                buttonExportHot.IsEnabled = true;
                buttonExport.IsEnabled = true;

            }                
            else
            {
                buttonExportHot.IsEnabled = false;
                buttonExport.IsEnabled = false;
            }                
        }
        private void DataGridCell_PreviewSelected(object sender, RoutedEventArgs e)
        {
            switch(tabControl.SelectedIndex)
            {
                case 0:
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
                    break;
                case 3:
                    if (dataGridViewEdit.SelectedIndex >= 0)
                    {
                        buttonEditInputHot.IsEnabled = true;
                        buttonDeleteHot.IsEnabled = true;
                    }
                    else
                    {
                        buttonDeleteHot.IsEnabled = false;
                        buttonEditInputHot.IsEnabled = false;
                    }
                    break;
                default:
                    buttonDeleteHot.IsEnabled = false;
                    buttonEditInputHot.IsEnabled = false;
                    break;
            }
            
                
        }

        private void buttonSaveFast_Click(object sender, RoutedEventArgs e)
        {
            SaveFastResult();  
        }
        public void SaveFastResult()//Сохранение краткой сводки
        {
            string pathFast = Directory.GetCurrentDirectory() + "\\" + Properties.Settings.Default.ViewFast;
            if (File.Exists(pathFast))
            {
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.FileName = "Краткая сводка по общему файлу (только для просмотра)";
                dlg.DefaultExt = Properties.Settings.Default.ViewFast.Substring(Properties.Settings.Default.ViewFast.LastIndexOf('.'));
                //dlg.DefaultExt = ".xlsx";
                //dlg.Filter = "Книга Excel(.xlsx)|*.xlsx";
                dlg.Filter = "Книга Excel(."+ dlg.DefaultExt + ")|*."+ dlg.DefaultExt;

                Nullable<bool> result = dlg.ShowDialog();

                if (result == true)
                {
                    string pathSave = dlg.FileName;
                    FileInfo localdata;
                    FileInfo globaldata = new FileInfo(pathFast);
                    if (File.Exists(pathSave))
                    {
                        localdata = new FileInfo(pathSave);
                        localdata.IsReadOnly = false;
                        File.Delete(pathSave);
                    }
                    File.Copy(pathFast, pathSave);
                    localdata = new FileInfo(pathSave);
                    localdata.IsReadOnly = true;
                }
            }            
        }

        public void AdminModeActive()
        {
            AdminModeActive(AdminMode);
        }
        public void AdminModeActive(bool AdminModeSwitch)
        {
            AdminMode = !AdminMode;
            if (AdminMode)
            {
                buttonAdminOff.IsEnabled = true;
                buttonAdminOff.Visibility = Visibility.Visible;
                MenuItemAdminOff.IsEnabled = true;
                MenuItemAdminOff.Visibility = Visibility.Visible;
                MenuItemOptions.IsEnabled = true;
                buttonAdmin.IsEnabled = false;
                buttonAdmin.Visibility = Visibility.Collapsed;
                MenuItemAdmin.IsEnabled = false;
                MenuItemAdmin.Visibility = Visibility.Collapsed;
                ViewEditTab.IsEnabled = true;
                ViewEditTab.Visibility = Visibility.Visible;
            }
            else
            {
                buttonAdminOff.IsEnabled = false;
                buttonAdminOff.Visibility = Visibility.Collapsed;
                MenuItemAdminOff.IsEnabled = false;
                MenuItemAdminOff.Visibility = Visibility.Collapsed;
                MenuItemOptions.IsEnabled = false;
                buttonAdmin.IsEnabled = true;
                buttonAdmin.Visibility = Visibility.Visible;
                MenuItemAdmin.IsEnabled = true;
                MenuItemAdmin.Visibility = Visibility.Visible;
                ViewEditTab.IsEnabled = false;
                ViewEditTab.Visibility = Visibility.Collapsed;
            }
        }

        private void buttonAdmin_Click(object sender, RoutedEventArgs e)
        {
            if (!AdminMode) {
                if (MessageBox.Show("Ты бэтмен?", "Переход в режим администрирования", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                {
                    InputPassword IPas = new InputPassword();
                    switch(IPas.ShowDialog())
                    {
                        case true:
                        AdminModeActive();
                            break;
                        case false:
                            break;
                        default:
                            break; 
                    }
                    
                }
            }else
                {
                    AdminModeActive();
                }
        }

        private void MenuItemOptions_Click(object sender, RoutedEventArgs e)
        {
            Options Op = new Options();
            Op.ShowDialog();
        }

        private void openWindowOpenFileTable(string[] dataString=null)
        {
            OpenFileTable f = new OpenFileTable(dataString,this);
            //f.Owner = this;
            f.Top = this.Top + 50;
            f.Left = this.Left + 50;
            //f.Show();
            //f.Hide();
            //if(dataString!=null)
              //  f.AddFilesToOpen(dataString);
            f.ShowDialog();

            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
        }


        private void buttonFileInputHot_Click(object sender, RoutedEventArgs e)
        {
            openWindowOpenFileTable();
        }

        private void SvodExcel_Drop(object sender, DragEventArgs e)
        {
            switch (tabControl.SelectedIndex)
            {
                case 0:
                    if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    {
                        string[] dataString = (string[])e.Data.GetData(DataFormats.FileDrop);
                        openWindowOpenFileTable(dataString);
                    }
                    break;
                default:
                    break;
            }
            
        }

        private void dataGridExport_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if(ClickToAddRow)
            {
                    openSingleInput();
            }
            else
            {
                ClickToAddRow = true;
            }
        }
        private void dataGridExportItemsChanges(object sender, NotifyCollectionChangedEventArgs e)
        {
            StatusStringCountRecordAllFile.Content = dataGridExport.Items.Count.ToString();
        }

        private void dataGridExport_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key==Key.Delete)
            {
                switch(tabControl.SelectedIndex)
                {
                case 0:
                        if(DTR.Count > 0 && dataGridExport.SelectedIndex >= 0)
                        DeleteItem(dataGridExport.SelectedIndex);
                    break;
                case 3:
                        if (vDTR.Count > 0 && dataGridViewEdit.SelectedIndex >= 0)
                            DeleteItemEdition(dataGridViewEdit.SelectedIndex);
                    break;
                default:
                    break;
                }
                
            }
        }

        private void MenuItemEditInput_Click(object sender, RoutedEventArgs e)
        {
            ChangeDataGrid();
        }

        private void MenuItemFileInput_Click(object sender, RoutedEventArgs e)
        {
            openWindowOpenFileTable();
        }

        private void dataGridViewEdit_Loaded(object sender, RoutedEventArgs e)
        {
            if (dataGridViewEdit.Columns.Count > 0)
            {
                dataGridViewEdit.Columns[0].Header = "Дата проведения";
                dataGridViewEdit.Columns[1].Header = "Время проведения";
                dataGridViewEdit.Columns[2].Header = "Преподаватель";
                dataGridViewEdit.Columns[3].Header = "Номер группы";
                dataGridViewEdit.Columns[4].Header = "Категория слушателей";
                dataGridViewEdit.Columns[5].Header = "Место проведения";
                dataGridViewEdit.Columns[0].MaxWidth = 120;
                dataGridViewEdit.Columns[1].MaxWidth = 120;
                dataGridViewEdit.Columns[2].MaxWidth = 200;
                dataGridViewEdit.Columns[3].MaxWidth = 120;
            }
            if(dataGridViewEdit.ItemsSource!=null)
                CollectionViewSource.GetDefaultView(dataGridViewEdit.ItemsSource).Refresh();
        }

        private void buttonDisconnect_Click(object sender, RoutedEventArgs e)
        {
            //if(NoneSave)
            if(MAD.Count>0)
            {
                if (MessageBox.Show("Вы действительно хотите отключиться от общих данных не сохранив изменений?", "Подтверждение отключения", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                {
                    ConnectDisconnect();
                }
            }
            else
            {
                ConnectDisconnect();
            }
            
        }        

        private void buttonViewEdit_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите подключиться к общим данным?\nВнимание пока вы не отключитесь от общих данных, все остальные пользователи не смогут вносить изменения в общие данные","Подтверждение подключения",MessageBoxButton.OKCancel,MessageBoxImage.Warning) == MessageBoxResult.OK)
            {
                ConnectDisconnect();
            }            
        }

        private void dataGridViewEdit_LayoutUpdated(object sender, EventArgs e)
        {
            if(MAD.Count>0)
            {
                buttonViewEdit_Download.IsEnabled = true;
            }
            else
            {
                buttonViewEdit_Download.IsEnabled = false;
            }
        }

        private void buttonViewEdit_Download_Click(object sender, RoutedEventArgs e)
        {
            if(MAD.Count>0)
                if(
                    MessageBox.Show("Отредактировано/удалено "+ MAD.Count.ToString()+" записей."+
                        "\nВы действительно хотите внести изменения в общий файл?",
                    "Подтверждение внесения изменений", MessageBoxButton.YesNo,MessageBoxImage.Question)==MessageBoxResult.Yes
                    )
                {
                    string pathC = Directory.GetCurrentDirectory() + "\\" + "ViewEdit_" + Properties.Settings.Default.GlobalData;
                    if (File.Exists(pathC))
                    {
                        FileInfo localdata = new FileInfo(pathC);
                        try {                               
                                localdata.IsReadOnly = false;
                            }
                        catch
                            {
                                MessageBox.Show("Ошибка обращения к локальной копии сводного документа \n("+pathC+").\nПерезапустите компьютер");
                                return;
                            }

                        var exBook = exApp.Workbooks.Open(pathC);
                        var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                        int BlinkEnd = 1;
                        var lastcell = ExSheet.Cells.SpecialCells(Type: Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                        //if (ExSheet.Cells[lastcell.Row, 2].Value != null || ExSheet.Cells[lastcell.Row, 3].Value != null || ExSheet.Cells[lastcell.Row, 4].Value != null || ExSheet.Cells[lastcell.Row, 5].Value != null || ExSheet.Cells[lastcell.Row, 6].Value != null || ExSheet.Cells[lastcell.Row, 7].Value != null)
                          //  BlinkEnd = 1;
                        Regex regexTime = new Regex(@"^( *[Cc] *)?\d{1,2} *[\.,:;\- ]? *\d{1,2}(( *[\-–\/\\\| ] *)|( +)|( *[Дд][Оо] *))\d{1,2} *[\.,:;\- ]? *\d{1,2} *$");
                        for(int i=1;i<=lastcell.Row;i++)
                        {
                            string bufstr = Convert.ToString(ExSheet.Cells[i, 3].Value2);
                            if(bufstr!=null)
                            {
                                if (bufstr.Length > 0)
                                {
                                    if (regexTime.IsMatch(bufstr))
                                    {
                                        //MessageBox.Show(i.ToString()+"\n"+ bufstr);
                                        BlinkEnd = i;
                                        break;
                                    }
                                    else
                                    {
                                        //MessageBox.Show("!"+i.ToString() + "\n" + bufstr);
                                    }
                                }
                            }
                        }
                            //for (int i = BlinkEnd; i < (DTR.Count + BlinkEnd); i++)
                            //{
                            //    ExSheet.Cells[lastcell.Row + i, 2] = DTR[i - BlinkEnd].Date;
                            //    ExSheet.Cells[lastcell.Row + i, 3] = DTR[i - BlinkEnd].Time;
                            //    ExSheet.Cells[lastcell.Row + i, 4] = DTR[i - BlinkEnd].Teacher;
                            //    ExSheet.Cells[lastcell.Row + i, 5] = DTR[i - BlinkEnd].Group;
                            //    ExSheet.Cells[lastcell.Row + i, 6] = DTR[i - BlinkEnd].Category;
                            //    ExSheet.Cells[lastcell.Row + i, 7] = DTR[i - BlinkEnd].Place;
                            //}
                        for (int i=0;i<MAD.Count;i++)
                        {
                            if(MAD[i].Action)
                            {
                                //Редактирование
                            }
                            else
                            {
                                ExSheet.Rows[MAD[i].IndexData+ BlinkEnd].Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                                labelTech2.Content = labelTech2.Content + "\n" + (MAD[i].IndexData + BlinkEnd).ToString();
                            }
                        }
                        exBook.Close(true);

                        localdata.IsReadOnly = true;
                    }

                    MAD.Clear();
                    buttonViewEdit.IsEnabled = false;
                }
        }

        public void ConnectDisconnect()
        {
            ConnectDisconnect(ConnectMode);
        }
        public void ConnectDisconnect(bool ConnectSwitch)
        {
            ConnectMode =!ConnectSwitch;
            string pathB = Properties.Settings.Default.PathToGlobal + "\\" + Properties.Settings.Default.GlobalMarker;
            string pathA = Properties.Settings.Default.PathToGlobalData;
            string pathC = Directory.GetCurrentDirectory() + "\\" + "ViewEdit_" + Properties.Settings.Default.GlobalData;
            if (ConnectMode)
            {
                if (File.Exists(pathB))
                {
                    MessageBox.Show("К сожалению, на данный момент обновление невозможно - другой пользователь обновляет общий файл.\nПопробуйте еще раз чуть позже");
                }
                else
                {
                    NoneSave = false;
                    MAD.Clear();
                    buttonDisconnect.IsEnabled = true;
                    buttonViewEdit.IsEnabled = false;
                    ExportTab.IsEnabled = false;
                    ViewSmallTab.IsEnabled = false;
                    ViewTab.IsEnabled = false;

                    StreamWriter sw = File.CreateText(pathB);
                    String host = System.Net.Dns.GetHostName();
                    System.Net.IPAddress ip = System.Net.Dns.GetHostEntry(host).AddressList[0];
                    sw.WriteLine(ip.ToString());
                    sw.Close();
                    if (File.Exists(pathC))
                    {
                        FileInfo localdata = new FileInfo(pathC);
                        FileInfo globaldata = new FileInfo(pathA);
                        //if (globaldata.LastWriteTime.ToLocalTime() > localdata.LastWriteTime.ToLocalTime())
                        {
                            localdata.IsReadOnly = false;
                            File.Delete(pathC);
                            File.Copy(pathA, pathC);
                            localdata.IsReadOnly = true;
                            CollectionViewSource.GetDefaultView(dataGridView.ItemsSource).Refresh();
                        }
                        //else
                        {
                            if (dataGridView.Items.Count > 0)
                                return;
                        }

                    }
                    else
                    {
                        File.Copy(pathA, pathC);
                        FileInfo localdata = new FileInfo(pathC);
                        localdata.IsReadOnly = true;

                    }

                    vDTR.Clear();

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
                    for (i = 0; i < dt_input.Rows.Count; i++)
                    {
                        vDTR.Add(new DataViewTableRow(
                                dt_input.Rows[i].ItemArray.GetValue(1).ToString().Length > 0 ? (dt_input.Rows[i].ItemArray.GetValue(1).ToString().LastIndexOf(" ")>0 ? dt_input.Rows[i].ItemArray.GetValue(1).ToString().Substring(0, (dt_input.Rows[i].ItemArray.GetValue(1).ToString().IndexOf(" "))) : dt_input.Rows[i].ItemArray.GetValue(1).ToString()) : "",
                                dt_input.Rows[i].ItemArray.GetValue(2).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(2).ToString() : "",
                                dt_input.Rows[i].ItemArray.GetValue(3).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(3).ToString() : "",
                                dt_input.Rows[i].ItemArray.GetValue(4).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(4).ToString() : "",
                                dt_input.Rows[i].ItemArray.GetValue(5).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(5).ToString() : "",
                                dt_input.Rows[i].ItemArray.GetValue(6).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(6).ToString() : "",
                                dt_input.Rows[i].ItemArray.GetValue(7).ToString().Length > 0 ? dt_input.Rows[i].ItemArray.GetValue(7).ToString() : ""
                                ));
                    }

                    //dataGridViewFast.ItemsSource = dt_input.AsDataView();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    //exApp.Quit();
                    dataGridViewEdit.ItemsSource = vDTR;
                    
                    CollectionViewSource.GetDefaultView(dataGridViewEdit.ItemsSource).Refresh();
                    if (dataGridViewEdit.Columns.Count > 0)
                    {
                        dataGridViewEdit.Columns.Remove(dataGridViewEdit.Columns[dataGridViewEdit.Columns.Count - 1]);
                        dataGridViewEdit.Columns[0].Header = "Дата проведения";
                        dataGridViewEdit.Columns[1].Header = "Время проведения";
                        dataGridViewEdit.Columns[2].Header = "Преподаватель";
                        dataGridViewEdit.Columns[3].Header = "Номер группы";
                        dataGridViewEdit.Columns[4].Header = "Категория слушателей";
                        dataGridViewEdit.Columns[5].Header = "Место проведения";
                        dataGridViewEdit.Columns[0].MaxWidth = 120;
                        dataGridViewEdit.Columns[1].MaxWidth = 120;
                        dataGridViewEdit.Columns[2].MaxWidth = 200;
                        dataGridViewEdit.Columns[3].MaxWidth = 120;
                    }
                    dataGridViewEdit.UpdateLayout();
                }
            }
            else
            {
                if (File.Exists(pathB))
                {
                    FileInfo employed = new FileInfo(pathB);
                    StreamReader sw = new StreamReader(pathB);
                    String host = System.Net.Dns.GetHostName();
                    System.Net.IPAddress ip = System.Net.Dns.GetHostEntry(host).AddressList[0];
                    string busy_customer = "";
                    busy_customer = sw.ReadLine();
                    sw.Close();
                    if ((busy_customer == ip.ToString()) && (DateTime.Now.Subtract(employed.CreationTime.ToLocalTime()).TotalSeconds > Properties.Settings.Default.WaitingDisconnect))
                    {
                        File.Delete(pathB);
                        int i = 0;
                        while (File.Exists(pathB))
                        {
                            if (i > 10000)
                                break;
                            i++;
                        }
                        buttonDisconnect.IsEnabled = false;
                        buttonViewEdit.IsEnabled = true;
                        ExportTab.IsEnabled = true;
                        ViewSmallTab.IsEnabled = true;
                        ViewTab.IsEnabled = true;
                        dataGridViewEdit.ItemsSource = null;
                        dataGridViewEdit.UpdateLayout();
                        MAD.Clear();
                        if (File.Exists(pathC))
                        {
                            FileInfo localdata = new FileInfo(pathC);
                            localdata.IsReadOnly = false;
                            File.Delete(pathC);
                            while (File.Exists(pathB))
                            {
                                if (i > 10000)
                                    break;
                                i++;
                            }
                        }
                    }
                    else
                    {
                        if(busy_customer != ip.ToString())
                        {
                            MessageBox.Show("Ошибка отключения: общий файл занят другим пользователем","Ошибка отключения",MessageBoxButton.OK,MessageBoxImage.Error);
                        }
                        if(DateTime.Now.Subtract(employed.CreationTime.ToLocalTime()).TotalMinutes > Properties.Settings.Default.WaitingInLine)
                        {
                            MessageBox.Show("Ошибка отключения: преждевременная попытка отключения", "Ошибка отключения", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
        }
    }
}
