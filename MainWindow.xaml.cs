﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Controls;
using System.Text.RegularExpressions;

namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
        public class DataTableRow
        {
            public string Date{ get; set; }
            public string Time { get; set; }
            public string Teacher { get; set; }
            public string Group { get; set; }
            public string Category { get; set; }
            public string Place { get; set; }

            public DataTableRow(string inputDate, string inputTime, string inputTeacher, string inputGroup, string inputCategory, string inputPlace)
            {
                Date = inputDate;
                if (inputTime.Length > 0)
                {
                    if (inputTime[0] == '0')
                    {
                        Time = inputTime.Substring(1).Replace(':', '.');
                    }
                    else
                    {
                        Time = inputTime.Replace(':', '.');
                    }
                    if (Time[Time.IndexOf("-") + 1] == '0')
                    {
                        Time = Time.Substring(0, Time.IndexOf("-") + 1) + Time.Substring(Time.IndexOf("-") + 2);
                    }
                }
                else
                {
                    Time = null;
                }
                
                Teacher = inputTeacher;
                Group = inputGroup;
                Category = inputCategory;
                Place = inputPlace;
            }
            public DataTableRow()
            {
                Date = null;
                Time = null;
                Teacher = null;
                Group = null;
                Category = null;
                Place = null;
            }
        }
        public class DataViewTableRow
        {
            public string Date { get; set; }
            public string Time { get; set; }
            public string Teacher { get; set; }
            public string Group { get; set; }
            public string Category { get; set; }
            public string Place { get; set; }
            public string Result { get; set; }

            public DataViewTableRow(string inputDate, string inputTime, string inputTeacher, string inputGroup, string inputCategory, string inputPlace, string inputResult=null)
            {
                Date = inputDate;
                if (inputTime.Length > 0)
                {
                    if (inputTime[0] == '0')
                    {
                        Time = inputTime.Substring(1).Replace(':', '.');
                    }
                    else
                    {
                        Time = inputTime.Replace(':', '.');
                    }
                    if (Time[Time.IndexOf("-") + 1] == '0')
                    {
                        Time = Time.Substring(0, Time.IndexOf("-") + 1) + Time.Substring(Time.IndexOf("-") + 2);
                    }
                }
                else
                {
                    Time = null;
                }

                Teacher = inputTeacher;
                Group = inputGroup;
                Category = inputCategory;
                Place = inputPlace;
                Result = inputResult;
            }
            public DataViewTableRow()
            {
                Date = null;
                Time = null;
                Teacher = null;
                Group = null;
                Category = null;
                Place = null;
                Result = null;
            }
            public DataViewTableRow(DataTableRow InputData)
            {
                Date = InputData.Date;
                Time = InputData.Time;
                Teacher = InputData.Teacher;
                Group = InputData.Group;
                Category = InputData.Category;
                Place = InputData.Place;
                Result = null;
            }
        }
        public List<DataTableRow> DTR = new List<DataTableRow>();
        public List<DataViewTableRow> vDTR = new List<DataViewTableRow>();

        public MainWindow()
        {
            InitializeComponent();
            DTR.Clear();
            vDTR.Clear();

            dataGridExport.ItemsSource = DTR;
            dataGridView.ItemsSource = vDTR;
        }
        private void SvodExcel_Loaded(object sender, RoutedEventArgs e)
        {
            DTR.Clear();
            vDTR.Clear();
 
            // example data
           //AddNewItem(new DataTableRow("06.11.2019", "10:00-16:40", "Пронина Л.Н.", "","******","!@#$%&"));
           // AddNewItem(new DataTableRow("07.11.2019", "12:00-18:40", "Радюхина Е.И.", "", "#######", "*?!~%$#"));
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            //----exmpla data

            ClearHang();
            //buttonDebug.Visibility = Visibility.Collapsed;
        }
        private void SvodExcel_Closed(object sender, EventArgs e)
        {
            ClearHang();
            exApp.Quit();
            System.Windows.Application.Current.Shutdown();
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

        private void MenuItemSingleInput_Click(object sender, RoutedEventArgs e)
        {
            SingleInput f = new SingleInput();
            f.Top = this.Top+50;
            f.Left = this.Left+50;
            f.RowIndex = -1;
            f.ShowDialog();
            //f.Show();

            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
        }

        public void AddNewItem(DataTableRow newDTR)
        {
            DTR.Add(newDTR);
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            buttonExport.IsEnabled = true;
            buttonExportHot.IsEnabled = true;
            buttonDeleteHot.IsEnabled = true;
        }


        private void DataGridCell_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;
            ChangeDataGrid();
        }
        private void ChangeDataGrid()
        {
            switch (tabControl.SelectedIndex)
            {
                case 0:
                    {
                        int SI = dataGridExport.SelectedIndex;
                        SingleInput f = new SingleInput();
                        f.Top = this.Top + 50;
                        f.Left = this.Left + 50;
                        f.RowIndex = dataGridExport.SelectedIndex;
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
                default:
                    break;
            }
            
        }
        public void EditItem(int RowIndex,DataTableRow newDTR)
        {
            DTR[RowIndex] = newDTR;
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
        }
        public void DeleteItem(int RowIndex)
        {
            if (RowIndex >= 0 && RowIndex < DTR.Count)
            {
                if(MessageBox.Show("Вы действительно хотите удалить из экспортируемых данных запись\n"+ DTR[RowIndex].Date + " " + DTR[RowIndex].Time + " " + DTR[RowIndex].Teacher+"\n?","Удаление элемента из экспорта",MessageBoxButton.YesNo,MessageBoxImage.Question,MessageBoxResult.No) ==MessageBoxResult.Yes)
                DTR.Remove(DTR[RowIndex]);
                CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
                if(DTR.Count<1)
                {
                    buttonDeleteHot.IsEnabled = false;
                    buttonExport.IsEnabled = false;
                    buttonExportHot.IsEnabled = false;
                }
            }
            else
                MessageBox.Show("Ошибка удаления элемента");
        }
        private void buttonDeleteHot_Click(object sender, RoutedEventArgs e)
        {
            DeleteItem(dataGridExport.SelectedIndex);
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
                ExportData();
                newWindowThread.Abort();
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
        public void ExportData()
        {           
            string pathB = Properties.Settings.Default.PathToGlobal+Properties.Settings.Default.GlobalMarker;
            ClearHang();
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению, на данный момент экспорт невозможен - другой пользователь уже начал оновлять общий файл!\nПопробуйте еще раз чуть позже");
            }
            else
            {
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
                /*
                 MessageBox.Show(
                    exBook.Path.ToString()+"\\"+exBook.Name.ToString()+ "\n" + (lastcell.Row -1).ToString() + " 4:\n" +
                    ExSheet.Cells[lastcell.Row-1, 4].Value.ToString()
                                       +"\n"+
                    pathC.ToString()+"\n"+(lastcell.Row + DTR.Count - 1).ToString()+" 4:\n"+
                    ExSheet.Cells[lastcell.Row + DTR.Count-1, 4].Value.ToString()
                    );
                    */
                exBook.Close(true);
                //exApp.Quit();
                File.Move(pathA, pathA.Substring(0, pathA.Length - 5) + " " + DateTime.Now.ToString().Replace(':', '_') + ".xlsx");
                File.Copy(pathC, pathA);
                //File.Replace(pathC,pathA,pathA.Substring(0, pathA.Length-5)+ " "+DateTime.Now.ToString().Replace(':','_')+ ".xlsx");
                File.Delete(pathC);
                File.Delete(pathB);
                DTR.Clear();
                CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            }
        }
        public void ClearHang()
        {
            string pathB = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
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
            string pathB = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
            ClearHang();
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению, на данный момент обновление невозможно - другой пользователь уже начал обновлять общий файл!\nПопробуйте еще раз чуть позже");
            }
            else
            {
                string pathC = Directory.GetCurrentDirectory() + "\\" + Properties.Settings.Default.GlobalData;
                if (File.Exists(pathC))
                {
                    File.Delete(pathC);
                }
                string pathA = Properties.Settings.Default.PathToGlobalData;
                File.Copy(pathA, pathC);
                //var exApp = new Microsoft.Office.Interop.Excel.Application();
                var exBook = exApp.Workbooks.Open(pathC);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                string FormulCalculate = ExSheet.Cells[16, 8].Formula;
                exBook.Close(true);
                //exApp.Quit();
                File.Delete(pathC);
                List<string> TimeTemplate = new List<string>();
                Regex regex = new Regex(@"\d{1,2}\.\d{2}\-\d{1,2}\.\d{2}");
                MatchCollection matchList = regex.Matches(FormulCalculate);
                for(int i=0;i<matchList.Count;i++)
                {
                    TimeTemplate.Add(matchList[i].Value);
                }
            }
        }

        private void tabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch(tabControl.SelectedIndex)
            {
                case 0:
                    menu_Hot_Export.IsEnabled = true;
                    menu_Hot_View.IsEnabled = false;

                    menu_Hot_Export.Visibility = Visibility.Visible;
                    menu_Hot_View.Visibility = Visibility.Collapsed;
                    break;
                case 2:
                    menu_Hot_Export.IsEnabled = false;
                    menu_Hot_View.IsEnabled = true;

                    menu_Hot_Export.Visibility = Visibility.Collapsed;
                    menu_Hot_View.Visibility = Visibility.Visible;
                    break;
                default:
                    menu_Hot_Export.IsEnabled = false;
                    menu_Hot_View.IsEnabled = false;

                    menu_Hot_Export.Visibility = Visibility.Visible;
                    menu_Hot_View.Visibility = Visibility.Visible;
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
                UpdateView();
                newWindowThread.Abort();
            }           
            this.Effect = null;
            UpdateLayout();

        }
        public void UpdateView()
        {
            string pathB = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
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
                
                var exBook = exApp.Workbooks.Open(pathC);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                var lastcell = ExSheet.Cells.SpecialCells(Type: Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                int BlinkEnd = 0;
                if (ExSheet.Cells[lastcell.Row, 2].Value != null || ExSheet.Cells[lastcell.Row, 3].Value != null || ExSheet.Cells[lastcell.Row, 4].Value != null || ExSheet.Cells[lastcell.Row, 5].Value != null || ExSheet.Cells[lastcell.Row, 6].Value != null || ExSheet.Cells[lastcell.Row, 7].Value != null)
                    BlinkEnd = 1;
                bool flag = true;
                if (lastcell.Row > 100)
                {
                    if (MessageBox.Show("Вы действительно хотите просмотреть данные из общего файла?\nЭто может занять несколько ДЕСЯТКОВ минут.\nВсего записей - " + (lastcell.Row + BlinkEnd - 15).ToString(), "Просмотр общих данных БОЛЬШОГО ОБЪЕМА", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
                    {
                        flag=false;
                    }
                }
                if (flag)
                {
                    for (int j = 15; j < lastcell.Row + BlinkEnd - 1; j++)
                    {
                        try
                        {
                            /*
                            vDTR.Add(new DataViewTableRow(ExSheet.Cells[j + 1, 2].Value == null ? "" : ExSheet.Cells[j + 1, 2].Value.ToString()
                                ,ExSheet.Cells[j + 1, 3].Value == null ? "" : ExSheet.Cells[j + 1, 3].Value.ToString()
                                ,ExSheet.Cells[j + 1, 4].Value == null ? "" : ExSheet.Cells[j + 1, 4].Value.ToString()
                                ,ExSheet.Cells[j + 1, 5].Value == null ? "" : ExSheet.Cells[j + 1, 5].Value.ToString()
                                ,ExSheet.Cells[j + 1, 6].Value == null ? "" : ExSheet.Cells[j + 1, 6].Value.ToString()
                                ,ExSheet.Cells[j + 1, 7].Value == null ? "" : ExSheet.Cells[j + 1, 7].Value.ToString()
                                ,ExSheet.Cells[j + 1, 8].Value == null ? "" : ExSheet.Cells[j + 1, 8].Value.ToString()
                                ));
                             */
                            vDTR.Add(new DataViewTableRow(ExSheet.Cells[j + 1, 2].Value == null ? "" : ExSheet.Cells[j + 1, 2].Value.ToString()
                               , ExSheet.Cells[j + 1, 3].Value == null ? "" : ExSheet.Cells[j + 1, 3].Value.ToString()
                               , ExSheet.Cells[j + 1, 4].Value == null ? "" : ExSheet.Cells[j + 1, 4].Value.ToString()
                               , ExSheet.Cells[j + 1, 5].Value == null ? "" : ExSheet.Cells[j + 1, 5].Value.ToString()
                               , ExSheet.Cells[j + 1, 6].Value == null ? "" : ExSheet.Cells[j + 1, 6].Value.ToString()
                               , ExSheet.Cells[j + 1, 7].Value == null ? "" : ExSheet.Cells[j + 1, 7].Value.ToString()
                               , "Технические работы"));
                        }
                        catch
                        {

                        }
                        CollectionViewSource.GetDefaultView(dataGridView.ItemsSource).Refresh();
                        //ListExcel.Add(ExSheet.Cells[j + 1, 4].Value.ToString());
                    }
                }
                          
                exBook.Close(false);
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
            string pathB = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
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
    }
}
