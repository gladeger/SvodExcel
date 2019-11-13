using System;
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
                if(inputTime[0]=='0')
                {
                    Time = inputTime.Substring(1).Replace(':', '.');
                }
                else
                {
                    Time = inputTime.Replace(':', '.');
                }            
                if(Time[Time.IndexOf("-")+1]=='0')
                {
                    Time = Time.Substring(0, Time.IndexOf("-")+1)+ Time.Substring(Time.IndexOf("-")+2);
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
        public List<DataTableRow> DTR = new List<DataTableRow>();

        public MainWindow()
        {
            InitializeComponent();
            DTR.Clear();
            dataGridExport.ItemsSource = DTR;
        }
        private void SvodExcel_Loaded(object sender, RoutedEventArgs e)
        {
            DTR.Clear();
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

            // example data
            AddNewItem(new DataTableRow("06.11.2019", "10:00-16:40", "Пронина Л.Н.", "","******","!@#$%&"));
            AddNewItem(new DataTableRow("07.11.2019", "12:00-18:40", "Радюхина Е.И.", "", "#######", "*?!~%$#"));
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
            //----exmpla data

            ClearHang();
        }
        private void SvodExcel_Closed(object sender, EventArgs e)
        {
            ClearHang();
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
            int SI = dataGridExport.SelectedIndex;
            SingleInput f = new SingleInput();
            f.Top = this.Top + 50;
            f.Left = this.Left + 50;
            f.RowIndex = dataGridExport.SelectedIndex;
            switch(dataGridExport.CurrentColumn.DisplayIndex)
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

            f.MaskedTextBoxStartTime.Text = DTR[SI].Time.Substring(0, 5).Replace('.',':');
            if(f.MaskedTextBoxStartTime.Text[0]=='_')
            {
                f.MaskedTextBoxStartTime.Text = "0"+DTR[SI].Time.Substring(0, 4).Replace('.', ':');
            }
            f.MaskedTextBoxEndTime.Text = DTR[SI].Time.Substring(DTR[SI].Time.Length-5, 5).Replace('.', ':');
            if (f.MaskedTextBoxEndTime.Text[0] == '_')
            {
                f.MaskedTextBoxEndTime.Text = "0" + DTR[SI].Time.Substring(DTR[SI].Time.Length - 4, 4).Replace('.', ':');
            }
            f.comboBoxTeacher.SelectedIndex = f.comboBoxTeacher.Items.IndexOf(DTR[SI].Teacher);
            f.textBoxCategory.Text = DTR[SI].Category;
            f.textBoxPlace.Text = DTR[SI].Place;
            f.Title = "Редактирование записи \""+ DTR[SI].Date+" "+DTR[SI].Time+" "+DTR[SI].Teacher+"\"";
            f.ButtonWriteAndContinue.IsEnabled = false;
            f.ButtonWriteAndContinue.Visibility = Visibility.Collapsed;
            f.ButtonWriteAndStop.Content = "Внести изменения";
            f.ButtonWriteAndStop.HorizontalAlignment = HorizontalAlignment.Left;
            f.ButtonWriteAndStop.Margin= new Thickness(10, 0, 0, 10);
            f.ShowDialog();
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

        private void buttonExport_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            if (MessageBox.Show("Вы действительно хотите добавить в общий файл все созданные ранее записи?\nВсего записей для экспорта: "+DTR.Count, "Экспот данных в общий файл", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                double This_TH2 = this.Top + this.Height / 2.0;
                double This_LW2 = this.Left + this.Width / 2.0;
                SvodExcel.ProgressBar PB = new SvodExcel.ProgressBar();
                PB.Top = This_TH2 - PB.Height / 2.0;
                PB.Left = This_LW2 - PB.Width / 2.0;
                PB.Topmost = true;
                PB.Show();
                ExportData();
                PB.Close();
            }
            this.Effect = null;
        }
        private void buttonExportHot_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            if (MessageBox.Show("Вы действительно хотите добавить в общий файл все созданные ранее записи?\nВсего записей для экспорта: " + DTR.Count, "Экспот данных в общий файл", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                
                double This_TH2 = this.Top + this.Height / 2.0;
                double This_LW2 = this.Left + this.Width / 2.0;
                    SvodExcel.ProgressBar PB = new SvodExcel.ProgressBar();
                    PB.Top = This_TH2 - PB.Height / 2.0;
                    PB.Left = This_LW2 - PB.Width / 2.0;
                    PB.Topmost = true;
                    PB.Show();
                    ExportData();
                PB.Close();
            }
            this.Effect = null;
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
                    File.Delete(pathC);
                }
                StreamWriter sw = File.CreateText(pathB);
                String host = System.Net.Dns.GetHostName();
                System.Net.IPAddress ip = System.Net.Dns.GetHostEntry(host).AddressList[0];
                sw.WriteLine(ip.ToString());
                sw.Close();
                string pathA = Properties.Settings.Default.PathToGlobalData;
                File.Copy(pathA, pathC);
                var exApp = new Microsoft.Office.Interop.Excel.Application();
                var exBook = exApp.Workbooks.Open(pathC);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                int BlinkEnd = 0;
                var lastcell = ExSheet.Cells.SpecialCells(Type: Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                if (ExSheet.Cells[lastcell.Row, 2].Value != null || ExSheet.Cells[lastcell.Row, 3].Value != null || ExSheet.Cells[lastcell.Row, 4].Value != null || ExSheet.Cells[lastcell.Row, 5].Value != null || ExSheet.Cells[lastcell.Row, 6].Value != null || ExSheet.Cells[lastcell.Row, 7].Value != null)
                    BlinkEnd = 1;
                for (int i= BlinkEnd; i<(DTR.Count+BlinkEnd); i++)
                {
                    ExSheet.Cells[lastcell.Row + i, 2] = DTR[i].Date;
                    ExSheet.Cells[lastcell.Row + i, 3] = DTR[i].Time;
                    ExSheet.Cells[lastcell.Row + i, 4] = DTR[i].Teacher;
                    ExSheet.Cells[lastcell.Row + i, 5] = DTR[i].Group;
                    ExSheet.Cells[lastcell.Row + i, 6] = DTR[i].Category;
                    ExSheet.Cells[lastcell.Row + i, 7] = DTR[i].Place;
                }
                MessageBox.Show(
                    exBook.Path.ToString()+"\\"+exBook.Name.ToString()+ "\n" + (lastcell.Row -1).ToString() + " 4:\n" +
                    ExSheet.Cells[lastcell.Row-1, 4].Value.ToString()
                                       +"\n"+
                    pathC.ToString()+"\n"+(lastcell.Row + DTR.Count - 1).ToString()+" 4:\n"+
                    ExSheet.Cells[lastcell.Row + DTR.Count-1, 4].Value.ToString()
                    );
                exBook.Close(true);
                exApp.Quit();
                File.Replace(pathC,pathA,pathA.Substring(0, pathA.Length-5)+ " "+DateTime.Now.ToString().Replace(':','_')+ ".xlsx");
                File.Delete(pathC);
                File.Delete(pathB);
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
                MessageBox.Show("К сожалению, на данный момент обновление невозможено - другой пользователь уже начал оновлять общий файл!\nПопробуйте еще раз чуть позже");
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
                var exApp = new Microsoft.Office.Interop.Excel.Application();
                var exBook = exApp.Workbooks.Open(pathC);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                string FormulCalculate = ExSheet.Cells[16, 8].Formula;
                exBook.Close(true);
                exApp.Quit();
                File.Delete(pathC);
                //MessageBox.Show(FormulCalculate);
                List<string> TimeTemplate = new List<string>();
                //@"^[А-Я][а-я]*\s[А-Я]\.[А-Я]\.$"
                Regex regex = new Regex(@"\d{1,2}\.\d{2}\-\d{1,2}\.\d{2}");
                MatchCollection matchList = regex.Matches(FormulCalculate);
                for(int i=0;i<matchList.Count;i++)
                {
                    TimeTemplate.Add(matchList[i].Value);
                }
                //TimeTemplate = regex.Matches(FormulCalculate).Val;
                //if (regex.IsMatch(Teacher))
                //List<int> TimeTemplateIndexs = ;
            }
        }
    }
}
