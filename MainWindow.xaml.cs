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
                Time = inputTime;
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
        }
        private void SvodExcel_Closed(object sender, EventArgs e)
        {
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
        }

        public void EditItem(int RowIndex,DataTableRow newDTR)
        {
            DTR[RowIndex] = newDTR;
            CollectionViewSource.GetDefaultView(dataGridExport.ItemsSource).Refresh();
        }

        private void dataGridExport_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            MessageBox.Show("Edit record");
        }
    }
}
