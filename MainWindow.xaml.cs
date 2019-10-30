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

        }
        public MainWindow()
        {
            InitializeComponent();
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
            f.ShowDialog();
        }

        private void SvodExcel_Closed(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }
    }
}
