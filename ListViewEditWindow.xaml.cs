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
using System.Threading;

namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для ListViewEditWindow.xaml
    /// </summary>
    public partial class ListViewEditWindow : Window
    {
        public List<string> LS2;
        public ListViewEditWindow(List<string> LS=null)
        {
            InitializeComponent();
            
            /*if(LS!=null)
            {
                
                LS2 = new List<string>();
                for(int i=0;i<LS.Count;i++)
                {
                    LS2.Add(LS[i]);
                }
                Binding bind = new Binding();
                bind.Source = LS2;
                bind.Path = new PropertyPath(".");
                bind.Mode = BindingMode.OneWay;
                //LVEW.dataGrid.ItemsSource = NoneTeacherTemplate;
                dataGrid.SetBinding(ItemsControl.ItemsSourceProperty, bind);

            }*/
        }

        private void Window_Closed(object sender, EventArgs e)
        {

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.Hide(); 
           e.Cancel = true;
        }

        private void dataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            MessageBox.Show("Ничего не происходит но будет потом");
        }
    }
}
