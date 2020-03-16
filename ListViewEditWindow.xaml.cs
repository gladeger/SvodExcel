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
using System.Collections.ObjectModel;
using System.Collections.Specialized;

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
            //CollectionViewSource.GetDefaultView(dataGrid.ItemsSource).Refresh();
            SaveEdit();
        }

       

        private void buttonEditInputHot_Click(object sender, RoutedEventArgs e)
        {
            dataGrid.Focus();
            dataGrid.BeginEdit() ;
            dataGrid.CommitEdit();
            MessageBox.Show("Простите, сервис временно не работает по этой кнопке.\nЕсли хотите редактировать запись в списке, сдеолайте по ней двойной клик или клавишу F2.");            
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(dataGrid.Items.Count>1)
            {
                if (!dataGrid.IsReadOnly)
                {
                    buttonEditInputHot.IsEnabled = true;
                    buttonDeleteHot.IsEnabled = true;
                }
            }
            else
            {
                buttonEditInputHot.IsEnabled = false;
                //buttonDeleteHot.IsEnabled = false;
            }
            //labelTech.Content = dataGrid.Items.Count.ToString();
        }

        private void MenuItemSingleInput_Click(object sender, RoutedEventArgs e)
        {
            dataGrid.SelectedIndex = dataGrid.Items.Count - 1;
            dataGrid.Focus();
        }

        private void buttonDeleteHot_Click(object sender, RoutedEventArgs e)
        {
            (dataGrid.ItemsSource as List<NoneTeacher>).Remove(dataGrid.SelectedItem as NoneTeacher);
            CollectionViewSource.GetDefaultView(dataGrid.ItemsSource).Refresh();

        }

        public void SaveEdit() 
        {
            switch (Owner.GetType().ToString())
            {
                case ("SvodExcel.MainWindow"):
                    break;
                case ("SvodExcel.OpenFileTable"):
                    (Owner as OpenFileTable).SaveDataListViewEditWindow();
                    break;
                default:
                    break;
            }
                    
        }

        private void dataGrid_UnloadingRow(object sender, DataGridRowEventArgs e)
        {
            SaveEdit();
        }

        private void dataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            //CollectionViewSource.GetDefaultView(dataGrid.ItemsSource).Refresh();
            SaveEdit();
        }
    }
}
