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


namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для SingleInput.xaml
    /// </summary>
    public partial class SingleInput : Window
    {
        public SingleInput()
        {
            InitializeComponent();
            buttonConfirmTime.Visibility = Visibility.Hidden;
            //labelTimeOut.Visibility = Visibility.Hidden;
        }

        private void MaskedTextBoxStartTime_GotFocus(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show("!");
            
            MaskedTextBoxStartTime.Select(1, 1);
            MaskedTextBoxStartTime.SelectionStart = 1;
            MaskedTextBoxStartTime.CaretIndex = 1;
            labelTimeOut.Content = MaskedTextBoxStartTime.CaretIndex.ToString();


        }
    }
}
