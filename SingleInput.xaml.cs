using System.Windows;
using System.Windows.Input;


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
            
            //MaskedTextBoxStartTime.Select(0, 0);
           // MaskedTextBoxStartTime.SelectionStart = 0;
            //MaskedTextBoxStartTime.CaretIndex = 0;
           // MaskedTextBoxStartTime.ScrollToHome();
            //labelTimeOut.Content = MaskedTextBoxStartTime.CaretIndex.ToString();



        }

        private void MaskedTextBoxStartTime_SelectionChanged(object sender, RoutedEventArgs e)
        {
            if(MaskedTextBoxStartTime.CaretIndex!=0)
                labelTimeOut.Content = MaskedTextBoxStartTime.CaretIndex.ToString();
        }
    }
}
