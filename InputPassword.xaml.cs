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
    /// Логика взаимодействия для InputPassword.xaml
    /// </summary>
    public partial class InputPassword : Window
    {
        public InputPassword()
        {
            InitializeComponent();
            pasBox.Focus();
        }

        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
            if(pasBox.Password== Properties.Settings.Default.AdminPassword)
            {
                this.DialogResult = true;
            }
            else
            {
                MessageBox.Show("Пароль неверный!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
//            this.Close();
        }
    }
}
