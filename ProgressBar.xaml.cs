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
    /// Логика взаимодействия для ProgressBar.xaml
    /// </summary>
    public partial class ProgressBar : Window
    {
        bool flagActive = true;
        public ProgressBar()
        {
            InitializeComponent();
            flagActive = true;
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            /*
            if (flagActive)
            {
                (this.Parent as Window).Activate();
                flagActive = false;
            }
            else
                flagActive = true;
                */
        }
    }
}
