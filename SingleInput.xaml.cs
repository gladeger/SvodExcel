using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Linq;
using System.Windows.Media;



namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для SingleInput.xaml
    /// </summary>
    public partial class SingleInput : System.Windows.Window
    {
        private string DefaultTimes;
        private bool FlagStartCursorMST = true;
        private bool itisclickcombobox = true;
        public SingleInput()
        {
            InitializeComponent();
            buttonConfirmTime.Visibility = Visibility.Hidden;
            DefaultTimes = MaskedTextBoxStartTime.Text;
            StartListTeacher();
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
        }

        private void MaskedTextBoxStartTime_GotFocus(object sender, RoutedEventArgs e)
        {
            StartCursorMST(MaskedTextBoxStartTime);
            FlagStartCursorMST = true;
        }

        private void MaskedTextBoxStartTime_SelectionChanged(object sender, RoutedEventArgs e)
        {
            if (FlagStartCursorMST)
            {
                FlagStartCursorMST = false;
                StartCursorMST(MaskedTextBoxStartTime);
            }
                
        }
        private void StartCursorMST(Xceed.Wpf.Toolkit.MaskedTextBox MST)
        {
            int IC = 0;
            string BufString= MST.Text;
            if(MST.SelectionStart>0)
            {
                if (BufString[0] == DefaultTimes[0])
                    IC = 0;
                else
                    if (BufString[1] == DefaultTimes[1])
                        IC = 1;
                    else
                        if (BufString[3] == DefaultTimes[3])
                        IC = 3;
                    else
                            if (BufString[4] == DefaultTimes[4])
                        IC = 4;
                MST.Select(IC, 0);
                MST.CaretIndex = IC;
            }
        }

        private void MaskedTextBoxEndTime_GotFocus(object sender, RoutedEventArgs e)
        {
            StartCursorMST(MaskedTextBoxEndTime);
            FlagStartCursorMST = true;
        }

        private void MaskedTextBoxEndTime_SelectionChanged(object sender, RoutedEventArgs e)
        {
            if (FlagStartCursorMST)
            {
                FlagStartCursorMST = false;
                StartCursorMST(MaskedTextBoxEndTime);
            }
        }
        private bool ReadyMST(Xceed.Wpf.Toolkit.MaskedTextBox MST)
        {
            string BufString = MST.Text;
            if (BufString[0] == DefaultTimes[0])
                return false;
            else
                if (BufString[1] == DefaultTimes[1])
                return false;
            else
                    if (BufString[3] == DefaultTimes[3])
                return false;
            else
                        if (BufString[4] == DefaultTimes[4])
                return false;
            return true;
        }

        private void MaskedTextBoxStartTime_LostFocus(object sender, RoutedEventArgs e)
        {
            if(checkBoxAutoEdit.IsChecked.Value)
            {
                NormMST(MaskedTextBoxStartTime);
                PositionMST(MaskedTextBoxStartTime, MaskedTextBoxEndTime);
            }
        }
        private void NormMST(Xceed.Wpf.Toolkit.MaskedTextBox MST)
        {
            if (ReadyMST(MST))
            {
                if(System.Convert.ToInt32(MST.Text.Substring(3, 2))>59)
                {
                    MST.Text = MST.Text.Substring(0,2)+":59";
                }
                if (System.Convert.ToInt32(MST.Text.Substring(0, 2) + MST.Text.Substring(3, 2)) < 840)
                {
                    MST.Text = "08:40";
                }
                else
                {
                    if (System.Convert.ToInt32(MST.Text.Substring(0, 2) + MST.Text.Substring(3, 2)) > 2100)
                    {
                        MST.Text = "21:00";
                    }
                }
            }
        }

        private void PositionMST(Xceed.Wpf.Toolkit.MaskedTextBox MSTS, Xceed.Wpf.Toolkit.MaskedTextBox MSTE)
        {
            if (ReadyMST(MSTS) && ReadyMST(MSTE))
            {
                if (System.Convert.ToInt32(MSTS.Text.Substring(0, 2) + MSTS.Text.Substring(3, 2)) > System.Convert.ToInt32(MSTE.Text.Substring(0, 2) + MSTE.Text.Substring(3, 2)))
                {
                    string BufString = MSTS.Text;
                    MSTS.Text = MSTE.Text;
                    MSTE.Text = BufString;
                }
            }
        }

        private void MaskedTextBoxEndTime_LostFocus(object sender, RoutedEventArgs e)
        {
            if (checkBoxAutoEdit.IsChecked.Value)
            {
                NormMST(MaskedTextBoxEndTime);
                PositionMST(MaskedTextBoxStartTime, MaskedTextBoxEndTime);
            }            
        }

        private void GetExcel()
        {
            string pathA = @"C:\\Users\\Администратор ОК\\source\\repos\\SvodExcel\\РАСП.xlsx";
            if(File.Exists(pathA))
            {
                ;
                string path = Directory.GetCurrentDirectory()+".\\РАСП.xlsx";
                //string path = "C:\\Users\\Администратор ОК\\source\\repos\\SvodExcel\\РАСП.xlsx";
                string pathB = @".\\РАСП.xlsx";
                if (File.Exists(pathB))
                    File.Delete(pathB);
                File.Copy(pathA,pathB);
                while (!File.Exists(pathB)) { };
                //string path = "C:\\Users\\Ilya\\Source\\Repos\\gladeger\\SvodExcel\\РАСП.xlsx";
                //Microsoft.Office.Interop.Excel.XLCel
                var exApp = new Microsoft.Office.Interop.Excel.Application();
                var exBook = exApp.Workbooks.Open(path);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                var lastcell = ExSheet.Cells.SpecialCells(Type: Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                List<string> ListExcel = new List<string>();
                for (int j = 15; j < lastcell.Row; j++)
                {
                    if (ExSheet.Cells[j + 1, 4].Value != null)
                    {
                        ListExcel.Add(ExSheet.Cells[j + 1, 4].Value.ToString());
                    }
                }
                exBook.Close(false);
                exApp.Quit();
                File.Delete(pathB);

                List<string> ListTeacher = new List<string>(ListExcel.Distinct());
                ListTeacher.Sort();
                string pathData = @".\ListTeacher.dat";
                File.WriteAllText(pathData, ListTeacher[0]);
                for (int i = 1; i < ListTeacher.Count; i++)
                {
                    File.AppendAllText(pathData, "\n" + ListTeacher[i]);
                }
                StartListTeacher();
            }
            else
            {
                MessageBox.Show("Не удалось подключиться к общему сводному файлу!", "Ошибка обновления", MessageBoxButton.OK, MessageBoxImage.Warning, MessageBoxResult.OK);
            }

        }
        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinEffectON();
            MessageBoxResult DR = MessageBox.Show("Сейчас программа попытается обновить список преподавателей. Этот процесс может занять несколько минут. Продолжить?", "Начать обновление списка преподавателей", MessageBoxButton.OKCancel,MessageBoxImage.Question,MessageBoxResult.OK);
            if(DR== MessageBoxResult.OK)
            {
                UpdateListTeacher();
            }            
            this.Effect = null;
            
        }
        private void WinEffectON()
        {
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
        }
        private void UpdateListTeacher()
        {            
            double This_TH2 = this.Top + this.Height / 2.0;
            double This_LW2 = this.Left + this.Width / 2.0;
            Thread newWindowThread = new Thread(new ThreadStart(() =>
            {
                SvodExcel.ProgressBar PB = new SvodExcel.ProgressBar();
                PB.Top = This_TH2 - PB.Height / 2.0;
                PB.Left = This_LW2 - PB.Width / 2.0;
                PB.Topmost = true;
                PB.Show();
                System.Windows.Threading.Dispatcher.Run();
            }));
            newWindowThread.SetApartmentState(ApartmentState.STA);
            newWindowThread.IsBackground = true;
            newWindowThread.Start();
            GetExcel();
            //PB.Close();
            newWindowThread.Abort();
        }
        private void StartListTeacher()
        {

            string path = @".\ListTeacher.dat";

            if (!File.Exists(path))
            {
                comboBoxTeacher.Items.Add("Пронина Л.Н.");
                comboBoxTeacher.Items.Add("Григорьева А.И.");
                File.WriteAllText(path, comboBoxTeacher.Items[0].ToString());
                for (int i = 1; i < comboBoxTeacher.Items.Count; i++)
                {
                    File.AppendAllText(path, "\n" + comboBoxTeacher.Items[i].ToString());
                }
            }
            else
            {
                string[] Teachers = File.ReadAllLines(path);
                comboBoxTeacher.Items.Clear();
                for (int i = 0; i < Teachers.Length; i++)
                {
                    comboBoxTeacher.Items.Add(Teachers[i]);
                }
            }
        }

        private void ButtonWriteAndStop_Click(object sender, RoutedEventArgs e)
        {
            switch(CorrectData())
            {
                case 0: MessageBox.Show("Ошибка во введенных данных");
                    break;
                case 1: MessageBox.Show("Пошла запись");
                    break;
                case 2: MessageBox.Show("Были внесены корректировки записи, Убедитесь что новые данные действительны");
                    break;
                default: MessageBox.Show("Неизвестная ошибка");
                    break;
            }
            this.Close();
        }

        private int CorrectData()
        {
            int flag = 1;
            int flag_time = 1;
            if (!ReadyMST(MaskedTextBoxStartTime) || !ReadyMST(MaskedTextBoxEndTime))
            {
                flag = 0;
                flag_time = 0;
            }
            else
            {
                    if (System.Convert.ToInt32(MaskedTextBoxStartTime.Text.Substring(0, 2) + MaskedTextBoxStartTime.Text.Substring(3, 2)) < 840)
                    {
                        flag = 0;
                        flag_time = 0;
                    }
                    else
                    {
                        if (System.Convert.ToInt32(MaskedTextBoxStartTime.Text.Substring(0, 2) + MaskedTextBoxStartTime.Text.Substring(3, 2)) > 2100)
                        {
                            flag = 0;
                            flag_time = 0;
                        }
                    }
                    if (System.Convert.ToInt32(MaskedTextBoxEndTime.Text.Substring(0, 2) + MaskedTextBoxEndTime.Text.Substring(3, 2)) < 840)
                    {
                        flag = 0;
                        flag_time = 0;
                    }
                    else
                    {
                        if (System.Convert.ToInt32(MaskedTextBoxEndTime.Text.Substring(0, 2) + MaskedTextBoxEndTime.Text.Substring(3, 2)) > 2100)
                        {
                            flag = 0;
                            flag_time = 0;
                        }
                    }
                    if (System.Convert.ToInt32(MaskedTextBoxStartTime.Text.Substring(0, 2) + MaskedTextBoxStartTime.Text.Substring(3, 2)) > System.Convert.ToInt32(MaskedTextBoxEndTime.Text.Substring(0, 2) + MaskedTextBoxEndTime.Text.Substring(3, 2)))
                    {
                        flag = 0;
                        flag_time = 0;
                    }

                if (checkBoxAutoEdit.IsChecked.Value)
                {
                    NormMST(MaskedTextBoxEndTime);
                    NormMST(MaskedTextBoxStartTime);
                    PositionMST(MaskedTextBoxStartTime, MaskedTextBoxEndTime);
                    if(flag_time == 0)
                    {
                        flag_time = 2;
                        flag = 2;
                    }
                }
            }

            if(DatePicker_Date.Text.Length==0)
            {
                flag = 2;
                GridDate.Background = new SolidColorBrush(Colors.Red);
            }
            else
            {
                GridDate.Background = null;
            }
            switch(flag_time)
            {
                case 0: GridTime.Background = new SolidColorBrush(Colors.Red);
                    break;
                case 2: GridTime.Background = new SolidColorBrush(Colors.Yellow);
                    break;
                default: GridTime.Background = null;
                    break;
            }
            return flag;
        }

        private void ButtonWriteAndContinue_Click(object sender, RoutedEventArgs e)
        {
            switch (CorrectData())
            {
                case 0:
                    MessageBox.Show("Ошибка во введенных данных");
                    break;
                case 1:
                    MessageBox.Show("Пошла запись");
                    break;
                case 2:
                    MessageBox.Show("Были внесены корректировки записи, Убедитесь что новые данные действительны");
                    break;
                default:
                    MessageBox.Show("Неизвестная ошибка");
                    break;
            }
        }

        private void comboBoxTeacher_LostFocus(object sender, RoutedEventArgs e)
        {
            if(itisclickcombobox)
                if(CorrectTeacher())
                {

                }
        }

        private bool CorrectTeacher()
        {
            if(comboBoxTeacher.SelectedIndex<0)
            {
                ButtonNewTeacher.IsEnabled = true;
                if (comboBoxTeacher.Text.Length>0)
                {
                    MessageBoxResult DR = MessageBox.Show("Указанного преподавателя нет в списке преподавателей. Добавить нового преподавателя?", "Новый преподаватель", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.OK);
                    if (DR==MessageBoxResult.Yes)
                    {
                        MessageBox.Show("Добавляем преподавателя");
                    }
                }
            }
            else
            {
                ButtonNewTeacher.IsEnabled = false;
            }
            
            //MessageBox.Show("!");
            return true;
        }

        private void Single_manual_entry_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if(comboBoxTeacher.IsFocused)
            {
                itisclickcombobox = false;
            }
        }
    }
}
