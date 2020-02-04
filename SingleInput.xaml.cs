using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Linq;
using System.Windows.Media;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;


namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для SingleInput.xaml
    /// </summary>
    public partial class SingleInput : System.Windows.Window
    {
        public int RowIndex;
        private string DefaultTimes="__:__";
        private bool FlagStartCursorMST = true;
        private bool itisclickcombobox = true;
        private bool itisclose = false;
        private List<string> NotCheckTeacher=new List<string>();
        private bool itisnotstart = false;
        List<string> TimeTemplate = new List<string>();
        public Microsoft.Office.Interop.Excel.Application exApp;


        System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
        
        public SingleInput()
        {
            InitializeComponent();
            GridCalcTime.Visibility = Visibility.Hidden;
            DefaultTimes = MaskedTextBoxStartTime.Text;
            StartListTeacher();
            StartListTimes();
            // = new System.Windows.Media.Effects.BlurEffect();     
            objBlur.Radius = 4;
            checkBoxAutoEdit.IsChecked = true;
            NotCheckTeacher.Add(ButtonNewTeacher.Name);
            NotCheckTeacher.Add(buttonUpdate.Name);
            NotCheckTeacher.Add(ButtonWriteAndContinue.Name);
            NotCheckTeacher.Add(ButtonWriteAndStop.Name);
            NotCheckTeacher.Add(ButtonCancel.Name);
            itisnotstart = true;
            DatePicker_Date.Focus();
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
            string BufString = MST.Text;
            if (MST.SelectionStart > 0)
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
            GridTime.Background = null;
            return true;
        }

        private void MaskedTextBoxStartTime_LostFocus(object sender, RoutedEventArgs e)
        {
            if(!itisclose)
            {
                string NewFocusElement = (FocusManager.GetFocusedElement(this) as FrameworkElement).Name;
                if (NewFocusElement != checkBoxAutoEdit.Name)
                {
                    if (checkBoxAutoEdit.IsChecked.Value)
                    {
                        NormMST(MaskedTextBoxStartTime);
                        PositionMST(MaskedTextBoxStartTime, MaskedTextBoxEndTime);
                    }
                }
            }                        
        }
        private void NormMST(Xceed.Wpf.Toolkit.MaskedTextBox MST)
        {
            if (ReadyMST(MST))
            {
                if (System.Convert.ToInt32(MST.Text.Substring(3, 2)) > 59)
                {
                    MST.Text = MST.Text.Substring(0, 2) + ":59";
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
            if (!itisclose)
            {
                string NewFocusElement = (FocusManager.GetFocusedElement(this) as FrameworkElement).Name;
                if (NewFocusElement != checkBoxAutoEdit.Name)
                {
                    if (checkBoxAutoEdit.IsChecked.Value)
                    {
                        NormMST(MaskedTextBoxEndTime);
                        PositionMST(MaskedTextBoxStartTime, MaskedTextBoxEndTime);
                    }
                }
            }                       
        }

        private void GetExcel()
        {
            //string pathA = @"C:\\Users\\Администратор ОК\\source\\repos\\SvodExcel\\РАСП.xlsx";
            string pathA = Properties.Settings.Default.PathToGlobalData;
            //string pathA = @"C:\\Users\\Илья\\Source\\Repos\\SvodExcel\\РАСП.xlsx";
            string pathER = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
            if (File.Exists(pathER))
            {
                MessageBox.Show("К сожалению обновление списка сейчас невозможно, обновляется общий сводный файл.\nПопробуйте чуть позже.");
            }
            else
                if (File.Exists(pathA))
                {
                    string pathC = Directory.GetCurrentDirectory() + ".\\РАСП.xlsx";
                    //string path = "C:\\Users\\Администратор ОК\\source\\repos\\SvodExcel\\РАСП.xlsx";
                    //string pathB = @".\\РАСП.xlsx";
                    string pathB = Properties.Settings.Default.PathToLocalData;
                    if (File.Exists(pathB))
                        File.Delete(pathB);
                    File.Copy(pathA, pathB);
                    while (!File.Exists(pathB)) { };
                //Microsoft.Office.Interop.Excel.XLCel
                /*
                var exBook = exApp.Workbooks.Open(pathC);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                var lastcell = ExSheet.Cells.SpecialCells(Type: Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
                int BlinkEnd = 0;
            if (ExSheet.Cells[lastcell.Row, 2].Value != null || ExSheet.Cells[lastcell.Row, 3].Value != null || ExSheet.Cells[lastcell.Row, 4].Value != null || ExSheet.Cells[lastcell.Row, 5].Value != null || ExSheet.Cells[lastcell.Row, 6].Value != null || ExSheet.Cells[lastcell.Row, 7].Value != null)
                BlinkEnd = 1;
            List<string> ListExcel = new List<string>();    
                for (int j = 15; j < lastcell.Row+BlinkEnd; j++)
                {
                    if (ExSheet.Cells[j + 1, 4].Value != null)
                    {
                        ListExcel.Add(ExSheet.Cells[j + 1, 4].Value.ToString());
                    }
                }
                exBook.Close(false);
                //exApp.Quit();
                */
                String connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathC + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                switch (pathC.Substring(pathC.LastIndexOf('.')))
                {
                    case ".xls":
                        connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathC + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                        break;
                    case ".xlsx":
                        connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathC + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
                        break;
                    default:
                        MessageBox.Show("Ошибка неизвестного формата файла " + pathC.Substring(pathC.LastIndexOf('.')), "Ошибка расширения", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                        break;
                }
                String Command = "Select * from [Лист1$A15:H]";
                OleDbConnection con = new OleDbConnection(connection);

                con.Open();
                OleDbCommand cmd = new OleDbCommand(Command, con);
                OleDbDataAdapter db = new OleDbDataAdapter(cmd);
                DataTable dt_input = new DataTable();
                db.Fill(dt_input);

                for (int i = 0; i < dt_input.Rows.Count; i++)
                {
                    if (dt_input.Rows[i].ItemArray.GetValue(2).ToString().Length == 0 && dt_input.Rows[i].ItemArray.GetValue(3).ToString().Length == 0)
                    {
                        dt_input.Rows[i].Delete();
                        //i -= 1;
                    }

                }
                dt_input.AcceptChanges();
                //dataGridViewFast.ItemsSource = dt_input.AsDataView();

                string BufStringExcel;
                List<string> ListTeacher = new List<string>();
                for (int j = 0; j < dt_input.Rows.Count; j++)
                {
                    BufStringExcel = dt_input.Rows[j].ItemArray.GetValue(3).ToString();
                    if (ListTeacher.IndexOf(BufStringExcel) < 0)
                    {
                        ListTeacher.Add(BufStringExcel);
                    }
                }
                cmd.Dispose();
                con.Close();
                con.Dispose();
                File.Delete(pathB);

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
            MessageBoxResult DR = MessageBox.Show("Сейчас программа попытается обновить список преподавателей. Этот процесс может занять несколько минут. \nВнимание! Новые проподаватели, еще не загруженные в общий доступ, будут удалены.\nПродолжить?", "Начать обновление списка преподавателей", MessageBoxButton.OKCancel, MessageBoxImage.Question, MessageBoxResult.OK);
            if (DR == MessageBoxResult.OK)
            {
                UpdateListTeacher();
            }
            this.Effect = null;
            UpdateLayout();

        }
        private void WinEffectON()
        {
            this.Effect = objBlur;
            UpdateLayout();
        }
        private void UpdateListTeacher()
        {
            /*double This_TH2 = this.Top + this.Height / 2.0;
            double This_LW2 = this.Left + this.Width / 2.0;
            Thread newWindowThread = new Thread(new ThreadStart(() =>
            {
                SvodExcel.ProgressBar PB = new SvodExcel.ProgressBar();
                PB.Top = This_TH2 - PB.Height / 2.0;
                PB.Left = This_LW2 - PB.Width / 2.0;
                PB.Topmost = false;
                PB.ShowDialog();
                System.Windows.Threading.Dispatcher.Run();
            }));
            
            newWindowThread.SetApartmentState(ApartmentState.STA);
            newWindowThread.IsBackground = true;
            newWindowThread.Start();
            */
            System.Windows.Media.Effects.BlurEffect objBlur = new System.Windows.Media.Effects.BlurEffect();
            objBlur.Radius = 4;
            this.Effect = objBlur;
            UpdateLayout();            
            GetExcel();
            this.Effect = null;
            UpdateLayout();
            //PB.Close();
            //newWindowThread.Abort();
            comboBoxTeacher.SelectedIndex = -1;
        }
        private void StartListTeacher()
        {

            string path = @".\ListTeacher.dat";

            if (!File.Exists(path))
            {
                comboBoxTeacher.Items.Add("Moodle");
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
        private void StartListTimes()
        {
            string pathT = @".\ListTime.dat";

            if (!File.Exists(pathT))
            {
                    File.WriteAllText(pathT, "10:00-16:40");
                    File.AppendAllText(pathT, "\n" + "12:00-18:40");
            }

            TimeTemplate.Clear();
            TimeTemplate = File.ReadAllLines(pathT).ToList<string>();
        }
        private void UpdateListTimes()
        {
            string pathT = @".\ListTime.dat";

                string pathB = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
                if (File.Exists(pathB))
                {
                    MessageBox.Show("К сожалению обновление списка сейчас невозможно, обновляется общий сводный файл.\nПопробуйте чуть позже.");
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
                
                    var exBook = exApp.Workbooks.Open(pathC);
                    var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                    string FormulCalculate = ExSheet.Cells[16, 8].Formula;
                    exBook.Close(true);
                    //exApp.Quit();
                    File.Delete(pathC);
                    //MessageBox.Show(FormulCalculate);
                    //@"^[А-Я][а-я]*\s[А-Я]\.[А-Я]\.$"
                    Regex regex = new Regex(@"\d{1,2}\.\d{2}\-\d{1,2}\.\d{2}");
                    MatchCollection matchList = regex.Matches(FormulCalculate);
                    File.WriteAllText(pathT, "10:00-16:40");
                    for (int i = 1; i < matchList.Count; i++)
                    {
                        string tempstring = matchList[i].Value.Replace('.', ':');
                        switch (tempstring.Length)
                        {
                            case 10:
                                if (matchList[i].Value.IndexOf('.') == 2)
                                {
                                    File.AppendAllText(pathT, "\n" + matchList[i].Value.Substring(0, 6).Replace('.', ':') + "0" + matchList[i].Value.Substring(6, 4).Replace('.', ':'));
                                }
                                else
                                {
                                    File.AppendAllText(pathT, "\n" + "0" + matchList[i].Value.Replace('.', ':'));
                                }
                                break;
                            case 9:
                                File.AppendAllText(pathT, "\n" + "0" + matchList[i].Value.Substring(0, 5).Replace('.', ':') + "0" + matchList[i].Value.Substring(5, 4).Replace('.', ':'));
                                break;
                            default:
                                File.AppendAllText(pathT, "\n" + matchList[i].Value.Replace('.', ':'));
                                break;
                        }

                    }
                    //TimeTemplate = regex.Matches(FormulCalculate).Val;
                    //if (regex.IsMatch(Teacher))
                    //List<int> TimeTemplateIndexs = ;
                }

            TimeTemplate.Clear();
            TimeTemplate = File.ReadAllLines(pathT).ToList<string>();
        }
        private void ButtonWriteAndStop_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show("Кнопка пока не работает");
            switch (CorrectData())
            {
                case 0:
                    MessageBox.Show("Ошибка во введенных данных");
                    break;
                case 1:
                    if (ConfirmTime())
                    {
                        NewRecord();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Введенное время занятия не соответсвует возможным диапазонам.\nПроверьте еще раз времена начала и завершения занятий");
                    }
                    
                    break;
                case 2:
                    MessageBox.Show("Были внесены корректировки записи, убедитесь что новые данные действительны");
                    break;
                case 3:
                    MessageBox.Show("Были внесены корректировки в записи, убедитесь что новые данные действительны. Однако не все данные удалось исправить.");
                    break;
                default:
                    MessageBox.Show("Неизвестная ошибка");
                    break;
            }
        }

        private int CorrectData()
        {
            int flag = 1;
            int flag_time = 1;
            int flag_date = 1;
            int flag_teacher = 1;
            if (!ReadyMST(MaskedTextBoxStartTime) || !ReadyMST(MaskedTextBoxEndTime))
            {
                flag = 0;
                flag_time = 0;
            }
            else
            {
                if(System.Convert.ToInt32(MaskedTextBoxStartTime.Text.Substring(3, 2))>59)
                {
                    flag = 0;
                    flag_time = 0;
                }
                if (System.Convert.ToInt32(MaskedTextBoxEndTime.Text.Substring(3, 2)) > 59)
                {
                    flag = 0;
                    flag_time = 0;
                }
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
                    if (flag_time == 0)
                    {
                        flag_time = 2;
                        flag = 2;
                    }
                }
            }
            if (DatePicker_Date.Text.Length == 0)
            {
                flag = 0;
                GridDate.Background = new SolidColorBrush(Colors.Red);
                DatePicker_Date.Focus();
                flag_date = 0;
            }
            else
            {
                GridDate.Background = null;
            }
            switch (flag_time)
            {
                case 0:
                    GridTime.Background = new SolidColorBrush(Colors.Red);
                    if (flag == 2)
                    {
                        flag = 3;
                    }
                    MaskedTextBoxStartTime.Focus();
                    break;
                case 2:
                    GridTime.Background = new SolidColorBrush(Colors.Yellow);
                    if (flag == 0)
                    {
                        flag = 3;
                    }
                    MaskedTextBoxStartTime.Focus();
                    break;
                default:
                    GridTime.Background = null;
                    break;
            }
            if(!CorrectAndAddTeacher()||comboBoxTeacher.Text.Length<1)
            {
                if (flag == 2)
                    flag = 3;
                else
                    flag = 0;
                GridTeacher.Background = new SolidColorBrush(Colors.Red);
                comboBoxTeacher.Focus();
                flag_teacher = 0;
            }
            else
            {
                GridTeacher.Background = null;
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
                    if(ConfirmTime())
                    {
                        NewRecord();
                        ClearData();
                    }
                    else
                    {
                        MessageBox.Show("Введенное время занятия не соответсвует возможным диапазонам.\nПроверьте еще раз времена начала и завершения занятий");
                    }
                    break;
                case 2:
                    MessageBox.Show("Были внесены корректировки записи, убедитесь что новые данные действительны");
                    break;
                case 3:
                    MessageBox.Show("Были внесены корректировки в записи, убедитесь что новые данные действительны. Однако не все данные удалось исправить.");
                    break;
                default:
                    MessageBox.Show("Неизвестная ошибка");
                    break;
            }
        }

        private void comboBoxTeacher_LostFocus(object sender, RoutedEventArgs e)
        {
           
                if(!itisclose)
                {
                string NewFocusElement = (FocusManager.GetFocusedElement(this) as FrameworkElement).Name;
                if(NotCheckTeacher.IndexOf(NewFocusElement)<0)
                        CorrectAndAddTeacher(!itisclickcombobox);
                }
  
        }

        private bool CorrectAndAddTeacher(bool silence=false)
        {
            if (comboBoxTeacher.Items.IndexOf(comboBoxTeacher.Text) < 0)
            {
                if(comboBoxTeacher.SelectedIndex >= 0)
                {
                    comboBoxTeacher.Text = comboBoxTeacher.SelectedValue.ToString();
                    ButtonNewTeacher.IsEnabled = false;
                }
                else
                {
                    if (comboBoxTeacher.Text.Length > 0)
                    {
                        MessageBoxResult DR= MessageBoxResult.No;
                        if(!silence)
                            DR = MessageBox.Show("Преподавателя \""+comboBoxTeacher.Text+"\" нет в списке преподавателей. Добавить нового преподавателя?", "Новый преподаватель", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
                        if (DR == MessageBoxResult.Yes)
                        {
                            
                            if(CorrectNewTeacher(comboBoxTeacher.Text))
                            {
                                if (NewTeacher(comboBoxTeacher.Text))
                                    MessageBox.Show("Запись нового преподавателя успеешно завершена.\nНо другие пользователи не увидят нового преподавателя, пока не будут сделаны новые записи в общее расписание.");
                                else
                                    MessageBox.Show("Ошибка записи нового преподавателя");
                            }
                            else
                            {
                                MessageBox.Show("Строка \"" + comboBoxTeacher.Text + "\" не удовлетворяет формату записи преподавателя - Фамилия и инициалы.\nК примеру, Иванов И.И.\nФИО должно записываться только из букв русского алфавита, пробела и символа точки.");
                                return false;
                            }                      
                        }
                        else
                        {
                            ButtonNewTeacher.IsEnabled = true;
                            return false;
                        }
                    }
                }           
            }
            else
            {
                ButtonNewTeacher.IsEnabled = false;
            }
            GridTeacher.Background = null;
            return true;
        }

        private void Single_manual_entry_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GC.Collect();
            itisclickcombobox = false;
            itisclose = true;
        }
        
        private bool CorrectNewTeacher(string Teacher)
        {
            Regex regex = new Regex(@"^[А-Я][а-я]*\s[А-Я]\.[А-Я]\.$");
            if(regex.IsMatch(Teacher))
                return true;
            return false;
        }
        private bool NewTeacher(string Teacher)
        {
            string path = @".\ListTeacher.dat";

            if (!File.Exists(path))
            {
                return false;
            }
            else
            {
                try
                {
                    List<string> Teachers = new List<string>(File.ReadAllLines(path));
                    Teachers.Add(Teacher);
                    Teachers.Sort();
                    comboBoxTeacher.Items.Clear();
                    File.WriteAllText(path, Teachers[0]);
                    for (int i = 1; i < Teachers.Count; i++)
                    {
                        File.AppendAllText(path, "\n" + Teachers[i]);
                    }
                    for (int i = 0; i < Teachers.Count; i++)
                    {
                        comboBoxTeacher.Items.Add(Teachers[i]);
                    }
                }
                catch
                {
                    return false;
                }                
            }
            return true;
        }

        private void comboBoxTeacher_MouseEnter(object sender, MouseEventArgs e)
        {
            itisclickcombobox = false;
            
        }

        private void comboBoxTeacher_MouseLeave(object sender, MouseEventArgs e)
        {
            itisclickcombobox = true;
            
        }

        private void comboBoxTeacher_GotFocus(object sender, RoutedEventArgs e)
        { 
            itisclickcombobox = true;
        }

        private void CheckBoxAutoEdit_Checked(object sender, RoutedEventArgs e)
        {
            NormMST(MaskedTextBoxStartTime);
            NormMST(MaskedTextBoxEndTime);
            PositionMST(MaskedTextBoxStartTime, MaskedTextBoxEndTime);
        }

        private void MaskedTextBoxStartTime_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (itisnotstart)
                OutCalcTime();
        }

        private void MaskedTextBoxEndTime_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if(itisnotstart)
                OutCalcTime();
        }

        private void OutCalcTime()
        {
            Xceed.Wpf.Toolkit.MaskedTextBox MSTS=MaskedTextBoxStartTime, MSTE=MaskedTextBoxEndTime;
            if (ReadyMST(MSTS)&& ReadyMST(MSTE))
            {
                GridCalcTime.Visibility = Visibility.Visible;
                labelCalcTime.Content=((System.Convert.ToInt32(MSTE.Text.Substring(0, 2)) * 60 + System.Convert.ToInt32(MSTE.Text.Substring(3, 2))) - (System.Convert.ToInt32(MSTS.Text.Substring(0, 2)) * 60 + System.Convert.ToInt32(MSTS.Text.Substring(3, 2)))).ToString();
                labelCalcTime.Content = labelCalcTime.Content + " мин.";

            }
            else
            {
                GridCalcTime.Visibility = Visibility.Hidden;
            }
        }

        private void ButtonNewTeacher_Click(object sender, RoutedEventArgs e)
        {
            CorrectAndAddTeacher();
        }

        private void NewRecord()
        {
            MainWindow home = Application.Current.MainWindow as MainWindow;
            if(RowIndex==-1)
                home.AddNewItem(new DataTableRow(DatePicker_Date.Text, MaskedTextBoxStartTime.Text + "-" + MaskedTextBoxEndTime.Text, comboBoxTeacher.Text, textboxGroup.Text, textBoxCategory.Text, textBoxPlace.Text));
            else
            {
                home.EditItem(RowIndex,new DataTableRow(DatePicker_Date.Text, MaskedTextBoxStartTime.Text + "-" + MaskedTextBoxEndTime.Text, comboBoxTeacher.Text, textboxGroup.Text, textBoxCategory.Text, textBoxPlace.Text));
            }
        }
        private void ClearData()
        {
            DatePicker_Date.Text = null;
            MaskedTextBoxStartTime.Text = DefaultTimes;
            MaskedTextBoxEndTime.Text = DefaultTimes;
            GridCalcTime.Visibility = Visibility.Hidden;
            comboBoxTeacher.Text = "";
            comboBoxTeacher.SelectedIndex = -1;
            ButtonNewTeacher.IsEnabled = false;
            textBoxCategory.Text = null;
            textBoxPlace.Text = null;
        }

        private void DatePicker_Date_GotFocus(object sender, RoutedEventArgs e)
        {
            GridDate.Background = null;
        }

        private void buttonUpdateTimeTemplates_Click(object sender, RoutedEventArgs e)
        {
            WinEffectON();
            MessageBoxResult DR = MessageBox.Show("Сейчас программа попытается обновить список возможного времени занятия. Этот процесс может занять несколько минут. \nВнимание! Новые проподаватели, еще не загруженные в общий доступ, будут удалены.\nПродолжить?", "Начать обновление списка преподавателей", MessageBoxButton.OKCancel, MessageBoxImage.Question, MessageBoxResult.OK);
            if (DR == MessageBoxResult.OK)
            {
                UpdateListTimes();
            }
            this.Effect = null;
            UpdateLayout();

        }
        public bool ConfirmTime()
        {
            if(TimeTemplate.IndexOf(MaskedTextBoxStartTime.Text+"-"+ MaskedTextBoxEndTime.Text)>=0)
            {
                return true;
            }
            return false;
        }

        private void buttonViewTimeTemplates_Click(object sender, RoutedEventArgs e)
        {
            string bufstr = "Диапазоны времени:\n";
            for(int i=0;i<TimeTemplate.Count;i++)
            {
                bufstr += TimeTemplate[i]+"\n";
            }
            MessageBox.Show(bufstr,"Возможные диапазоны времени",MessageBoxButton.OK);
        }
    }
}
