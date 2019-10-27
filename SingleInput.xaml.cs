using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System;



namespace SvodExcel
{
    /// <summary>
    /// Логика взаимодействия для SingleInput.xaml
    /// </summary>
    public partial class SingleInput : System.Windows.Window
    {
        private string DefaultTimes;
        private bool FlagStartCursorMST = true;
        public SingleInput()
        {
            InitializeComponent();
            buttonConfirmTime.Visibility = Visibility.Hidden;
            labelTimeOut.Visibility = Visibility.Hidden;
            DefaultTimes = MaskedTextBoxStartTime.Text;
            StartListTeacher();
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
            //string path = ".\\РАСП.xlsx";
            //string path = "C:\\Users\\Администратор ОК\\source\\repos\\SvodExcel\\РАСП.xlsx";
            string path = "C:\\Users\\Ilya\\Source\\Repos\\gladeger\\SvodExcel\\РАСП.xlsx";
            //Microsoft.Office.Interop.Excel.XLCel
            var exApp = new Microsoft.Office.Interop.Excel.Application();
            var exBook = exApp.Workbooks.Open(path);
            //if (exBook == null) throw new ArgumentNullException("exBook");
            var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
            var lastcell = ExSheet.Cells.SpecialCells(Type: Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            //string[,] list = new string[lastcell.Row, lastcell.Column];
            List<List<string>> list = new List<List<string>>();
            /*for (int i = 0; i < lastcell.Column; i++) //Все колонки
            {
                list.Add(new List<string>());
                for (int j = 0; j < lastcell.Row; j++) //строки
                    list[i].Add(ExSheet.Cells[j + 1, i + 1].Value.ToString());
            }*/
            // ReSharper disable once CoVariantArrayConversion
            //comboBoxTeacher.Items.AddRange(items: list[0].ToArray());
            exBook.Close(false);
            //exBook.Close(false, Type.Missing, Type.Missing);
            exApp.Quit();
            //GC.Collect();
        }

        private void ButtonUpdate_Click(object sender, RoutedEventArgs e)
        {
            //GetExcel();
            UpdateListTeacher();
        }
        private void UpdateListTeacher()
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
                string[] Teachers=File.ReadAllLines(path);
                comboBoxTeacher.Items.Clear();
                for (int i = 0; i < Teachers.Length;i++)
                {
                    comboBoxTeacher.Items.Add(Teachers[i]);
                }
            }          
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
    }
}
