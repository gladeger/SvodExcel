using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Data;

namespace SvodExcel
{
    class DataWork
    {

        public static void UpdateListTimes(Microsoft.Office.Interop.Excel.Application exApp, bool WatchonTimeCreate = false)
        {
            string pathT = @".\ListTime.dat";
            FileInfo localdata;

            string pathB = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
            if (File.Exists(pathB))
            {
                MessageBox.Show("К сожалению обновление списка сейчас невозможно, обновляется общий сводный файл.\nПопробуйте чуть позже.");
            }
            else
            {
                string pathC = Directory.GetCurrentDirectory() + "\\" + Properties.Settings.Default.GlobalData;
                try
                {
                    if (File.Exists(pathC))
                    {
                        localdata = new FileInfo(pathC);
                        localdata.IsReadOnly = false;
                        File.Delete(pathC);
                    }
                }
                catch
                {
                    MessageBox.Show("Ошибка доступа к " + pathC, "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                string pathA = Properties.Settings.Default.PathToGlobalData;
                File.Copy(pathA, pathC);

                var exBook = exApp.Workbooks.Open(pathC);
                var ExSheet = (Microsoft.Office.Interop.Excel.Worksheet)exBook.Sheets[1];
                string FormulCalculate = ExSheet.Cells[16, 8].Formula;
                exBook.Close(true);
                //exApp.Quit();
                localdata = new FileInfo(pathC);
                localdata.IsReadOnly = false;
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
        }

        public static void UpdateTeachersList(Microsoft.Office.Interop.Excel.Application exApp)
        {
            string pathA = Properties.Settings.Default.PathToGlobalData;
            string pathER = Properties.Settings.Default.PathToGlobal + Properties.Settings.Default.GlobalMarker;
            FileInfo localdata;
            if (File.Exists(pathER))
            {
                MessageBox.Show("К сожалению обновление списка сейчас невозможно, обновляется общий сводный файл.\nПопробуйте чуть позже.");
            }
            else
                if (File.Exists(pathA))
            {
                string pathC = Directory.GetCurrentDirectory() + ".\\РАСП.xlsx";
                string pathB = Properties.Settings.Default.PathToLocalData;
                if (File.Exists(pathB))
                {
                    localdata = new FileInfo(pathB);
                    localdata.IsReadOnly = false;
                    File.Delete(pathB);
                }

                File.Copy(pathA, pathB);
                while (!File.Exists(pathB)) { };
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
                localdata = new FileInfo(pathB);
                localdata.IsReadOnly = false;
                File.Delete(pathB);

                ListTeacher.Sort();
                string pathData = @".\ListTeacher.dat";
                File.WriteAllText(pathData, ListTeacher[0]);
                for (int i = 1; i < ListTeacher.Count; i++)
                {
                    File.AppendAllText(pathData, "\n" + ListTeacher[i]);
                }
            }
            else
            {
                MessageBox.Show("Не удалось подключиться к общему сводному файлу!", "Ошибка обновления", MessageBoxButton.OK, MessageBoxImage.Warning, MessageBoxResult.OK);
            }
        }

    }
}
