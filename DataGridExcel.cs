using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Windows;
using System.Text.RegularExpressions;

namespace SvodExcel
{
    public class DataTableRow
    {
        
        public string Date { get; set; }
        public string Time { get; set; }
        public string Teacher { get; set; }
        public string Group { get; set; }
        public string Category { get; set; }
        public string Place { get; set; }
        
        public DataTableRow(string inputDate, string inputTime, string inputTeacher, string inputGroup, string inputCategory, string inputPlace, bool testcorrect=false)
        {
            // DataTableRow("06.11.2019", "10:00-16:40", "Пронина Л.Н.", "","-----","-----"));
            Date = "";
            Time = "";
            Teacher = "";
            Group = "";
            Category = "";
            Place = "";
            if (testcorrect)
            {
                string bufInputDate=inputDate;
                if(bufInputDate.Length>0)
                {
                    while(bufInputDate[0]<'0'|| bufInputDate[0] > '9')
                    {
                        bufInputDate = bufInputDate.Substring(1);
                    }
                    Regex regexDateDigit = new Regex(@"\d");
                    Regex regexDateSeparate = new Regex(@"[\.,:;\-\/ ]");
                    MatchCollection matchesDate = regexDateDigit.Matches(bufInputDate);
                    MatchCollection matchesDateSeparate = regexDateSeparate.Matches(bufInputDate);
                    bufInputDate.Substring(0, matchesDate[matchesDate.Count - 1].Index + 1);
                    while (bufInputDate[bufInputDate.Length-1] < '0' || bufInputDate[bufInputDate.Length - 1] > '9')
                    {
                        bufInputDate = bufInputDate.Substring(0, bufInputDate.Length - 1);
                    }
                    matchesDate = regexDateDigit.Matches(bufInputDate);
                    matchesDateSeparate = regexDateSeparate.Matches(bufInputDate);
                    bufInputDate =regexDateSeparate.Replace(bufInputDate, "/");
                    //MessageBox.Show(Convert.ToDateTime(bufInputDate).ToString());
                    MessageBox.Show(bufInputDate);
                    switch (matchesDate.Count)
                    {
                        case 8:
                            switch (matchesDateSeparate[0].Index)
                            {
                                case 2:
                                    Date = matchesDate[0].Value + matchesDate[1].Value + "." + matchesDate[2].Value + matchesDate[3].Value + "." + matchesDate[4].Value + matchesDate[5].Value + matchesDate[6].Value + matchesDate[7].Value;
                                    break;
                                case 4:
                                    Date = matchesDate[6].Value + matchesDate[7].Value + "." + matchesDate[4].Value + matchesDate[5].Value + "." + matchesDate[0].Value + matchesDate[1].Value + matchesDate[2].Value + matchesDate[3].Value;
                                    break;
                                default:
                                    Date = "";
                                    break;
                            }
                            
                            break;
                        case 7:
                            switch (matchesDateSeparate[0].Index)
                            {
                                case 1:
                                    Date = "0" + matchesDate[0].Value + "." + matchesDate[1].Value + matchesDate[2].Value + "." + matchesDate[3].Value + matchesDate[4].Value + matchesDate[5].Value + matchesDate[6].Value;
                                    break;

                            }
                            break;
                        case 6:
                            break;
                        case 5:
                            break;
                        case 4:
                            break;
                        default:
                            Date = "";
                            break;

                    }
                }                
            }
            else
            {
                Date = inputDate;
                if (inputTime.Length > 0)
                {
                    if (inputTime[0] == '0')
                    {
                        Time = inputTime.Substring(1).Replace(':', '.');
                    }
                    else
                    {
                        Time = inputTime.Replace(':', '.');
                    }
                    if (Time[Time.IndexOf("-") + 1] == '0')
                    {
                        Time = Time.Substring(0, Time.IndexOf("-") + 1) + Time.Substring(Time.IndexOf("-") + 2);
                    }
                }
                else
                {
                    Time = null;
                }

                Teacher = inputTeacher;
                Group = inputGroup;
                Category = inputCategory;
                Place = inputPlace;
            }
            
        }
        public DataTableRow()
        {
            Date = null;
            Time = null;
            Teacher = null;
            Group = null;
            Category = null;
            Place = null;
        }
    }
    public class DataViewTableRow
    {
        public string Date { get; set; }
        public string Time { get; set; }
        public string Teacher { get; set; }
        public string Group { get; set; }
        public string Category { get; set; }
        public string Place { get; set; }
        public string Result { get; set; }

        public DataViewTableRow(string inputDate, string inputTime, string inputTeacher, string inputGroup, string inputCategory, string inputPlace, string inputResult = null)
        {
            Date = inputDate;
            if (inputTime.Length > 0)
            {
                if (inputTime[0] == '0')
                {
                    Time = inputTime.Substring(1).Replace(':', '.');
                }
                else
                {
                    Time = inputTime.Replace(':', '.');
                }
                if (Time[Time.IndexOf("-") + 1] == '0')
                {
                    Time = Time.Substring(0, Time.IndexOf("-") + 1) + Time.Substring(Time.IndexOf("-") + 2);
                }
            }
            else
            {
                Time = null;
            }

            Teacher = inputTeacher;
            Group = inputGroup;
            Category = inputCategory;
            Place = inputPlace;
            Result = inputResult;
        }
        public DataViewTableRow()
        {
            Date = null;
            Time = null;
            Teacher = null;
            Group = null;
            Category = null;
            Place = null;
            Result = null;
        }
        public DataViewTableRow(DataTableRow InputData)
        {
            Date = InputData.Date;
            Time = InputData.Time;
            Teacher = InputData.Teacher;
            Group = InputData.Group;
            Category = InputData.Category;
            Place = InputData.Place;
            Result = null;
        }
    }

    public class DataViewFastTableRow
    {
        public string Teacher { get; set; }
        public string Result { get; set; }

        public DataViewFastTableRow(string inputTeacher, string inputResult = null)
        {
            Teacher = inputTeacher;
            Result = inputResult;
        }
        public DataViewFastTableRow()
        {
            Teacher = null;
            Result = null;
        }
        public DataViewFastTableRow(DataTableRow InputData)
        {
            Teacher = InputData.Teacher;
            Result = null;
        }
    }


    public class InputDataFile
    {
        public List<DataTableRow> InputDataFileRows { get; }
        public InputDataFile()
        {
            InputDataFileRows = new List<DataTableRow>();
            InputDataFileRows.Clear();
        }
        public InputDataFile(string FileName)
        {
            InputDataFileRows = new List<DataTableRow>();
            OpenFile(FileName);
        }
        public InputDataFile(DataTableRow inputdatafilerow)
        {
            InputDataFileRows = new List<DataTableRow>();
            InputDataFileRows.Clear();
            InputDataFileRows.Add(inputdatafilerow);
        }

        public void Clear()
        {
            InputDataFileRows.Clear();
        }

        public void Add(DataTableRow inputdatafilerow)
        {
            InputDataFileRows.Add(inputdatafilerow);
        }

        public void Add(String str1="", String str2="", String str3="", String str4="", String str5="", String str6="")
        {
            if(str1==null) str1 = "";
            if (str2 == null) str2 = "";
            if (str3 == null) str3 = "";
            if (str4 == null) str4 = "";
            if (str5 == null) str5 = "";
            if (str6 == null) str6 = "";

        }

        public bool OpenFile(string FileName)
        {
            InputDataFileRows.Clear();
            if(File.Exists(FileName))
            {
                String connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                switch (FileName.Substring(FileName.LastIndexOf('.')))
                {
                    case ".xls":
                        connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                        break;
                    case ".xlsx":
                        connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
                        break;
                    default:
                        return false;
                }


                //String Command = "Show tables";
                try
                {
                    OleDbConnection con = new OleDbConnection(connection);

                    con.Open();
                    DataTable dtExcelSchema;
                    dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //con.Close();
                    DataSet ds = new DataSet();

                    string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    String Command = "Select * from [" + SheetName + "]";

                    OleDbCommand cmd = new OleDbCommand(Command, con);
                    OleDbDataAdapter db = new OleDbDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    db.Fill(dt);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if ((dt.Rows[i].ItemArray.GetValue(0).ToString().Length+ dt.Rows[i].ItemArray.GetValue(1).ToString().Length+ dt.Rows[i].ItemArray.GetValue(2).ToString().Length) == 0)
                        {
                            dt.Rows[i].Delete();
                            //i -= 1;
                        }

                    }
                    dt.AcceptChanges();
                    //dataGridViewFast.ItemsSource = dt.AsDataView();
                    cmd.Dispose();
                    con.Close();
                    con.Dispose();
                    //int StartRow = 0;
                    //int StartColumn = 0;
                    for(int i=0;i< dt.Rows.Count;i++)
                    {
                        for(int j=0;j<dt.Columns.Count;j++)
                        {
                            Regex regexTime = new Regex(@"^( *[Cc] *)?\d{1,2} *[\.,:;\- ]? *\d{1,2}(( *[\-–\/\\\| ] *)|( +)|( *[Дд][Оо] *))\d{1,2} *[\.,:;\- ]? *\d{1,2} *$");
                            if (regexTime.IsMatch(dt.Rows[i].ItemArray.GetValue(j).ToString()))
                            {
                                Regex regexName = new Regex(@"^ *([А-Я]|Ё)([а-я]|ё)+ +(([А-Я]|Ё) *\. *){0,2} *$");
                                Regex regexDate = new Regex(@"^ *\d{1,2} *[\.,:;\- \/]? *\d{1,2} *[\.,:;\- \/]? *((\d{2})|(\d{4})) *$");
                                if(j>0)
                                {
                                    if (regexDate.IsMatch(dt.Rows[i].ItemArray.GetValue(j - 1).ToString()))
                                        InputDataFileRows.Add(new DataTableRow(dt.Rows[i].ItemArray.GetValue(j-1).ToString(), dt.Rows[i].ItemArray.GetValue(j).ToString(), dt.Rows[i].ItemArray.GetValue(j+1).ToString(), dt.Rows[i].ItemArray.GetValue(j + 2).ToString(), dt.Rows[i].ItemArray.GetValue(j + 3).ToString(), dt.Rows[i].ItemArray.GetValue(j + 4).ToString(),true));
                                    else
                                    {
                                        if (regexName.IsMatch(dt.Rows[i].ItemArray.GetValue(j + 1).ToString()))
                                            InputDataFileRows.Add(new DataTableRow(dt.Rows[i].ItemArray.GetValue(j - 1).ToString(), dt.Rows[i].ItemArray.GetValue(j).ToString(), dt.Rows[i].ItemArray.GetValue(j + 1).ToString(), dt.Rows[i].ItemArray.GetValue(j + 2).ToString(), dt.Rows[i].ItemArray.GetValue(j + 3).ToString(), dt.Rows[i].ItemArray.GetValue(j + 4).ToString(), true));
                                    }    
                                }                                
                                else
                                {
                                    if (regexName.IsMatch(dt.Rows[i].ItemArray.GetValue(j + 1).ToString()))
                                        InputDataFileRows.Add(new DataTableRow(dt.Rows[i].ItemArray.GetValue(j - 1).ToString(), dt.Rows[i].ItemArray.GetValue(j).ToString(), dt.Rows[i].ItemArray.GetValue(j + 1).ToString(), dt.Rows[i].ItemArray.GetValue(j + 2).ToString(), dt.Rows[i].ItemArray.GetValue(j + 3).ToString(), dt.Rows[i].ItemArray.GetValue(j + 4).ToString(), true));
                                }
                            }
                        }
                    }
                    //InputDataFileRows.Add(new DataTableRow(dt.Rows[StartRow].ItemArray.GetValue(StartColumn).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+1).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+2).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+3).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+4).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+5).ToString()));
                }
                catch
                {
                    return false;
                }
                
            }
            if (InputDataFileRows.Count == 0)
                return false;
            //InputDataFileRows.Add(new DataTableRow("1", "2", "3", "4", "5", "6"));
            return true;
        }

    }
}

