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

        public DataTableRow(string inputDate, string inputTime, string inputTeacher, string inputGroup, string inputCategory, string inputPlace, bool testcorrect = false)
        {
            // DataTableRow("06.11.2019", "10:00-16:40", "Пронина Л.Н.", "","-----","-----"));
            if (inputDate != null)
                Date = inputDate;
            else
                Date = "";
            Time = "";
            if (inputTeacher != null)
                Teacher = inputTeacher;
            else
                Teacher = "";
            if (inputGroup != null)
                Group = inputGroup;
            else
                Group = "";
            if (inputCategory != null)
                Category = inputCategory;
            else
                Category = "";
            if (inputPlace != null)
                Place = inputPlace;
            else
                Place = "";
            if (testcorrect)
            {
                string bufInputDate = inputDate;
                if (bufInputDate == null)
                    bufInputDate = "";
                if (bufInputDate.Length > 0)
                {
                    int startindex = 0;
                    while (bufInputDate[startindex] < '0' || bufInputDate[startindex] > '9')
                    {
                        startindex += 1;
                    }
                    bufInputDate = bufInputDate.Substring(startindex);
                    int endindex = bufInputDate.Length - 1;
                    while (bufInputDate[endindex] < '0' || bufInputDate[endindex] > '9')
                    {
                        endindex -= 1;
                    }
                    bufInputDate = bufInputDate.Substring(0, endindex + 1);
                    Regex regexDateDigit = new Regex(@"\d");
                    Regex regexDateSeparate = new Regex(@"[\.,:;\-\/ ]");
                    List<bool> bufFlags = new List<bool>();
                    for (int i = 0; i < bufInputDate.Length; i++)
                    {
                        bufFlags.Add(true);
                    }
                    MatchCollection matchesDate = regexDateDigit.Matches(bufInputDate);
                    MatchCollection matchesDateSeparate = regexDateSeparate.Matches(bufInputDate);
                    for (int i = 1; i < matchesDateSeparate.Count; i++)
                    {
                        if (matchesDateSeparate[i].Index == (matchesDateSeparate[i - 1].Index - 1))
                        {
                            bufFlags[matchesDateSeparate[i].Index] = false;
                        }
                    }
                    List<char> bufCharDate = new List<char>();
                    for (int i = 0; i < bufInputDate.Length; i++)
                    {
                        if (bufFlags[i])
                            bufCharDate.Add(bufInputDate[i]);
                    }
                    bufInputDate = new string(bufCharDate.ToArray());
                    matchesDate = regexDateDigit.Matches(bufInputDate);
                    matchesDateSeparate = regexDateSeparate.Matches(bufInputDate);
                    bufInputDate = regexDateSeparate.Replace(bufInputDate, ".");
                    try
                    {
                        Date = Convert.ToDateTime(bufInputDate).GetDateTimeFormats('d')[0];
                    }
                    catch
                    {
                        Date = "";
                    }
                }
                string bufInputTime = inputTime;
                if (bufInputTime == null)
                    bufInputTime = "";
                if (bufInputTime.Length > 0)
                {
                    int startindex = 0;
                    while (bufInputTime[startindex] < '0' || bufInputTime[startindex] > '9')
                    {
                        startindex += 1;
                    }
                    bufInputTime = bufInputTime.Substring(startindex);
                    int endindex = bufInputTime.Length - 1;
                    while (bufInputTime[endindex] < '0' || bufInputTime[endindex] > '9')
                    {
                        endindex -= 1;
                    }
                    bufInputTime = bufInputTime.Substring(0, endindex + 1);
                    Regex regexTimeDigit = new Regex(@"\d");
                    Regex regexTimeSeparate = new Regex(@"[\.,:;\-\/ ]");
                    Regex regexTimeSeparateDubl = new Regex(@"\D{2,}");
                    //bufInputTime = bufInputTime.Replace(" ", "");
                    MatchCollection matchesTimeSeparateDubl = regexTimeSeparateDubl.Matches(bufInputTime);
                    int sdvig = 0;
                    for (int i=0;i<matchesTimeSeparateDubl.Count;i++)
                    {
                        bufInputTime=bufInputTime.Remove(matchesTimeSeparateDubl[i].Index-sdvig+1, matchesTimeSeparateDubl[i].Length-1);
                        sdvig += matchesTimeSeparateDubl[i].Length - 1;
                    }
                    List<bool> bufFlags = new List<bool>();
                    for (int i = 0; i < bufInputTime.Length; i++)
                    {
                        bufFlags.Add(true);
                    }
                    MatchCollection matchesTime = regexTimeDigit.Matches(bufInputTime);
                    MatchCollection matchesTimeSeparate = regexTimeSeparate.Matches(bufInputTime);
                    for (int i = 1; i < matchesTimeSeparate.Count; i++)
                    {
                        if (matchesTimeSeparate[i].Index == (matchesTimeSeparate[i - 1].Index - 1))
                        {
                            bufFlags[matchesTimeSeparate[i].Index] = false;
                        }
                    }
                    List<char> bufCharTime = new List<char>();
                    for (int i = 0; i < bufInputTime.Length; i++)
                    {
                        if (bufFlags[i])
                            bufCharTime.Add(bufInputTime[i]);
                    }
                    bufInputTime = new string(bufCharTime.ToArray());
                    matchesTime = regexTimeDigit.Matches(bufInputTime);
                    matchesTimeSeparate = regexTimeSeparate.Matches(bufInputTime);
                    bufInputTime = regexTimeSeparate.Replace(bufInputTime, ":");
                    if (matchesTimeSeparate.Count == 3)
                    {
                        string bufInputTimeStart = bufInputTime.Substring(0, matchesTimeSeparate[1].Index);
                        string bufInputTimeEnd = bufInputTime.Substring(matchesTimeSeparate[1].Index + 1);
                        try
                        {
                            if (Convert.ToDateTime(bufInputTimeStart) > Convert.ToDateTime(bufInputTimeEnd))
                            {
                                string buftemptime = bufInputTimeStart;
                                bufInputTimeStart = bufInputTimeEnd;
                                bufInputTimeEnd = buftemptime;
                            }
                            Time = (Convert.ToDateTime(bufInputTimeStart).GetDateTimeFormats('t')[0] + "-" + Convert.ToDateTime(bufInputTimeEnd).GetDateTimeFormats('t')[0]).Replace(':', '.');
                        }
                        catch
                        {
                            Time = "";
                        }
                    }
                    else
                    {
                        Time = "";
                    }
                    //MessageBox.Show(Time);
                }

                string bufInputTeacher = inputTeacher;
                if (bufInputTeacher == null)
                    bufInputTeacher = "";
                if (bufInputTeacher.Length > 0)
                {
                    Regex regexTeacherMoodle = new Regex(@"moodle", RegexOptions.IgnoreCase);
                    if (regexTeacherMoodle.IsMatch(bufInputTeacher))
                    {
                        bufInputTeacher = "Moodle";
                        Teacher = "Moodle";
                       // MessageBox.Show("To Moodle");
                    }
                    else
                    {
                        //MessageBox.Show("From " + bufInputTeacher);
                        //Regex regexTeacherSeparate = new Regex(@"[\.,:;\-\/ ]");
                        Regex regexTeacherSeparate = new Regex(@"([А-Я]|Ё)|([а-я]|ё)");
                        Regex regexTeacherChar = new Regex(@"([А-Я]|Ё)|([а-я]|ё)");
                        bufInputTeacher += ".";
                        MatchCollection matchTeacherChar = regexTeacherChar.Matches(bufInputTeacher);
                        bufInputTeacher = bufInputTeacher.Substring(matchTeacherChar[0].Index, matchTeacherChar[matchTeacherChar.Count - 1].Index + 2);
                        MatchCollection matchesTeacherSeparate = regexTeacherSeparate.Matches(bufInputTeacher);
                        //MessageBox.Show("From (" + startindex + " - " + (endindex + 1) + ") " + bufInputTeacher);
                        //bufInputTeacher = bufInputTeacher.Substring(0, endindex + 1);
                        //MessageBox.Show("From ("+startindex+" - "+endindex+") "+bufInputTeacher);
                        /*List<bool> bufFlags = new List<bool>();
                        for (int i = 0; i < bufInputTeacher.Length; i++)
                        {
                            bufFlags.Add(true);
                        }
                        
                        for (int i = 1; i < matchesTeacherSeparate.Count; i++)
                        {
                            if (matchesTeacherSeparate[i].Index == (matchesTeacherSeparate[i - 1].Index - 1))
                            {
                                bufFlags[matchesTeacherSeparate[i].Index] = false;
                            }
                        }
                        List<char> bufCharTeacher = new List<char>();
                        for (int i = 0; i < bufInputTeacher.Length; i++)
                        {
                            if (bufFlags[i])
                                bufCharTeacher.Add(bufInputTeacher[i]);
                        }
                        */
                        string bufCharTeacher = "";
                        bufCharTeacher += matchTeacherChar[0].Value;
                        for (int i = 1; i < matchTeacherChar.Count; i++)
                        {
                            if (matchTeacherChar[i].Index!=(matchTeacherChar[i-1].Index+1))
                            {
                                bufCharTeacher += " ";
                                bufCharTeacher += matchTeacherChar[i].Value.ToUpper();
                            }
                            else
                            {
                                if((((matchTeacherChar[i].Value[0]>='А')&& (matchTeacherChar[i].Value[0] <= 'Я'))|| (matchTeacherChar[i].Value[0] == 'Ё')) || ((matchTeacherChar[i].Value[0] >= 'A') && (matchTeacherChar[i].Value[0] <= 'Z')))
                                {
                                    if ((((matchTeacherChar[i-1].Value[0] >= 'А') && (matchTeacherChar[i-1].Value[0] <= 'Я')) || (matchTeacherChar[i-1].Value[0] == 'Ё')) || ((matchTeacherChar[i-1].Value[0] >= 'A') && (matchTeacherChar[i-1].Value[0] <= 'Z')))
                                    {
                                        bufCharTeacher += matchTeacherChar[i].Value.ToLower();
                                    }
                                    else
                                    {
                                        bufCharTeacher += " ";
                                        bufCharTeacher += matchTeacherChar[i].Value;
                                    }
                                        
                                }
                                else
                                {
                                    bufCharTeacher += matchTeacherChar[i].Value;
                                }
                            }
                                
                        }
                        bufInputTeacher = new string(bufCharTeacher.ToArray());
                        Regex regexTeacherNames = new Regex(@"(([А-Я]|Ё)(([а-я]|ё)*))|(([A-Z])(([a-z])*))");
                        MatchCollection matchTeacherNames = regexTeacherNames.Matches(bufInputTeacher);
                        if(matchTeacherNames.Count>2)
                        {
                            int fathernameIndex = -1;
                            for(int i=1;i<matchTeacherNames.Count;i++)
                            {
                                if (matchTeacherNames[i].Length>=3)
                                    if((matchTeacherNames[i].Value.Substring(matchTeacherNames[i].Length-3,3)=="вна")|| (matchTeacherNames[i].Value.Substring(matchTeacherNames[i].Length - 3, 3) == "вич"))
                                    {
                                        fathernameIndex = i;
                                        break;
                                    }
                            }
                            if(fathernameIndex>0)
                            {
                                if(fathernameIndex>1)
                                    Teacher = matchTeacherNames[fathernameIndex - 2].Value + " " + matchTeacherNames[fathernameIndex - 1].Value[0] + "." + matchTeacherNames[fathernameIndex].Value[0] + ".";
                                else
                                    Teacher = matchTeacherNames[fathernameIndex + 1].Value + " " + matchTeacherNames[fathernameIndex - 1].Value[0] + "." + matchTeacherNames[fathernameIndex].Value[0] + ".";
                            }
                            else
                            {
                                int longnameindex=-1;
                                for (int i = 0; i < matchTeacherNames.Count; i++)
                                {
                                    if(matchTeacherNames[i].Length > 2)
                                    {
                                        longnameindex = i;
                                        break;
                                    }
                                }
                                if(longnameindex == -1)
                                {
                                    Teacher = matchTeacherNames[0].Value+" "+ matchTeacherNames[1].Value[0]+"."+ matchTeacherNames[2].Value[0]+".";
                                }
                                else
                                {
                                    Teacher = matchTeacherNames[longnameindex].Value + " ";
                                    if ((longnameindex + 2) < matchTeacherNames.Count)
                                        Teacher += matchTeacherNames[longnameindex + 1].Value[0] + "." + matchTeacherNames[longnameindex + 2].Value[0] + ".";
                                    else
                                    {
                                        if ((longnameindex - 2) >= 0)
                                            Teacher += matchTeacherNames[longnameindex - 2].Value[0] + "." + matchTeacherNames[longnameindex - 1].Value[0] + ".";
                                        else
                                        {
                                            Teacher += matchTeacherNames[longnameindex - 1].Value[0] + "." + matchTeacherNames[longnameindex + 1].Value[0] + ".";
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            if(matchTeacherNames.Count > 0)
                            Teacher = matchTeacherNames[0].Value;
                            if(matchTeacherNames.Count > 1)
                                Teacher += " "+matchTeacherNames[0].Value[0]+".";
                        }
                        /*
                        matchesTeacherSeparate = regexTeacherSeparate.Matches(bufInputTeacher);
                        bufInputTeacher = regexTeacherSeparate.Replace(bufInputTeacher, ".");
                        if (matchesTeacherSeparate.Count > 0)
                            bufInputTeacher = bufInputTeacher.Remove(matchesTeacherSeparate[0].Index, 1).Insert(matchesTeacherSeparate[0].Index, " ");
                        Regex regexName = new Regex(@"^([А-Я]|Ё)([а-я]|ё)+ +(([А-Я]|Ё)*\. *){2}$");
                        if (regexName.IsMatch(bufInputTeacher))
                        {
                            Teacher = bufInputTeacher;
                        }
                        else
                        {
                            Regex regexInvertName = new Regex(@"^(([А-Я]|Ё)[\.,:;\-\/ ]*){2}[\.,:;\-\/ ]+([А-Я]|Ё)([а-я]|ё)+[\.,:;\-\/ ]*$");
                            if (regexInvertName.IsMatch(bufInputTeacher))
                            {
                                Regex regexBigChar = new Regex(@"([А-Я]|Ё)");
                                MatchCollection matchBigChar = regexBigChar.Matches(bufInputTeacher);
                                if (matchBigChar.Count == 3)
                                    Teacher = bufInputTeacher.Substring(matchBigChar[2].Index, matchesTeacherSeparate[matchesTeacherSeparate.Count - 1].Index - matchBigChar[2].Index) + " " + matchBigChar[0].Value + "." + matchBigChar[1].Value + ".";
                                else
                                    Teacher = "";
                            }
                            else
                                Teacher = "";
                        }
                        */
                    }
                }
            }
            else
            {

                string bufTime;
                if (inputTime != null)
                    bufTime = inputTime;
                else
                    bufTime = "";
                if (bufTime.Length > 0)
                {
                    if (bufTime[0] == '0')
                    {
                        Time = bufTime.Substring(1).Replace(':', '.');
                    }
                    else
                    {
                        Time = bufTime.Replace(':', '.');
                    }
                    if (Time[Time.IndexOf("-") + 1] == '0')
                    {
                        Time = Time.Substring(0, Time.IndexOf("-") + 1) + Time.Substring(Time.IndexOf("-") + 2);
                    }
                }
                else
                {
                    Time = "";
                }
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

        public bool Intersection(DataTableRow B = null)
        {
            return Intersection(this, B);
        }
        public static bool Intersection(DataTableRow A = null, DataTableRow B = null)
        {
            if (A == null || B == null)
                return false;
            if (A.Teacher != B.Teacher)
                return false;
            else
            {
                if (A.Teacher == null || B.Teacher == null)
                    return false;
                else
                {
                    if (A.Teacher == "" || B.Teacher == "" || A.Teacher == "Moodle" || B.Teacher == "Moodle")
                        return false;
                    else
                    {
                        if (A.Date != B.Date)
                            return false;
                        else
                        {
                            if (A.Date == null || B.Date == null)
                                return false;
                            else
                            {
                                if (A.Date == "" || B.Date == "")
                                    return false;
                                else
                                {
                                    if (A.Time != B.Time)
                                        return false;
                                    else
                                    {
                                        if (A.Time == null || B.Time == null)
                                            return false;
                                        else
                                        {
                                            if (A.Time == "" || B.Time == "")
                                                return false;
                                            else
                                            {
                                                string AStartTime, BStartTime, AEndTime, BEndTime;
                                                AStartTime = A.Time.Substring(0, A.Time.IndexOf('-')).Replace('.', ':');
                                                BStartTime = B.Time.Substring(0, B.Time.IndexOf('-')).Replace('.', ':');
                                                AEndTime = A.Time.Substring(A.Time.IndexOf('-') + 1).Replace('.', ':');
                                                BEndTime = B.Time.Substring(B.Time.IndexOf('-') + 1).Replace('.', ':');
                                                DateTime AStart, BStart, AEnd, BEnd;
                                                AStart = Convert.ToDateTime(AStartTime);
                                                BStart = Convert.ToDateTime(BStartTime);
                                                AEnd = Convert.ToDateTime(AEndTime);
                                                BEnd = Convert.ToDateTime(BEndTime);
                                                if (BEnd <= AStart || BStart >= AEnd)
                                                {
                                                    return false;
                                                }
                                                else
                                                    return true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return false;
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

        public void Add(String str1 = "", String str2 = "", String str3 = "", String str4 = "", String str5 = "", String str6 = "")
        {
            if (str1 == null) str1 = "";
            if (str2 == null) str2 = "";
            if (str3 == null) str3 = "";
            if (str4 == null) str4 = "";
            if (str5 == null) str5 = "";
            if (str6 == null) str6 = "";

        }

        public bool OpenFile(string FileName)
        {
            InputDataFileRows.Clear();
            if (File.Exists(FileName))
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
                        if ((dt.Rows[i].ItemArray.GetValue(0).ToString().Length + dt.Rows[i].ItemArray.GetValue(1).ToString().Length + dt.Rows[i].ItemArray.GetValue(2).ToString().Length) == 0)
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
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            Regex regexTime = new Regex(@"^( *[Cc] *)?\d{1,2} *[\.,:;\- ]? *\d{1,2}(( *[\-–\/\\\| ] *)|( +)|( *[Дд][Оо] *))\d{1,2} *[\.,:;\- ]? *\d{1,2} *$");
                            if (regexTime.IsMatch(dt.Rows[i].ItemArray.GetValue(j).ToString()))
                            {
                                Regex regexName = new Regex(@"^ *([А-Я]|Ё)([а-я]|ё)+ +(([А-Я]|Ё) *\. *){0,2} *$");
                                Regex regexDate = new Regex(@"^ *\d{1,2} *[\.,:;\- \/]? *\d{1,2} *[\.,:;\- \/]? *((\d{2})|(\d{4})) *$");
                                if (j > 0)
                                {
                                    if (regexDate.IsMatch(dt.Rows[i].ItemArray.GetValue(j - 1).ToString()))
                                        InputDataFileRows.Add(new DataTableRow(dt.Rows[i].ItemArray.GetValue(j - 1).ToString(), dt.Rows[i].ItemArray.GetValue(j).ToString(), dt.Rows[i].ItemArray.GetValue(j + 1).ToString(), dt.Rows[i].ItemArray.GetValue(j + 2).ToString(), dt.Rows[i].ItemArray.GetValue(j + 3).ToString(), dt.Rows[i].ItemArray.GetValue(j + 4).ToString(), true));
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
                    for (int i = 0; i < InputDataFileRows.Count; i++)
                    {
                        if (InputDataFileRows[i].Time == null)
                        {
                            InputDataFileRows.RemoveAt(i);
                            i -= 1;
                        }
                        else
                        {
                            if (InputDataFileRows[i].Time.Length == 0)
                            {
                                InputDataFileRows.RemoveAt(i);
                                i -= 1;
                            }
                            else
                            {

                            }
                        }
                    }
                    //InputDataFileRows.Add(new DataTableRow(dt.Rows[StartRow].ItemArray.GetValue(StartColumn).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+1).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+2).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+3).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+4).ToString(), dt.Rows[StartRow].ItemArray.GetValue(StartColumn+5).ToString()));
                }
                catch
                {
                    //MessageBox.Show("!");
                    return false;
                }

            }
            if (InputDataFileRows.Count == 0)
            {
                //MessageBox.Show("?");
                return false;
            }

            //InputDataFileRows.Add(new DataTableRow("1", "2", "3", "4", "5", "6"));
            return true;
        }

    }
}

