using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public DataTableRow(string inputDate, string inputTime, string inputTeacher, string inputGroup, string inputCategory, string inputPlace)
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

}

