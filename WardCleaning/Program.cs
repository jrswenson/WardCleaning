using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WardCleaning
{
    class Program
    {
        static void Main(string[] args)
        {
            var tmpString = ConfigurationManager.AppSettings["StartDate"];
            DateTime start;
            if (DateTime.TryParse(tmpString, out start) == false)
                throw new Exception($"StartDate = {tmpString}");

            tmpString = ConfigurationManager.AppSettings["EndDate"];
            DateTime end;
            if (DateTime.TryParse(tmpString, out end) == false)
                throw new Exception($"EndDate = {tmpString}");

            var weekAssignments = new Dictionary<DateTime, HashSet<string>>();
            var tmpDate = start;
            while (tmpDate <= end)
            {
                if (tmpDate.DayOfWeek == DayOfWeek.Saturday)
                {
                    weekAssignments.Add(tmpDate, new HashSet<string>());
                    tmpDate = tmpDate.AddDays(7);
                }
                else
                {
                    var daysToAdd = ((int)DayOfWeek.Saturday - (int)start.DayOfWeek + 7) % 7;
                    tmpDate = tmpDate.AddDays(daysToAdd);
                }
            }

            var numTimesAssigned = 1;
            tmpString = ConfigurationManager.AppSettings["TimesAssigned"];
            if (int.TryParse(tmpString, out numTimesAssigned) == false)
                throw new Exception($"TimesAssigned = {tmpString}");

            var members = File.ReadAllLines("MemberList.txt");
            var randomMembers = members.Randomize();

            var assignmentsPerWeek = Math.Floor((members.Count() * numTimesAssigned) / weekAssignments.Count() * 1.0);
            var memberAssignments = new Dictionary<string, ICollection<DateTime>>();
            var random = new Random();
            foreach (var member in randomMembers)
            {
                var dateList = new List<DateTime>();
                memberAssignments.Add(member, dateList);
                int lastRan = 0;
                while (dateList.Count() < 2)
                {
                    var index = random.Next(1, 17);
                    DateTime dateVal;
                    if (lastRan != 0)
                    {
                        while (index == lastRan || index == lastRan + 1 || index == lastRan - 1)
                        {
                            index = random.Next(1, 18);
                        }
                    }
                    lastRan = index;

                    dateVal = weekAssignments.Keys.ElementAt(index - 1);

                    HashSet<string> assigned;
                    if (weekAssignments.TryGetValue(dateVal, out assigned) == false)
                        throw new Exception("Something went wrong with the dates");

                    if (assigned.Count < assignmentsPerWeek)
                    {
                        assigned.Add(member);
                        dateList.Add(dateVal);
                    }
                }
            }

            var fileName = $@"Cleaning Assignment {start.Month}-{end.Month}-{end.Year}.xlsx";
            File.Delete(fileName);
            var newFile = new FileInfo(fileName);
            using (ExcelPackage pckg = new ExcelPackage(newFile))
            {
                var ws1 = pckg.Workbook.Worksheets.Add("Weekly Assignments");
                var rowIndex = 1;
                var colName = "A";
                foreach (var item in weekAssignments)
                {
                    var x = new List<string>() { $"{item.Key.Month}/{item.Key.Day}/{item.Key.Year}" };
                    foreach (var val in item.Value)
                    {
                        x.Add(val);
                    }

                    var rowCount = rowIndex + x.Count;
                    ws1.Cells[$"{colName}{rowIndex}:{colName}{rowCount}"].LoadFromCollection(x);

                    switch (colName)
                    {
                        case "A":
                            colName = "C";
                            break;
                        case "C":
                            colName = "E";
                            break;
                        default:
                            colName = "A";
                            rowIndex = rowCount + 1;
                            break;
                    }
                }

                var ws2 = pckg.Workbook.Worksheets.Add("Family Assignments");

                var alpha = memberAssignments.OrderBy(k => k.Key);
                rowIndex = 1;
                colName = "B";
                foreach (var item in alpha)
                {
                    ws2.Cells[$"A{rowIndex}:A{rowIndex}"].Value = item.Key;
                    foreach (var val in item.Value)
                    {
                        ws2.Cells[$"{colName}{rowIndex}:{colName}{rowIndex}"].Value = $"{val.Month}/{val.Day}/{val.Year}";
                        switch (colName)
                        {
                            case "B":
                                colName = "C";
                                break;
                            default:
                                colName = "B";
                                break;
                        }
                    }

                    rowIndex++;
                }

                pckg.Save();
            }
        }
    }
}
