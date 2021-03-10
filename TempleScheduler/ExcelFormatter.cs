using Prism.Mvvm;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using MessageBox = System.Windows.MessageBox;
using System.Text.RegularExpressions;

namespace TempleScheduler
{
    public class ExcelFormatter
    {
        /**
         * !param path
         * Used for the constructor and is expected to be passed by the user when they select a path using the "Path" button
         *
         * !param weekdays
         * A short list to create new worksheets, the intention is to add an "Overview" tab at a later date.
         *
         * !param time
         * using EPPlus, I map this onto the excel worksheet using the LoadFromCollection function. This function is very powerful and if I knew about it earlier, I would have abused it like wildfire.
         * For future reference use this for anything whenever you want to map stuff.
         *
         * !param times
         * This dictionary is used to manually map a position for the hours selected by the user and represent the row of insertion. These are likely not to change unless the person maintaining wants to change the layout.
         *
         *
         */
        public string path;

        public string semester;

        private List<string> weekdays = new List<string> {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"};

        List<string> time = new List<string>
        {
            "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM",
            "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM",
            "5:00 PM", "5:30 PM",
        };

        public Dictionary<string, int> times = new()
        {
            {"8:00 AM", 3},
            {"8:30 AM", 4},
            {"9:00 AM", 5},
            {"9:30 AM", 6},
            {"10:00 AM", 7},
            {"10:30 AM", 8},
            {"11:00 AM", 9},
            {"11:30 AM", 10},
            {"12:00 PM", 11},
            {"12:30 PM", 12},
            {"1:00 PM", 13},
            {"1:30 PM", 14},
            {"2:00 PM", 15},
            {"2:30 PM", 16},
            {"3:00 PM", 17},
            {"3:30 PM", 18},
            {"4:00 PM", 19},
            {"4:30 PM", 20},
            {"5:00 PM", 21},
            {"5:30 PM", 22},
        };

        public ExcelFormatter(string path, string semester)
        {
            this.path = path;
            this.semester = semester;
        }

        /**
         *!function PersonsJSONDeserializer():
         * Using the path selected by the user, it'll look through the files available and put together a list of Deserialized jsons into Schedule objects.
         *
         */
        public async Task<List<Schedule>> PersonsJSONDeserializer()
        {
            string[] fileEntries = Directory.GetFiles(this.path);
            Schedule person;
            List<Schedule> schedules = new List<Schedule>();

            foreach (string staff in fileEntries)
            {
                if (staff.Contains("_Schedules.json"))
                {
                    string json = System.IO.File.ReadAllText(staff);
                    person = JsonConvert.DeserializeObject<Schedule>(json);
                    schedules.Add(person);
                }
            }

            return schedules;
        }

        /**
         * !function ExcelCreator():
         *<summary>wow this comment markup is awful!</summary>
         * <param name="yourMom"></param>
         */
        public async Task ExcelCreator()
        {
            DateTime today = DateTime.Today;
            string outputFile = $"{this.path}\\{today.ToString(format: "yyyyMMddfff")}_Schedules.xlsx";
            Console.WriteLine(value: outputFile);
            var file = new FileInfo(fileName: outputFile);
            DeleteFileIfExists(file: file);

            var staff = await PersonsJSONDeserializer();

            /**
             * The using keyword allows us to use a file and not worrying about closing it manually later. The old school way of doing this would have been
             * package.Dispose() or in VBA it'd be Workbooks(file).Close
             */
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(newFile: file))
            {
                var ws = package.Workbook.Worksheets.Add(Name: "Overview");
                ws.Cells[Address: "A1:F1"].Merge = true;
                if (semester != "")
                {
                    ws.Cells[Address: "A1"].Value = "Student Workers " + semester + " Schedule";
                }
                else
                {
                    ws.Cells[Address: "A1"].Value = "Student Workers Schedule";
                }

                ws.Cells[Address: "A1"].Style.Font.Bold = true;
                ws.Cells[Address: "A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[Address: "A1"].Style.Font.Size = 26;
                int globalOverviewSpotMax = 0;
                int max = 3;
                int currentOverviewSpot = 3;
                for (int k = 0; k < staff.Count; k++)
                {
                    ws.Cells[currentOverviewSpot, 1].Value = staff[k].name;
                    if (staff[k].phone != "")
                    {
                        ws.Cells[currentOverviewSpot + 1, 1].Value =
                            "Phone: " + Convert.ToInt64(staff[k].phone).ToString("###-###-####");
                    }
                    else
                    {
                        ws.Cells[currentOverviewSpot + 1, 1].Value =
                            "Phone: N/A";
                    }

                    if (staff[k].office != "")
                    {
                        ws.Cells[currentOverviewSpot + 2, 1].Value =
                            "Office: " + Convert.ToInt64(staff[k].office).ToString("###-###-####");
                    }
                    else
                    {
                        ws.Cells[currentOverviewSpot + 2, 1].Value =
                            "Office: N/A";
                    }
                    for (int j = 0; j < 5; j++)
                    {
                        ws.Cells[2, j + 2].Value = weekdays[j];




                        int nRange = 0;
                        int fRange = 0;

                        int flexAndNormalRangeCount = staff[k].normalRanges[j].Count() + staff[k].flexRanges[j].Count();
                        for (int i = 0; i < flexAndNormalRangeCount; i++)
                        {
                            if (max < flexAndNormalRangeCount)
                            {
                                max = flexAndNormalRangeCount;
                            }

                            if (staff[k].normalRanges[j].Count() == 0 || staff[k].normalRanges[j].Count() == nRange)
                            {
                                ws.Cells[currentOverviewSpot + i, j + 2].Value =
                                    time[staff[k].flexRanges[j].ElementAt(fRange).Item1] + "-" +
                                    time[staff[k].flexRanges[j].ElementAt(fRange).Item2];
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.BackgroundColor
                                    .SetColor(color: System.Drawing.Color.PowderBlue);
                                fRange++;
                            }
                            else if (staff[k].flexRanges[j].Count() == 0 || staff[k].flexRanges[j].Count() == fRange)
                            {
                                ws.Cells[currentOverviewSpot + i, j + 2].Value =
                                    time[staff[k].normalRanges[j].ElementAt(nRange).Item1] + "-" +
                                    time[staff[k].normalRanges[j].ElementAt(nRange).Item2];
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.BackgroundColor
                                    .SetColor(color: System.Drawing.Color.LightGreen);
                                nRange++;
                            }
                            else if (staff[k].normalRanges[j].ElementAt(nRange).Item1 <
                                     staff[k].flexRanges[j].ElementAt(fRange).Item1)
                            {
                                ws.Cells[currentOverviewSpot + i, j + 2].Value =
                                    time[staff[k].normalRanges[j].ElementAt(nRange).Item1] + "-" +
                                    time[staff[k].normalRanges[j].ElementAt(nRange).Item2];
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.BackgroundColor
                                    .SetColor(color: System.Drawing.Color.LightGreen);
                                nRange++;
                            }
                            else if (staff[k].flexRanges[j].ElementAt(fRange).Item1 <
                                     staff[k].normalRanges[j].ElementAt(nRange).Item1)
                            {
                                ws.Cells[currentOverviewSpot + i, j + 2].Value =
                                    time[staff[k].flexRanges[j].ElementAt(fRange).Item1] + "-" +
                                    time[staff[k].flexRanges[j].ElementAt(fRange).Item2];
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[currentOverviewSpot + i, j + 2].Style.Fill.BackgroundColor
                                    .SetColor(color: System.Drawing.Color.PowderBlue);
                                fRange++;
                            }
                        }



                    }
                    currentOverviewSpot += max + 1;
                    if (globalOverviewSpotMax < currentOverviewSpot)
                    {
                        globalOverviewSpotMax = currentOverviewSpot;
                    }
                }

                ws.Cells.AutoFitColumns();
                StylingKey(ws: ws, currentCol: 6, shift: 1, globalOverviewSpotMax);


                for (int j = 0; j < 5; j++)
                {
                    ws = package.Workbook.Worksheets.Add(Name: weekdays[index: j]);
                    ws.Cells[Row: 3, Col: 1].LoadFromCollection(Collection: this.time);
                    ws.Cells[Address: "A1:F1"].Merge = true;
                    ws.Cells[Address: "A1"].Value = weekdays[index: j];
                    ws.Cells[Address: "A1"].Style.Font.Bold = true;
                    ws.Cells[Address: "A1"].Style.Font.Size = 26;
                    int currentCol = 2;
                    for (int i = 0; i < staff.Count(); i++)
                    {
                        ws.Cells[Row: 2, Col: currentCol].Value = staff[index: i].name;
                        ws.Cells[Row: 2, Col: currentCol].Style.Font.Bold = true;
                        foreach (string hour in staff[index: i].normalTimes[index: j])
                        {
                            ws.Cells[Row: times[key: hour], Col: currentCol].Style.Fill.PatternType =
                                ExcelFillStyle.Solid;
                            ws.Cells[Row: times[key: hour], Col: currentCol].Style.Fill.BackgroundColor
                                .SetColor(color: System.Drawing.Color.LightGreen);
                        }

                        foreach (string hour in staff[index: i].flexTimes[index: j])
                        {
                            ws.Cells[Row: times[key: hour], Col: currentCol].Style.Fill.PatternType =
                                ExcelFillStyle.Solid;
                            ws.Cells[Row: times[key: hour], Col: currentCol].Style.Fill.BackgroundColor
                                .SetColor(color: System.Drawing.Color.PowderBlue);
                        }

                        currentCol = currentCol + 1;
                        if (i == staff.Count - 1)
                        {
                            ws.Cells[Row: 3, Col: currentCol].LoadFromCollection(Collection: this.time);
                            ws.Cells[FromRow: 2, FromCol: 1, ToRow: 22, ToCol: currentCol].Style.Border.Top.Style =
                                ExcelBorderStyle.Thin;
                            ws.Cells[FromRow: 2, FromCol: 1, ToRow: 22, ToCol: currentCol].Style.Border.Left.Style =
                                ExcelBorderStyle.Thin;
                            ws.Cells[FromRow: 2, FromCol: 1, ToRow: 22, ToCol: currentCol].Style.Border.Right.Style =
                                ExcelBorderStyle.Thin;
                            ws.Cells[FromRow: 2, FromCol: 1, ToRow: 22, ToCol: currentCol].Style.Border.Bottom.Style =
                                ExcelBorderStyle.Thin;

                            if (currentCol % 2 == 0)
                            {
                                StylingKey(ws: ws, currentCol: currentCol, shift: 1, 24);
                            }
                            else
                            {
                                StylingKey(ws: ws, currentCol: currentCol, shift: 2, 24);
                            }
                        }
                    }

                    ws.Cells.AutoFitColumns();
                    ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                await package.SaveAsAsync(file: file);
            }
        }

        private static void StylingKey(ExcelWorksheet ws, int currentCol, int shift, int row)
        {
            ws.Cells[row, currentCol / 2].Value = "Regular";
            ws.Cells[row, currentCol / 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, currentCol / 2].Style.Fill.BackgroundColor
                .SetColor(System.Drawing.Color.LightGreen);
            ws.Cells[row, currentCol / 2].Style.Font.Bold = true;

            ws.Cells[row, currentCol / 2 + shift].Value = "Flex";
            ws.Cells[row, currentCol / 2 + shift].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, currentCol / 2 + shift].Style.Fill.BackgroundColor
                .SetColor(System.Drawing.Color.PowderBlue);
            ws.Cells[row, currentCol / 2 + shift].Style.Font.Bold = true;
        }


        private void DeleteFileIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }
    }
}