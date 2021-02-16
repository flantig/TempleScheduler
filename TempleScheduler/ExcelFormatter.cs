using Prism.Mvvm;
using System;
using System.Collections;
using System.Collections.Generic;
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

namespace TempleScheduler
{
    public class ExcelFormatter
    {
        public string path;
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

        public ExcelFormatter(string path)
        {
            this.path = path;
        }

        public async Task<List<Schedule>> PersonsJSONDeserializer()
        {
            string[] fileEntries = Directory.GetFiles(this.path);
            Schedule person;
            List<Schedule> schedules = new List<Schedule>();

            foreach (string staff in fileEntries)
            {
                if (staff.Contains(".json"))
                {
                    Console.WriteLine("We found one!");
                    string json = System.IO.File.ReadAllText(staff);
                    person = JsonConvert.DeserializeObject<Schedule>(json);
                    schedules.Add(person);
                }
            }

            return schedules;
        }

        public async Task ExcelCreator()
        {
            DateTime today = DateTime.Today;
            string outputFile = $"{this.path}\\{today.ToString("yyyyMMddfff")}_Schedules.xlsx";
            Console.WriteLine(outputFile);
            var file = new FileInfo(outputFile);
            DeleteFileIfExists(file);

            var staff = await PersonsJSONDeserializer();

            /*
             * The using keyword allows us to use a file and not worrying about closing it manually later. The old school way of doing this would have been
             * package.Dispose() or in VBA it'd be Workbooks(file).Close
             */
            using (var package = new ExcelPackage(file))
            {
                for (int j = 0; j < 5; j++)
                {
                    var ws = package.Workbook.Worksheets.Add(weekdays[j]);
                    ws.Cells[3, 1].LoadFromCollection(this.time);
                    ws.Cells["A1:F1"].Merge = true;
                    ws.Cells["A1"].Value = weekdays[j];
                    ws.Cells["A1"].Style.Font.Bold = true;
                    ws.Cells["A1"].Style.Font.Size = 26;
                    int currentCol = 2;
                    for (int i = 0; i < staff.Count(); i++)
                    {
                        ws.Cells[2, currentCol].Value = staff[i].name;
                        ws.Cells[2, currentCol].Style.Font.Bold = true;
                        foreach (string hour in staff[i].normalTimes[j])
                        {
                            ws.Cells[times[hour], currentCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[times[hour], currentCol].Style.Fill.BackgroundColor
                                .SetColor(System.Drawing.Color.LightGreen);
                        }

                        foreach (string hour in staff[i].flexTimes[j])
                        {
                            ws.Cells[times[hour], currentCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[times[hour], currentCol].Style.Fill.BackgroundColor
                                .SetColor(System.Drawing.Color.PowderBlue);
                        }

                        currentCol = currentCol + 1;
                        if (i == staff.Count - 1)
                        {
                            ws.Cells[3, currentCol].LoadFromCollection(this.time);
                            ws.Cells[2, 1, 22, currentCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            ws.Cells[2, 1, 22, currentCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells[2, 1, 22, currentCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            ws.Cells[2, 1, 22, currentCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            if (currentCol % 2 == 0)
                            {
                                ws.Cells[24,currentCol/2].Value = "Regular";
                                ws.Cells[24,currentCol/2].Style.Font.Bold = true;
                                ws.Cells[24,currentCol/2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[24,currentCol/2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                                ws.Cells[24,currentCol/2 + 1].Style.Font.Bold = true;
                                ws.Cells[24,currentCol/2 + 1].Value = "Flex";
                                ws.Cells[24,currentCol/2 + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[24,currentCol/2 + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PowderBlue);
                            }
                            else
                            {
                                ws.Cells[24, currentCol/2 ].Value = "Regular";
                                ws.Cells[24,currentCol/2 ].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[24,currentCol/2 ].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                ws.Cells[24,currentCol/2 ].Style.Font.Bold = true;

                                ws.Cells[24,currentCol/2 + 2].Value = "Flex";
                                ws.Cells[24,currentCol/2 + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[24,currentCol/2 + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PowderBlue);
                                ws.Cells[24,currentCol/2 + 2].Style.Font.Bold = true;
                            }
                        }
                    }
                    ws.Cells.AutoFitColumns();
                    ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                }

                package.SaveAs(file);
            }

            MessageBox.Show("Merge is complete!");
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