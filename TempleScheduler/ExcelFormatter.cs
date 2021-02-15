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

namespace TempleScheduler
{
   public class ExcelFormatter
    {
        public string path;

        public Hashtable cities = new Hashtable()
        {
            {"8:00 AM",""},
        };

        public ExcelFormatter(string path)
        {

            this.path = path;
        }

        public List<Schedule> PersonsJSONDeserializer()
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

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DateTime today = DateTime.Today;
            string outputFile = $"{this.path}\\{today.ToString("yyyyMMddss")}_Schedules.xlsx";
            Console.WriteLine(outputFile);
            var file = new FileInfo(outputFile);


            var staff = PersonsJSONDeserializer();

            /*
             * The using keyword allows us to use a file and not worrying about closing it manually later. The old school way of doing this would have been
             * package.Dispose() or in VBA it'd be Workbooks(file).Close
             */
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add("Monday");
                package.SaveAs(file);
            }



        }




    }
}
