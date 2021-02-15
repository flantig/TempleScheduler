using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace TempleScheduler
{
   public class ExcelFormatter
    {
        public string path;
        public ExcelFormatter(string path)
        {
            this.path = path;
        }

        public void FileNames()
        {
            string[] fileEntries = Directory.GetFiles(this.path);
            Console.WriteLine(fileEntries[0]);
        }


    }
}
