using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Newtonsoft.Json;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using Brushes = System.Windows.Media.Brushes;
using MessageBox = System.Windows.MessageBox;

namespace TempleScheduler
{
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
            BindLife();
        }

        private void BindLife()
        {
            List<string> time = new List<string> { "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM", "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", };

            List<System.Windows.Controls.ListBox> listBoxes = new List<System.Windows.Controls.ListBox> { monday, tuesday, wednesday, jueves, friday };
            foreach (System.Windows.Controls.ListBox lister in listBoxes)
            {
                List<TimeLord> Bindables = new List<TimeLord>();

                foreach (string item in time)
                {
                    Bindables.Add(new TimeLord() { Flex = "off", Time = item });
                }

                lister.ItemsSource = Bindables;
            }

        }


        private void PathDialog_Event(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            tb.Text = dialog.SelectedPath;
        }

        private void SaveSchedule_Event(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(tb.Text))
            {
                MessageBox.Show("That's not a valid directory...");
                return;
            }

            List<System.Windows.Controls.ListBox> weekdays = new List<System.Windows.Controls.ListBox> { monday, tuesday, wednesday, jueves, friday};
            Schedule schedule = new Schedule();
            

            List<string> flexHours;
            List<string> normalHours;


            schedule.name = nameTB.Text;
            schedule.normalTimes = new List<List<string>>();
            schedule.flexTimes = new List<List<string>>();
            for (var i = 0; i < weekdays.Count; i++)
            {
            
                normalHours = new List<string>();
                flexHours = new List<string>();

                foreach (TimeLord item in weekdays[i].Items)
                {
                    if(item.Flex == "normal")
                    {
                        normalHours.Add(item.Time);
                    } else if (item.Flex == "flex")
                    {
                        flexHours.Add(item.Time);
                    }
                   
                   
                }
                
                
                schedule.normalTimes.Add(normalHours);
                schedule.flexTimes.Add(flexHours);
                
            }

            string json = JsonConvert.SerializeObject(schedule);
            System.IO.File.WriteAllText(tb.Text + $"\\{schedule.name}.json", json);
            MessageBox.Show("Save Complete!!");
        }


        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            TimeLord item = (TimeLord)(sender as TextBlock).DataContext;
            var textBlock = sender as TextBlock;

            switch (item.Flex)
            {
                case "off":
                    textBlock.Foreground = Brushes.Red;
                    item.Flex = "normal";
                    break;
                case "normal":
                    textBlock.Foreground = Brushes.Green;
                    item.Flex = "flex";
                    break;
                case "flex":
                    textBlock.Foreground = Brushes.Black;
                    item.Flex = "off";
                    break;
                
            }
        }

        private void Merge_OnClick(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(tb.Text))
            {
                ExcelFormatter export = new ExcelFormatter(tb.Text);
                export.ExcelCreator();
                MessageBox.Show("Merge is complete!");
            }
            else
            {
                MessageBox.Show("That's not a valid directory...");
            }
        }
    }
}
