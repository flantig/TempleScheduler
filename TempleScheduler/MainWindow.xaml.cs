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
using System;
using Control = System.Windows.Controls.Control;

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
            List<string> time = new List<string>
            {
                "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM",
                "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM",
                "5:00 PM", "5:30 PM",
            };

            List<System.Windows.Controls.ListBox> listBoxes = new List<System.Windows.Controls.ListBox>
                {monday, tuesday, wednesday, jueves, friday};
            foreach (System.Windows.Controls.ListBox lister in listBoxes)
            {
                List<TimeLord> Bindables = new List<TimeLord>();

                foreach (string item in time)
                {
                    Bindables.Add(new TimeLord() {Flex = "off", Time = item});
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

            List<System.Windows.Controls.ListBox> weekdays = new List<System.Windows.Controls.ListBox>
                {monday, tuesday, wednesday, jueves, friday};
            Schedule schedule = new Schedule();


            List<string> flexHours;

            List<string> normalHours;


            schedule.name = nameTB.Text;
            schedule.phone = phoneTB.Text;
            schedule.office = officeTB.Text;
            schedule.semester = semesterTB.Text;
            schedule.normalTimes = new List<List<string>>();
            schedule.flexTimes = new List<List<string>>();
            schedule.normalRanges = new List<IEnumerable<Tuple<int, int>>>();
            schedule.flexRanges = new List<IEnumerable<Tuple<int, int>>>();

            for (var i = 0; i < weekdays.Count; i++)
            {
                normalHours = new List<string>();
                var normalIndex = new List<int>();
                flexHours = new List<string>();
                var flexIndex = new List<int>();

                int index = 0;
                foreach (TimeLord item in weekdays[i].Items)
                {
                    if (item.Flex == "normal")
                    {
                        normalIndex.Add(index);
                        normalHours.Add(item.Time);
                    }
                    else if (item.Flex == "flex")
                    {
                        flexIndex.Add(index);
                        flexHours.Add(item.Time);
                    }

                    index++;
                }


                schedule.normalRanges.Add(numListToPossiblyDegenerateRanges(normalIndex));
                schedule.flexRanges.Add(numListToPossiblyDegenerateRanges(flexIndex));
                schedule.normalTimes.Add(normalHours);
                schedule.flexTimes.Add(flexHours);
            }

            string json = JsonConvert.SerializeObject(schedule);
            System.IO.File.WriteAllText(tb.Text + $"\\{schedule.name}_Schedules.json", json);
            MessageBox.Show("Save Complete!!");
        }

        public static IEnumerable<Tuple<int, int>> numListToPossiblyDegenerateRanges(IEnumerable<int> numList)
        {
            Tuple<int, int> currentRange = null;
            foreach (var num in numList)
            {
                if (currentRange == null)
                {
                    currentRange = Tuple.Create(num, num);
                }
                else if (currentRange.Item2 == num - 1)
                {
                    currentRange = Tuple.Create(currentRange.Item1, num);
                }
                else
                {
                    yield return currentRange;
                    currentRange = Tuple.Create(num, num);
                }
            }

            if (currentRange != null)
            {
                yield return currentRange;
            }
        }

        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            TimeLord item = (TimeLord) (sender as TextBlock).DataContext;
            var textBlock = sender as TextBlock;

            switch (item.Flex)
            {
                case "off":
                    item.Flex = "normal";
                    break;
                case "normal":
                    item.Flex = "flex";
                    break;
                case "flex":
                    item.Flex = "off";
                    break;
            }
        }


        private void Merge_OnClick(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(tb.Text))
            {
                ExcelFormatter export = new ExcelFormatter(tb.Text, semesterTB.Text);
                export.ExcelCreator();
                MessageBox.Show("Merge is complete!");
            }
            else
            {
                MessageBox.Show("That's not a valid directory...");
            }
        }

        public List<System.Windows.Controls.Control> AllChildren(DependencyObject parent)
        {
            var _List = new List<System.Windows.Controls.Control> { };
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var chil = VisualTreeHelper.GetChild(parent, i);
                if (chil is Control)
                    _List.Add(chil as Control);
                _List.AddRange(AllChildren(chil));
            }

            return _List;
        }

        private void clear_Click(object sender, RoutedEventArgs e)
        {
            List<System.Windows.Controls.ListBox> weekdays = new List<System.Windows.Controls.ListBox>
                {monday, tuesday, wednesday, jueves, friday};

            for (var i = 0; i < weekdays.Count; i++)
            {
                foreach (TimeLord item in weekdays[i].Items)
                {
                    item.Flex = "off";
                }
            }
        }
    }
}