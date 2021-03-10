using System;
using System.Collections.Generic;

namespace TempleScheduler
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    [Serializable]
    public class Schedule
    {
        public string name;
        public string phone;
        public string office;
        public string semester;
        public List<List<string>> normalTimes;
        public List<List<string>> flexTimes;
        public List<IEnumerable<Tuple<int, int>>> flexRanges;
        public List<IEnumerable<Tuple<int, int>>> normalRanges;

    }
}