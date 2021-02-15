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
        public List<List<string>> normalTimes;
        public List<List<string>> flexTimes;
        
    }
}