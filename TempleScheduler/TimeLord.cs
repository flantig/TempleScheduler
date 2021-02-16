using System.ComponentModel;
using System.Runtime.CompilerServices;
using TempleScheduler.Annotations;

namespace TempleScheduler
{

    public class TimeLord : INotifyPropertyChanged {
    protected void OnPropertyChanged(PropertyChangedEventArgs e)
    {
        PropertyChangedEventHandler handler = PropertyChanged;
        if (handler != null)
            handler(this, e);
    }

    public string flex;
    public string Flex
    {
        get { return flex;}
        set
        {
            flex = value;
            OnPropertyChanged();
        }
    }

    public string Time
    {
        get;
        set;
    }

    public event PropertyChangedEventHandler PropertyChanged;

    [NotifyPropertyChangedInvocator]
    public void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        if (PropertyChanged != null)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    }
}