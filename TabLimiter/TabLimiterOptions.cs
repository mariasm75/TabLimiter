using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace TabLimiter
{
    internal class TabLimiterOptions : DialogPage, INotifyPropertyChanged
    {
        private int _maxNumberOfTabs;

        [Category("General")]
        [DisplayName("Tab limet")]
        [Description("Maximum number of open tabs")]
        public int MaxNumberOfTabs
        {
            get => _maxNumberOfTabs; set
            {
                if (value != _maxNumberOfTabs)
                {
                    _maxNumberOfTabs = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MaxNumberOfTabs)));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}