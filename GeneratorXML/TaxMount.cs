using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace GeneratorXML
{
    public class TaxMount : INotifyPropertyChanged
    {
        private readonly ObservableCollection<TaxLine> _taxLines = new ObservableCollection<TaxLine>();
        private readonly string _affiliationWithQuarter;

        public TaxMount(string month, string parentQuarter)
        {
            Month = month;
            _affiliationWithQuarter = parentQuarter;
        }

        public string Month { get; }
        public string ParentQuarter => _affiliationWithQuarter;
        public ObservableCollection<TaxLine> TaxLines => _taxLines;

        public override string ToString()
        {
            int quarter = TaxQuarter.ConvertToInt(_affiliationWithQuarter);
            int month = ConvertToInt(Month);
            return ((quarter - 1) * 3 + month).ToString();
        }

        public event PropertyChangedEventHandler PropertyChanged
        {
            add
            {
                ((INotifyPropertyChanged)_taxLines).PropertyChanged += value;
            }

            remove
            {
                ((INotifyPropertyChanged)_taxLines).PropertyChanged -= value;
            }
        }

        public static int ConvertToInt(string value)
        {
            return Convert.ToInt32(value) % 10;
        }
    }
}
