using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneratorXML
{
    public class TaxMount : INotifyPropertyChanged
    {
        private readonly string _AffiliationWithQ;
        public TaxMount(string month, string parent)
        {
            _Month = month;
            _AffiliationWithQ = parent;
        }

        public readonly ObservableCollection<TaxLine> _taxLines = new ObservableCollection<TaxLine>();
        public readonly string _Month;

        public override string ToString()
        {
            int quarter = TaxQuarter.ConvertToInt(_AffiliationWithQ);
            int month = TaxMount.ConvertToInt(_Month);
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
