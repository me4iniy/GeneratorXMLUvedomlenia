using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneratorXML
{
    public class TaxQuarter : INotifyPropertyChanged
    {
        public readonly ObservableCollection<TaxMount> _mounts = new ObservableCollection<TaxMount>();
        public TaxQuarter(string quarter)
        {
            _Quarter = quarter;
        }

        public readonly string _Quarter;

        public event PropertyChangedEventHandler PropertyChanged
        {
            add
            {
                ((INotifyPropertyChanged)_mounts).PropertyChanged += value;
            }

            remove
            {
                ((INotifyPropertyChanged)_mounts).PropertyChanged -= value;
            }
        }

        public override string ToString()
        {
            return TaxQuarter.ConvertToInt(_Quarter).ToString();
        }

        public static int ConvertToInt(string value)
        {
            switch (value)
            {
                case "21":
                    return 1;
                case "31":
                    return 2;
                case "33":
                    return 3;
                case "34":
                    return 4;
                default:
                    throw new Exception($"Код квартала '{value}', указан не верно.");
            }
        }
    }
}
