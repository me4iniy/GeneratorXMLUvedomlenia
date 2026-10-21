using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace GeneratorXML
{
    public class TaxQuarter : INotifyPropertyChanged
    {
        private readonly ObservableCollection<TaxMount> _mounts = new ObservableCollection<TaxMount>();

        public TaxQuarter(string quarter)
        {
            Quarter = quarter;
        }

        public string Quarter { get; }
        public ObservableCollection<TaxMount> Mounts => _mounts;

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
            return ConvertToInt(Quarter).ToString();
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
