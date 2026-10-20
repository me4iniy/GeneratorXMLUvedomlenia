using System.ComponentModel;

namespace GeneratorXML
{
    public class TaxLine : INotifyPropertyChanged
    {
        private string _oktmo = "";
        private string _kbk = "";
        private string _sum = "";
        private string _period = "";
        private string _nomerMesKvart = "";
        private string _year = "";

        public string Oktmo
        {
            get => _oktmo;
            set { _oktmo = value; OnPropertyChanged(nameof(Oktmo)); }
        }

        public string Kbk
        {
            get => _kbk;
            set { _kbk = value; OnPropertyChanged(nameof(Kbk)); }
        }

        public string Sum
        {
            get => _sum;
            set { _sum = value; OnPropertyChanged(nameof(Sum)); }
        }

        public string Period
        {
            get => _period;
            set { _period = value; OnPropertyChanged(nameof(Period)); }
        }

        public string NomerMesKvart
        {
            get => _nomerMesKvart;
            set { _nomerMesKvart = value; OnPropertyChanged(nameof(NomerMesKvart)); }
        }

        public string Year
        {
            get => _year;
            set { _year = value; OnPropertyChanged(nameof(Year)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}