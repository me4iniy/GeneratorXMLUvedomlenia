using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using System.Xml.Linq;

namespace GeneratorXML
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public const string SenderINN = "742600021392";
        public const string SenderFam = "Вековшинина";
        public const string SenderName = "Марина";
        public const string SenderPatronymic = "Анатольевна";
        public const string DocNameDefault = "Доверенность";
        public const string PrPodpDefault = "2";

        public const string PrefixRTDefault = "UT_UVISCHSUMNAL";
        public const string KNDDefault = "1110355";
        public const string VersProgDefault = "Налогоплательщик ЮЛ 4.94.1";
        public const string VersFormDefault = "5.03";

        private readonly TaxPeriodsManager _taxPeriodsManager = new TaxPeriodsManager();

        private string _dataPath = @"C:\Users\kille\Dropbox\Рабочая\Табель расчетов\Аккуина Г.В..xlsm";
        private TaxQuarter _selectedQuarter;
        private string _fileId = "";
        private string _xmlPreview = "";
        private string _fromFNS = "7460";
        private string _toFNS = "7460";
        private string _dateDoc = DateTime.Today.ToString("dd.MM.yyyy");
        private string _kodNO = "7460";
        private string _innFL = "742601130309";

        public string DataPath
        {
            get => _dataPath;
            set { _dataPath = value; OnPropertyChanged(nameof(DataPath)); }
        }

        public ObservableCollection<TaxQuarter> TaxQuarters => _taxPeriodsManager.TaxQuarters;

        public ObservableCollection<TaxLine> TaxLines { get; private set; } = new ObservableCollection<TaxLine>();

        public TaxQuarter SelectedQuarter
        {
            get => _selectedQuarter;
            set
            {
                if (Equals(_selectedQuarter, value))
                    return;

                _selectedQuarter = value;
                OnPropertyChanged(nameof(SelectedQuarter));
                LoadSelectedQuarterData();
            }
        }

        public string FileId
        {
            get => _fileId;
            set { _fileId = value; OnPropertyChanged(nameof(FileId)); OnPropertyChanged(nameof(FileNameLength)); }
        }

        public int FileNameLength => string.IsNullOrEmpty(FileId) ? 0 : FileId.Length;

        public string XmlPreview
        {
            get => _xmlPreview;
            set { _xmlPreview = value; OnPropertyChanged(nameof(XmlPreview)); }
        }

        public string FromFNS
        {
            get => _fromFNS;
            set { _fromFNS = value; OnPropertyChanged(nameof(FromFNS)); }
        }

        public string ToFNS
        {
            get => _toFNS;
            set { _toFNS = value; OnPropertyChanged(nameof(ToFNS)); }
        }

        public string KND { get; set; } = KNDDefault;
        public string DateDoc
        {
            get => _dateDoc;
            set { _dateDoc = value; OnPropertyChanged(nameof(DateDoc)); }
        }

        public string KodNO
        {
            get => _kodNO;
            set { _kodNO = value; OnPropertyChanged(nameof(KodNO)); }
        }

        public string InnFL
        {
            get => _innFL;
            set { _innFL = value; OnPropertyChanged(nameof(InnFL)); }
        }

        public string Oktmo { get; set; } = "75506000";

        public MainViewModel()
        {
            SelectExcelCommand = new RelayCommand(_ => SelectExcel());
            LoadPeriodDataCommand = new RelayCommand(_ => LoadPeriodData());
            GenerateFileNameCommand = new RelayCommand(_ => GenerateFileName());
            GenerateXmlCommand = new RelayCommand(_ => GenerateXml());

            if (TaxQuarters.Any())
                SelectedQuarter = TaxQuarters.First();
        }

        public ICommand SelectExcelCommand { get; }
        public ICommand LoadPeriodDataCommand { get; }
        public ICommand GenerateFileNameCommand { get; }
        public ICommand GenerateXmlCommand { get; }

        private void SelectExcel()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsm)|*.xlsm|Excel files (*.xlsx)|*.xlsx",
                Title = "Выберите файл Excel",
                InitialDirectory = Path.GetDirectoryName(DataPath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dialog.ShowDialog() == true)
            {
                DataPath = dialog.FileName;
                LoadPeriodsFromExcel(DataPath);
            }
        }

        private void LoadPeriodsFromExcel(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    return;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    _taxPeriodsManager.UpdateData(package);
                }

                OnPropertyChanged(nameof(TaxQuarters));
                SelectedQuarter = TaxQuarters.FirstOrDefault();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка чтения: {ex.Message}", "Ошибка Excel",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadPeriodData()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(DataPath) || !File.Exists(DataPath))
                {
                    MessageBox.Show("Выберите корректный Excel-файл.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                using (var package = new ExcelPackage(new FileInfo(DataPath)))
                {
                    _taxPeriodsManager.UpdateData(package);
                }

                OnPropertyChanged(nameof(TaxQuarters));
                SelectedQuarter = TaxQuarters.FirstOrDefault();

                MessageBox.Show($"Данные загружены. Найдено кварталов: {TaxQuarters.Count}", "Готово",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка чтения Excel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadSelectedQuarterData()
        {
            TaxLines.Clear();

            if (SelectedQuarter == null)
                return;

            foreach (var taxLine in SelectedQuarter.Mounts.SelectMany(m => m.TaxLines))
                TaxLines.Add(taxLine);
        }

        private void GenerateFileName()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(FromFNS))
                    throw new Exception("Поле A не должно быть пустым.");

                if (string.IsNullOrWhiteSpace(ToFNS))
                    throw new Exception("Поле K не должно быть пустым.");

                var datePart = DateTime.Today.ToString("yyyyMMdd");
                var guid = Guid.NewGuid().ToString("D").ToUpperInvariant();

                FileId = $"{PrefixRTDefault}_{FromFNS}_{ToFNS}_{SenderINN}_{datePart}_{guid}";

                MessageBox.Show($"Имя файла сформировано:\n{FileId}\n\nДлина: {FileNameLength}", "Генерация имени",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка генерации имени", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateXml()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(FileId))
                    throw new Exception("Сначала сформируйте имя файла.");

                if (TaxLines == null || TaxLines.Count == 0)
                    throw new Exception("Нельзя генерировать XML без строк.");

                var xml = BuildXml(FileId, TaxLines);
                XmlPreview = xml;

                var saveDialog = new SaveFileDialog
                {
                    Filter = "XML files (*.xml)|*.xml",
                    Title = "Сохранить XML",
                    FileName = FileId + ".xml"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    File.WriteAllText(saveDialog.FileName, xml, Encoding.GetEncoding(1251));
                    MessageBox.Show($"Файл сохранён:\n{saveDialog.FileName}", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка генерации XML", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string BuildXml(string fileId, IEnumerable<TaxLine> lines)
        {
            var document = new XDocument(
                new XDeclaration("1.0", "windows-1251", null),
                new XElement("Файл",
                    new XAttribute(XNamespace.Xmlns + "xsi", "http://www.w3.org/2001/XMLSchema-instance"),
                    new XAttribute("ИдФайл", fileId),
                    new XAttribute("ВерсПрог", VersProgDefault),
                    new XAttribute("ВерсФорм", VersFormDefault),
                    new XElement("Документ",
                        new XAttribute("КНД", KND),
                        new XAttribute("ДатаДок", DateDoc),
                        new XAttribute("КодНО", KodNO),
                        new XElement("СвНП",
                            new XElement("НПИП", new XAttribute("ИННФЛ", InnFL))
                        ),
                        new XElement("Подписант",
                            new XAttribute("ПрПодп", PrPodpDefault),
                            new XElement("ФИО",
                                new XAttribute("Фамилия", SenderFam),
                                new XAttribute("Имя", SenderName),
                                new XAttribute("Отчество", SenderPatronymic)
                            ),
                            new XElement("СвПред", new XAttribute("НаимДок", DocNameDefault))
                        ),
                        lines.Select(line =>
                            new XElement("УвИсчСумНалог",
                                new XAttribute("ОКТМО", line.Oktmo),
                                new XAttribute("КБК", line.Kbk),
                                new XAttribute("СумНалогАванс", line.Sum.ToString("0.00", CultureInfo.InvariantCulture)),
                                new XAttribute("Период", line.Period),
                                new XAttribute("НомерМесКварт", line.NomerMesKvart),
                                new XAttribute("Год", line.Year)
                            )
                        )
                    )
                )
            );

            var sb = new StringBuilder();
            var settings = new XmlWriterSettings
            {
                Encoding = Encoding.GetEncoding(1251),
                OmitXmlDeclaration = false,
                Indent = false,
                NewLineHandling = NewLineHandling.None
            };

            using (var writer = XmlWriter.Create(sb, settings))
            {
                document.Save(writer);
            }

            return sb.ToString();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
