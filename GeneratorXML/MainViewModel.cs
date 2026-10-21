using GeneratorXML;
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
        public const string _fam_SenderSur = "Вековшинина";
        public const string _name_SenderName = "Марина";
        public const string _patronymic_SenderSecName = "Анатольевна";
        public const string _docName_SenderDoc = "Доверенность";
        public const string _prPodp_SenderStatus = "2";

        public const string PrefixRT = "UT_UVISCHSUMNAL";
        public const string KND = "1110355";

        public const string VersProg = "Налогоплательщик ЮЛ 4.94.1";
        public const string VersForm = "5.03";

        private string _dataPath = @"C:\Users\kille\Dropbox\Рабочая\Табель расчетов\Аккуина Г.В..xlsm";
        public string DataPath
        {
            get { return _dataPath; }
            set {  _dataPath = value; OnPropertyChanged(nameof(DataPath)); }
        }

        public ObservableCollection<TaxLine> TaxLines = new ObservableCollection<TaxLine>();
        public ObservableCollection<TaxMount> TaxMounts = new ObservableCollection<TaxMount>();
        public ObservableCollection<TaxQuarter> TaxQuarters = new ObservableCollection<TaxQuarter>();

        private TaxPeriodsManager _TaxPeriodsManager = new TaxPeriodsManager();

        private string _fileId = "";
        public string FileId
        {
            get { return _fileId; }
            set { _fileId = value; OnPropertyChanged(nameof(FileId));}
        }
        private string _xmlPreview = "";
        public string XmlPreview
        {
            get { return _xmlPreview; }
            set { _xmlPreview = value; OnPropertyChanged(nameof(XmlPreview)); }
        }
        private string _fromFNS = "7460";
        public string FromFNS
        {
            get { return _fromFNS; }
            set { _fromFNS = value; OnPropertyChanged(nameof(FromFNS)); }
        }
        private string _toFNS = "7460";
        public string ToFNS
        {
            get { return _toFNS; }
            set { _toFNS = value; OnPropertyChanged(nameof(ToFNS)); }
        }
        private string _dateDoc = DateTime.Today.ToString("dd.MM.yyyy");
        private string _kodNO = "7460";
        private string _innFL = "742601130309";
        public string InnFL
        {
            get { return _innFL; }
            set { _innFL = value; OnPropertyChanged(nameof(InnFL)); }
        }

        private int _selectedQuarter = -1;
        public int SelectedQuarter
        {
            get { return _selectedQuarter; }
            set { _selectedQuarter = value; OnPropertyChanged(nameof(SelectedQuarter)); }
        }
        private int _selectedMonth = -1;
        public int SelectedMonth
        {
            get { return _selectedMonth; }
            set { _selectedMonth = value; OnPropertyChanged(nameof(SelectedMonth)); }
        }


        public MainViewModel()
        {
            SelectExcelCommand = new RelayCommand(_ => SelectExcel());
            LoadPeriodDataCommand = new RelayCommand(_ => LoadPeriodData());
            GenerateFileNameCommand = new RelayCommand(_ => GenerateFileName());
            GenerateXmlCommand = new RelayCommand(_ => GenerateXml());
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
                    _TaxPeriodsManager.UpdateData(package);

                //SelectedMonth = -1;
                //SelectedQuarter = -1;
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

                TaxLines = new ObservableCollection<TaxLine>(rows);
                MessageBox.Show($"Данные загружены за период: {SelectedPeriod}", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка чтения Excel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<TaxLine> ReadMonthRows(int selectedMonth)
        {
            List<TaxLine> result = new List<TaxLine>();

            using (var package = new ExcelPackage(new FileInfo(ExcelPath)))
            {
                var worksheet = package.Workbook.Worksheets[SheetNameUved]
                    ?? throw new Exception($"Лист '{SheetNameUved}' не найден.");

                for (int row = 3; row < 30; row++)
                {
                    var periodText = worksheet.Cells[row, 2].Text?.Trim();

                    if (string.IsNullOrWhiteSpace(periodText))
                        continue;

                    int normalizedMonth = NormalizeMonthKey(periodText);

                    if (normalizedMonth != selectedMonth)
                        continue;

                    float? valueC = worksheet.Cells[row, 3].GetValue<float?>();
                    float? valueD = worksheet.Cells[row, 4].GetValue<float?>();
                    float? valueE = worksheet.Cells[row, 5].GetValue<float?>();

                    if (valueC.HasValue) totalC += valueC.Value;
                    if (valueD.HasValue) totalD += valueD.Value;
                    if (valueE.HasValue) totalE += valueE.Value;

                    if (string.IsNullOrEmpty(monthCode))
                    {
                        monthCode = GetNomerMesKvart(periodText);
                        periodCode = GetPeriodCode(periodText);
                    }
                }

                if (totalC != 0)
                    result.Add(new TaxLine
                    {
                        Oktmo = Oktmo,
                        Kbk = "18210102010011000110",
                        Sum = totalC.ToString("0.00", CultureInfo.InvariantCulture),
                        Period = periodCode,
                        NomerMesKvart = monthCode,
                        Year = DateTime.Today.Year.ToString()
                    });

                if (totalD != 0)
                    result.Add(new TaxLine
                    {
                        Oktmo = Oktmo,
                        Kbk = "18210102210011000110",
                        Sum = totalD.ToString("0.00", CultureInfo.InvariantCulture),
                        Period = periodCode,
                        NomerMesKvart = monthCode,
                        Year = DateTime.Today.Year.ToString()
                    });

                if (totalE != 0)
                    result.Add(new TaxLine
                    {
                        Oktmo = Oktmo,
                        Kbk = "18210201000011000160",
                        Sum = totalE.ToString("0.00", CultureInfo.InvariantCulture),
                        Period = periodCode,
                        NomerMesKvart = monthCode,
                        Year = DateTime.Today.Year.ToString()
                    });
            }

            return result;
        }

        private string GetPeriodCode(string periodText)
        {
            // Делает из "21/01" -> "31" по схеме, у которой вы описали квартальные значения.
            // В реальном проекте можно сделать более точную формулу: 21=>31, 31=>31, 33=>31, 34=>31.
            if (string.IsNullOrWhiteSpace(periodText))
                return "31";

            var text = periodText.Trim();
            if (text.Contains("/"))
            {
                var parts = text.Split('/');
                if (parts.Length >= 2)
                {
                    var first = parts[0].Trim();
                    if (first == "21" || first == "31" || first == "33" || first == "34")
                        return "31";
                }
            }

            return "31";
        }

        private string GetNomerMesKvart(string periodText)
        {
            // Из строки вида "21/01" делает "01"
            // Из "31/02" -> "02"
            // По сути: месяц внутри квартала.
            if (string.IsNullOrWhiteSpace(periodText))
                return "01";

            var text = periodText.Trim();
            if (text.Contains("/"))
            {
                var parts = text.Split('/');
                if (parts.Length >= 2)
                {
                    var month = parts[1].Trim();
                    if (int.TryParse(month, out _))
                        return month.PadLeft(2, '0');
                }
            }

            return "01";
        }

        private void GenerateFileName()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(PrefixRT))
                    throw new Exception("Поле R_T не должно быть пустым.");

                if (string.IsNullOrWhiteSpace(A))
                    throw new Exception("Поле A не должно быть пустым.");

                if (string.IsNullOrWhiteSpace(K))
                    throw new Exception("Поле K не должно быть пустым.");

                if (string.IsNullOrWhiteSpace(O))
                    throw new Exception("Поле O не должно быть пустым.");

                var datePart = DateTime.Today.ToString("yyyyMMdd");
                var guid = Guid.NewGuid().ToString("D").ToUpperInvariant();

                FileId = $"{PrefixRT}_{A}_{K}_{O}_{datePart}_{guid}";
                OnPropertyChanged(nameof(FileNameLength));

                MessageBox.Show($"Имя файла сформировано:\n{FileId}\n\nДлина: {FileNameLength}", "Генерация имени", MessageBoxButton.OK, MessageBoxImage.Information);
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
                    new XAttribute("ВерсПрог", VersProg),
                    new XAttribute("ВерсФорм", VersForm),
                    new XElement("Документ",
                        new XAttribute("КНД", KND),
                        new XAttribute("ДатаДок", DateDoc),
                        new XAttribute("КодНО", KodNO),
                        new XElement("СвНП",
                            new XElement("НПИП", new XAttribute("ИННФЛ", InnFL))
                        ),
                        new XElement("Подписант",
                            new XAttribute("ПрПодп", PrPodp),
                            new XElement("ФИО",
                                new XAttribute("Фамилия", Fam),
                                new XAttribute("Имя", Name),
                                new XAttribute("Отчество", Patronymic)
                            ),
                            new XElement("СвПред", new XAttribute("НаимДок", DocName))
                        ),
                        lines.Select(line =>
                            new XElement("УвИсчСумНалог",
                                new XAttribute("ОКТМО", line.Oktmo),
                                new XAttribute("КБК", line.Kbk),
                                new XAttribute("СумНалогАванс", line.Sum),
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
