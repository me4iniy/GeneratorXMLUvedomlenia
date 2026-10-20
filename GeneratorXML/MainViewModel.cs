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
        private const string SheetName = "УведомленияИПерсы";

        private string _excelPath = @"C:\Users\kille\Dropbox\Рабочая\Табель расчетов\Аккуина Г.В..xlsm";
        private string _selectedPeriod = "";
        private ObservableCollection<string> _periods = new ObservableCollection<string>();
        private ObservableCollection<TaxLine> _taxLines = new ObservableCollection<TaxLine>();
        private string _fileId = "";
        private string _xmlPreview = "";
        private string _prefixRT = "UT_UVISCHSUMNAL";
        private string _a = "7460";
        private string _k = "7460";
        private string _o = "742600021392";

        private string _knd = "1110355";
        private string _dateDoc = DateTime.Today.ToString("dd.MM.yyyy");
        private string _kodNO = "7460";
        private string _innFL = "742601130309";
        private string _fam = "Вековшинина";
        private string _name = "Марина";
        private string _patronymic = "Анатольевна";
        private string _docName = "Доверенность";
        private string _versProg = "Налогоплательщик ЮЛ 4.94.1";
        private string _versForm = "5.03";
        private string _oktmo = "75506000";
        private string _prPodp = "2";

        public MainViewModel()
        {
            // Инициализация стартовых строк
            TaxLines = new ObservableCollection<TaxLine>
            {
                new TaxLine { Oktmo = "75506000", Kbk = "18210102010011000110", Sum = "0.00", Period = "31", NomerMesKvart = "02", Year = DateTime.Today.Year.ToString() },
                new TaxLine { Oktmo = "75506000", Kbk = "18210102210011000110", Sum = "0.00", Period = "31", NomerMesKvart = "02", Year = DateTime.Today.Year.ToString() },
                new TaxLine { Oktmo = "75506000", Kbk = "18210201000011000160", Sum = "0.00", Period = "31", NomerMesKvart = "02", Year = DateTime.Today.Year.ToString() }
            };

            LoadPeriodsFromExcel(_excelPath);

            SelectExcelCommand = new RelayCommand(_ => SelectExcel());
            LoadPeriodDataCommand = new RelayCommand(_ => LoadPeriodData());
            GenerateFileNameCommand = new RelayCommand(_ => GenerateFileName());
            GenerateXmlCommand = new RelayCommand(_ => GenerateXml());
        }

        public string ExcelPath
        {
            get => _excelPath;
            set { _excelPath = value; OnPropertyChanged(nameof(ExcelPath)); }
        }

        public string SelectedPeriod
        {
            get => _selectedPeriod;
            set { _selectedPeriod = value; OnPropertyChanged(nameof(SelectedPeriod)); }
        }

        public ObservableCollection<string> Periods
        {
            get => _periods;
            set { _periods = value; OnPropertyChanged(nameof(Periods)); }
        }

        public ObservableCollection<TaxLine> TaxLines
        {
            get => _taxLines;
            set { _taxLines = value; OnPropertyChanged(nameof(TaxLines)); }
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

        public string PrefixRT
        {
            get => _prefixRT;
            set { _prefixRT = value; OnPropertyChanged(nameof(PrefixRT)); }
        }

        public string A
        {
            get => _a;
            set { _a = value; OnPropertyChanged(nameof(A)); }
        }

        public string K
        {
            get => _k;
            set { _k = value; OnPropertyChanged(nameof(K)); }
        }

        public string O
        {
            get => _o;
            set { _o = value; OnPropertyChanged(nameof(O)); }
        }

        public string KND
        {
            get => _knd;
            set { _knd = value; OnPropertyChanged(nameof(KND)); }
        }

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

        public string Fam
        {
            get => _fam;
            set { _fam = value; OnPropertyChanged(nameof(Fam)); }
        }

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        public string Patronymic
        {
            get => _patronymic;
            set { _patronymic = value; OnPropertyChanged(nameof(Patronymic)); }
        }

        public string DocName
        {
            get => _docName;
            set { _docName = value; OnPropertyChanged(nameof(DocName)); }
        }

        public string VersProg
        {
            get => _versProg;
            set { _versProg = value; OnPropertyChanged(nameof(VersProg)); }
        }

        public string VersForm
        {
            get => _versForm;
            set { _versForm = value; OnPropertyChanged(nameof(VersForm)); }
        }

        public string Oktmo
        {
            get => _oktmo;
            set { _oktmo = value; OnPropertyChanged(nameof(Oktmo)); }
        }

        public string PrPodp
        {
            get => _prPodp;
            set { _prPodp = value; OnPropertyChanged(nameof(PrPodp)); }
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
                InitialDirectory = Path.GetDirectoryName(ExcelPath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dialog.ShowDialog() == true)
            {
                ExcelPath = dialog.FileName;
                LoadPeriodsFromExcel(ExcelPath);
            }
        }

        private void LoadPeriodsFromExcel(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                    return;

                var periods = new List<string>();

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[SheetName];
                    if (worksheet == null)
                        throw new Exception($"Лист '{SheetName}' не найден.");

                    for (int row = 3; row < 5000; row++)
                    {
                        var cell = worksheet.Cells[row, 2]; // B
                        var value = cell.Text?.Trim();

                        if (string.IsNullOrWhiteSpace(value))
                        {
                            // Пустая строка может быть разделителем между кварталами,
                            // поэтому пропускаем её, а не прерываем чтение файла.
                            continue;
                        }

                        if (!periods.Contains(value))
                            periods.Add(value);
                    }
                }

                Periods = new ObservableCollection<string>(periods);

                if (Periods.Count > 0)
                    SelectedPeriod = Periods[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка чтения листа '{SheetName}': {ex.Message}", "Ошибка Excel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadPeriodData()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ExcelPath) || !File.Exists(ExcelPath))
                {
                    MessageBox.Show("Выберите корректный Excel-файл.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(SelectedPeriod))
                {
                    MessageBox.Show("Выберите период.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var rows = ReadPeriodRows(ExcelPath, SelectedPeriod);

                if (rows == null || rows.Count == 0)
                {
                    MessageBox.Show("Для выбранного периода данные не найдены.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
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

        private List<TaxLine> ReadPeriodRows(string filePath, string selectedPeriod)
        {
            var result = new List<TaxLine>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[SheetName];
                if (worksheet == null)
                    throw new Exception($"Лист '{SheetName}' не найден.");

                for (int row = 3; row < 5000; row++)
                {
                    var periodCell = worksheet.Cells[row, 2]; // B
                    var periodText = periodCell.Text?.Trim();

                    if (string.IsNullOrWhiteSpace(periodText))
                    {
                        // Пустая строка — не конец файла, а просто разделитель между блоками.
                        continue;
                    }

                    if (!string.Equals(periodText, selectedPeriod, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var valueC = worksheet.Cells[row, 3].GetValue<decimal?>();
                    var valueD = worksheet.Cells[row, 4].GetValue<decimal?>();
                    var valueE = worksheet.Cells[row, 5].GetValue<decimal?>();

                    var periodCode = GetPeriodCode(periodText);
                    var monthCode = GetNomerMesKvart(periodText);
                    var year = DateTime.Today.Year.ToString();

                    if (valueC.HasValue)
                    {
                        result.Add(new TaxLine
                        {
                            Oktmo = Oktmo,
                            Kbk = "18210102010011000110",
                            Sum = valueC.Value.ToString("0.00", CultureInfo.InvariantCulture),
                            Period = periodCode,
                            NomerMesKvart = monthCode,
                            Year = year
                        });
                    }

                    if (valueD.HasValue)
                    {
                        result.Add(new TaxLine
                        {
                            Oktmo = Oktmo,
                            Kbk = "18210102210011000110",
                            Sum = valueD.Value.ToString("0.00", CultureInfo.InvariantCulture),
                            Period = periodCode,
                            NomerMesKvart = monthCode,
                            Year = year
                        });
                    }

                    if (valueE.HasValue)
                    {
                        result.Add(new TaxLine
                        {
                            Oktmo = Oktmo,
                            Kbk = "18210201000011000160",
                            Sum = valueE.Value.ToString("0.00", CultureInfo.InvariantCulture),
                            Period = periodCode,
                            NomerMesKvart = monthCode,
                            Year = year
                        });
                    }

                    // Если в файле один период = одна строка, можно оставить break;
                    // если строк с этим периодом несколько — не прерываем чтение.
                    // break;
                }
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
