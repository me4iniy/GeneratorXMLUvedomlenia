using OfficeOpenXml;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace GeneratorXML
{
    public class TaxPeriodsManager
    {
        private const string SheetNameUved = "УведомленияИПерсы";

        private readonly int _oktmoNDFL = 75506000;
        private readonly int _oktmoInsuranse = 75506000;
        private readonly int _year = 2026;

        public ObservableCollection<TaxQuarter> TaxQuarters { get; } = new ObservableCollection<TaxQuarter>();

        public void UpdateData(ExcelPackage package)
        {
            var worksheet = package.Workbook.Worksheets[SheetNameUved]
                ?? throw new Exception($"Лист '{SheetNameUved}' не найден.");

            TaxQuarters.Clear();

            for (int row = 3; row < 30; row++)
            {
                string periodValue = worksheet.Cells[row, 2].Text?.Trim();
                if (string.IsNullOrWhiteSpace(periodValue))
                    continue;

                string[] periodParts = periodValue.Split('/');
                if (periodParts.Length != 2 || string.IsNullOrWhiteSpace(periodParts[0]))
                    continue;

                string quarter = periodParts[0].Trim();
                string month = periodParts[1].Trim();

                var taxSumLines = new[]
                {
                    (Type: TaxLine.Type.NDFL, Oktmo: _oktmoNDFL, Sum: worksheet.Cells[row, 3].GetValue<float?>()),
                    (Type: TaxLine.Type.sNDFL, Oktmo: _oktmoNDFL, Sum: worksheet.Cells[row, 4].GetValue<float?>()),
                    (Type: TaxLine.Type.Insuranse, Oktmo: _oktmoInsuranse, Sum: worksheet.Cells[row, 5].GetValue<float?>())
                };

                if (!taxSumLines.Any(x => !IsFlEmpty(x.Sum)))
                    continue;

                var quarterItem = TaxQuarters.FirstOrDefault(x => x.Quarter == quarter);
                if (quarterItem == null)
                {
                    quarterItem = new TaxQuarter(quarter);
                    TaxQuarters.Add(quarterItem);
                }

                var mountItem = quarterItem.Mounts.FirstOrDefault(x => x.Month == month);
                if (mountItem == null)
                {
                    mountItem = new TaxMount(month, quarter);
                    quarterItem.Mounts.Add(mountItem);
                }

                foreach (var (type, oktmo, sum) in taxSumLines)
                {
                    if (IsFlEmpty(sum))
                        continue;

                    mountItem.TaxLines.Add(new TaxLine(type, oktmo, Convert.ToSingle(sum), quarter, month, _year));
                }
            }
        }

        private static bool IsFlEmpty(float? value) => !value.HasValue || value == 0;
    }
}
