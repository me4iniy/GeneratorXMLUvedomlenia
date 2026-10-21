using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Shapes;
using static OfficeOpenXml.ExcelErrorValue;

namespace GeneratorXML
{
    public class TaxPeriodsManager
    {
        private const string SheetNameUved = "УведомленияИПерсы";

        private int _oktmo_NDFL = 75506000;
        private int _oktmo_Insuranse = 75506000;
        private int _year = 2026;

        public readonly ObservableCollection<TaxQuarter> TaxQuarters = new ObservableCollection<TaxQuarter>();

        public void UpdateData(ExcelPackage package)
        {
            var worksheet = package.Workbook.Worksheets[SheetNameUved]
                ?? throw new Exception($"Лист '{SheetNameUved}' не найден.");

            //3 and 30 - fixed values in template
            for (int row = 3; row < 30; row++)
            {
                string[] period = worksheet.Cells[row, 2].Text?.Trim().Split('/');

                if (string.IsNullOrWhiteSpace(period[0]) || period.Length != 2)
                    continue;

                var taxSumLines = new[]
                {
                    (TaxLine.Type.NDFL, _oktmo_NDFL,worksheet.Cells[row, 3].GetValue<float?>()),
                    (TaxLine.Type.sNDFL, _oktmo_NDFL,worksheet.Cells[row, 4].GetValue<float?>()),
                    (TaxLine.Type.Insuranse, _oktmo_Insuranse,worksheet.Cells[row, 5].GetValue<float?>())
                };

                if (!taxSumLines.Any(x => !IsFlEmpty(x.Item3)))
                    continue;

                string quarter = period[0];
                string month = period[1];

                TaxQuarter taxQuarter = TaxQuarters.FirstOrDefault(x => x._Quarter == quarter);

                if (taxQuarter == null)
                {
                    taxQuarter = new TaxQuarter(quarter);
                    TaxQuarters.Add(taxQuarter);
                }

                TaxMount taxMount = taxQuarter._mounts.FirstOrDefault(x => x._Month == month);

                if (taxMount == null)
                {
                    taxMount = new TaxMount(month, quarter);
                    taxQuarter._mounts.Add(taxMount);
                }


                foreach (var (type, oktmo, sum) in taxSumLines)
                {
                    if (IsFlEmpty(sum))
                        continue;

                    taxMount._taxLines.Add(new TaxLine(type, oktmo, Convert.ToSingle(sum), quarter, month, _year));
                }
            }
        }
        
        private bool IsFlEmpty(float? value) => !value.HasValue || value == 0;
    }
}
