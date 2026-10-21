using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneratorXML
{
    public class TaxLine
    {
        public TaxLine(Type type, int oktmo, float sum, string period, string nomerMesKvart, int year)
        {
            switch (type)
            {
                case Type.Insuranse:
                    Kbk = "18210201000011000160";
                    break;
                case Type.NDFL:
                    Kbk = "18210102010011000110";
                    break;
                case Type.sNDFL:
                    Kbk = "18210102210011000110";
                    break;
            }

            Oktmo = oktmo;
            Sum = sum;
            Period = period;
            NomerMesKvart = nomerMesKvart;
            Year = year;
        }

        public enum Type { Insuranse, NDFL, sNDFL }

        public readonly int Oktmo;
        public readonly string Kbk;
        public readonly float Sum;
        public readonly string Period;
        public readonly string NomerMesKvart;
        public readonly int Year;
    }
}
