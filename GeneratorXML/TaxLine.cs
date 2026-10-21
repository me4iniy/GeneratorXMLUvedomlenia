using System;

namespace GeneratorXML
{
    public class TaxLine
    {
        public enum Type { Insuranse, NDFL, sNDFL }

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
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }

            LineType = type;
            Oktmo = oktmo;
            Sum = sum;
            Period = period;
            NomerMesKvart = nomerMesKvart;
            Year = year;
        }

        public Type LineType { get; }
        public int Oktmo { get; }
        public string Kbk { get; }
        public float Sum { get; }
        public string Period { get; }
        public string NomerMesKvart { get; }
        public int Year { get; }
    }
}
