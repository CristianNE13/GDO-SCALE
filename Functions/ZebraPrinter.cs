using System;
using System.Globalization;
using System.Text;

namespace Scale_Program.Functions
{
    internal class ZebraPrinter
    {
        private const string ZplHeader =
            "\u0010CT~~CD,~CC^~CT~\r\n^XA\r\n~TA000\r\n~JSN\r\n^LT0\r\n^MNW\r\n^MTT\r\n^PON\r\n^PMN\r\n^LH0,0\r\n^JMA\r\n^PR4,4\r\n~SD30\r\n^JUS\r\n^LRN\r\n^CI27\r\n^PA0,1,1,0\r\n^XZ\r\n^XA\r\n^MMT\r\n^PW168\r\n^LL328\r\n^LS0\r\n^FT42,283^BQN,2,4\r\n";

        public string JualianDate { get; }

        public static (string zlp, string integerPart, string fractionalPart) GenerateZplBody(string pieza)
        {
            var (integer, fractional) = SplitJulianDate(GetJulianDayNumber(DateTime.Now));
            var sb = new StringBuilder();
            sb.AppendLine(ZplHeader);
            sb.AppendLine($"^FH\\^FDLA2{integer}.{fractional}^FS\r\n");
            sb.AppendLine("^FT71,179^A0B,34,33^FH\\^CI28^FDPieza:^FS^CI27\r\n");
            sb.AppendLine($"^FT96,179^A0B,11,13^FH\\^CI28^FD{pieza}^FS^CI27\r\n");
            sb.AppendLine("^PQ1,0,1,Y\r\n^XZ");

            return (sb.ToString(), integer, fractional);
        }

        public static double GetJulianDayNumber(DateTime date)
        {
            var year = date.Year;
            var month = date.Month;
            var day = date.Day;

            if (month <= 2)
            {
                year -= 1;
                month += 12;
            }

            var A = year / 100;
            var B = 2 - A + A / 4;

            var jd = Math.Floor(365.25 * (year + 4716)) +
                Math.Floor(30.6001 * (month + 1)) +
                day + B - 1524.5;

            var fraction = (date.Hour + date.Minute / 60.0 + date.Second / 3600.0) / 24.0;
            return jd + fraction;
        }

        public static (string, string) SplitJulianDate(double julianDate)
        {
            var julianDateString = julianDate.ToString("F6", CultureInfo.InvariantCulture);
            var parts = julianDateString.Split('.');
            return (parts[0], parts.Length > 1 ? parts[1] : "000000");
        }
    }
}