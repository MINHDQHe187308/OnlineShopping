using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Reflection;

namespace ASP.BaseCommon
{
    public static class Helper
    {
        public static string GetDescription(this Enum value)
        {
            var field = value.GetType().GetField(value.ToString());
            var attr = field.GetCustomAttributes(typeof(DescriptionAttribute), false);
            return attr.Length == 0 ? value.ToString() : (attr[0] as DescriptionAttribute).Description;
        }

        public static string? GetDisplayName(this Enum enumeration)
        {
            return enumeration.GetType()
                    .GetMember(enumeration.ToString())
                    .First()
                    .GetCustomAttribute<DisplayAttribute>()
                    ?.GetName();
        }

        public static string ConvertStringToStringFromExcel(string value) 
        { 
            if(string.IsNullOrEmpty(value)) return "";
            return value.Trim();
        }

        public static short ConvertStringToShortFromExcel(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            return Convert.ToInt16(value.Replace(" ", ""));
        }

        public static int ConvertStringToIntFromExcel(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            return Convert.ToInt32(value.Replace(" ", ""));
        }

        public static int ConvertStringToIntWithNegativeFromExcel(string value, bool isCheckNegative)
        {
            if (string.IsNullOrEmpty(value)) return 0;

            if (isCheckNegative)
            {
                var sign = value.Contains("-");

                if (sign)
                {
                    var newVal = value.Replace("-", "");
                    var val = Convert.ToInt32(newVal.Replace(" ", ""));
                    return val * -1;
                }
            }

            return Convert.ToInt32(value.Replace(" ", ""));
        }

        public static float ConvertStringToFloatFromExcel(string value) 
        {
            if (string.IsNullOrEmpty(value)) return 0;
            return Convert.ToSingle(value.Replace(" ", ""));
        }

        public static double ConvertStringToDoubleFromExcel(string value, bool isCheckNegative)
        {
            if (string.IsNullOrEmpty(value)) return 0;

            if(isCheckNegative)
            {
                var sign = value.Contains("-");

                if(sign)
                {
                    var newVal = value.Replace("-", "");
                    var val = Convert.ToDouble(newVal.Replace(" ", ""));
                    return val * -1;
                }
            }

            return Convert.ToDouble(value.Replace(" ", ""));
        }

        public static bool ConvertStringToBoolFromExcel(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            return Convert.ToBoolean(value);
        }

        public static bool CheckFileIsExcel(string extension)
        {
            if (extension.Equals(".xlsx")) return true;
            if (extension.Equals(".xls")) return true;
            if (extension.Equals(".xlsm")) return true;
            return false;
        }

        public static bool CheckFileIsPdf(string extension)
        {
            if (extension.Equals(".pdf")) return true;
            return false;
        }

        public static bool CheckStringContains(string main, string compare)
        {
            if (main.Contains(compare)) return true;
            return false;
        }

        public static bool CheckStringStartWith(string main, string compare)
        {
            if (main.StartsWith(compare)) return true;
            return false;
        }

        public static DateTime ConvertStringToDateFromExcel(string value)
        {
            if (string.IsNullOrEmpty(value)) return DateTime.Today;
            return DateTime.ParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        public static DateTime ConvertStringToDateTimeFromExcel(string value)
        {
            if (string.IsNullOrEmpty(value)) return DateTime.Now;
            return DateTime.ParseExact(value, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        public static DateTime ConvertStringToDateTimeFromExcelWithFormat(string value, string format)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(format)) return DateTime.Now;
            return DateTime.ParseExact(value, format, CultureInfo.InvariantCulture);
        }

        public static string GenFileName(string fileName, string extension)
        {
            return $"{fileName}_{DateTime.Now.ToString("yyMMdd_HHmmss")}.{extension}";
        }

        public static int GetFYByMonthYear(int inputYear, int inputMonth)
        {
            if (inputMonth < 4) return inputYear - 1;
            return inputYear;
        }

        public static decimal CalcPercentage(int value, int total)
        {
            if (total == 0) return 0;
            decimal res = (100 * value) / total;
            return Math.Round(res);
        }

        public static List<int> ConvertStringToListInt(string inputString)
        {
            if (string.IsNullOrEmpty(inputString)) return new List<int>();

            var raw = inputString.Split(',', System.StringSplitOptions.RemoveEmptyEntries);
            if (raw == null || raw.Length == 0) return new List<int> { };

            var lstItems = new List<int>();
            int tempVal = 0;

            foreach (var item in raw)
            {
                int.TryParse(item, out tempVal);
                lstItems.Add(tempVal);
            }

            return lstItems;
        }

        public static List<Guid> ConvertStringToGuidList(string inputString) {
            if (string.IsNullOrEmpty(inputString)) return new List<Guid>();

            var raw = inputString.Split(',', System.StringSplitOptions.RemoveEmptyEntries);
            if (raw == null || raw.Length == 0) return new List<Guid> { };

            var lstItems = new List<Guid>();
            Guid tempVal = Guid.Empty;

            foreach (var item in raw)
            {
                Guid.TryParse(item, out tempVal);
                lstItems.Add(tempVal);
            }

            return lstItems;
        }

        public static int GetInventoryMonthByThisMonth(int thisMonth)
        {
            int firstInventory = 8, secondInventory = 2;
            if (thisMonth >= secondInventory & thisMonth < firstInventory) return secondInventory;

            return firstInventory;
        }
    }
}
