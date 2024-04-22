using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

internal class Program
{
    private static void Main(string[] args)
    {
        // Для .xlsx файла
        IWorkbook workbook;
        using (FileStream file = new FileStream(@"C:\Users\ATPIX\Desktop\HMF_24_dataset_sch.xlsx", FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(file);
        }

        int count1 = 0, count2 = 0, count3 = 0, count4 = 0, count5 = 0, count6 = 0, count7 = 0, count8 = 0, count9 = 0, count10 = 0, count11 = 0, count12 = 0, count13 = 0, count14 = 0, count15 = 0, count16 = 0;
        ISheet sheet = workbook.GetSheetAt(0);
        IRow row = sheet.GetRow(2);
        ICell cell = row.GetCell(4);
        var value = cell.StringCellValue;

        for(int i=2;i<11859;i++)
        {
            sheet = workbook.GetSheetAt(0);
            row = sheet.GetRow(i);
            cell = row.GetCell(5);
            value = cell.StringCellValue;
            Console.WriteLine(Convert.ToString(value)+"-"+Convert.ToString(i));
            
            if (value == "район Богородское")
                count1++;
            else if (value == "район Вешняки")
                count2++;
            else if (value == "район Восточное Измайлово")
                count3++;
            else if (value == "район Восточный")
                count4++;
            else if (value == "район Гольяново")
                count5++;
            else if (value == "район Ивановское")
                count6++;
            else if (value == "район Измайлово")
                count7++;
            else if (value == "район Косино-Ухтомский")
                count8++;
            else if (value == "район Метрогородок")
                count9++;
            else if (value == "район Новогиреево")
                count10++;
            else if (value == "район Новокосино")
                count11++;
            else if (value == "район Перово")
                count12++;
            else if (value == "район Преображенское")
                count13++;
            else if (value == "район Северное Измайлово")
                count14++;
            else if (value == "район Соколиная Гора")
                count15++;
            else if (value == "район Сокольники")
                count16++;
        }
        for (int i = 11859; i < 20857; i++)
        {
            sheet = workbook.GetSheetAt(0);
            row = sheet.GetRow(i);
            cell = row.GetCell(5);
            value = cell.StringCellValue;
            Console.WriteLine(Convert.ToString(value) + "-" + Convert.ToString(i));

            if (value == "район Богородское")
                count1++;
            else if (value == "район Вешняки")
                count2++;
            else if (value == "район Восточное Измайлово")
                count3++;
            else if (value == "район Восточный")
                count4++;
            else if (value == "район Гольяново")
                count5++;
            else if (value == "район Ивановское")
                count6++;
            else if (value == "район Измайлово")
                count7++;
            else if (value == "район Косино-Ухтомский")
                count8++;
            else if (value == "район Метрогородок")
                count9++;
            else if (value == "район Новогиреево")
                count10++;
            else if (value == "район Новокосино")
                count11++;
            else if (value == "район Перово")
                count12++;
            else if (value == "район Преображенское")
                count13++;
            else if (value == "район Северное Измайлово")
                count14++;
            else if (value == "район Соколиная Гора")
                count15++;
            else if (value == "район Сокольники")
                count16++;
        }

        Console.WriteLine(Convert.ToString(count1)," ",Convert.ToString(count2), " ", count3, " ", count4, " ", count5, " ", count6, " ", count7, " ", count8, " ", count9, " ", count10, " ", count11, " ", count12, " ", count13, " ", count14, " ", count15, " ", count16);
        Console.WriteLine(Convert.ToString(count2));
        Console.WriteLine(Convert.ToString(count3));
        Console.WriteLine(Convert.ToString(count4));
        Console.WriteLine(Convert.ToString(count5));
        Console.WriteLine(Convert.ToString(count6));

        Console.ReadLine();
    }
}