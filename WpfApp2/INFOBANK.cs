using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WpfApp2
{
    public static class INFOBANK
    {
        //переменные которые доступны во всем приложении
        public static string pathtoexcel1 = "1.xlsx"; //путь до 1 ексел файла
        public static string pathtoexcel2 = "2.xlsx"; //путь до 2 ексел файла
        public static string yachdo1 = ""; // номер ячейки докуда начинать перебор в файле екселя 1
        public static string yachdo2 = ""; // номер ячейки докуда начинать перебор в файле екселя 2
        public static string yachs1 = ""; // номер ячейки с которой начинать перебор ексел1
        public static string yachs2 = ""; // номер ячейки с которой начинать перебор ексел2
        public static string stolbets1 = ""; // столбец 1 из файла ексел1
        public static string stolbets2 = ""; // столбец 2 из файла ексел2
        public static string SchetSStolbca = ""; //считывание столбца из файла ексел 2 в ексел1
        public static string ZapisVStolbets = ""; //запись в столбец из ексел 2 в ексел 1
        public static string List1 = ""; //лист файла ексел 1
        public static string LIst2 = ""; //лист файла ексел 2
        public static bool start = false; //запуск функции при тру
        public static bool worksbutton = true; // кнопка работает если тру иначе фолз
        public static Thread thread; //поток в котором запускается функция
        public static double minimumProgress = 0; //прогресс бар минимум
        public static double maximumProgress = 1; //прогресс бра максимум
        public static double progress = 0; // значение ползунка в прогрез баре между минимумом и максимумом
        public static string error = "";
        public static Microsoft.Office.Interop.Excel.Application ex1; //переменная excel 1
        public static Microsoft.Office.Interop.Excel.Application ex2; //переменная excel 2

        public static Microsoft.Office.Interop.Excel.Worksheet sheet; //листы екселя
        public static Microsoft.Office.Interop.Excel.Worksheet sheet2;
        public static void quit()
        {

            try
            {
                //заккоментированное отключено и работает у других у меня лично работает без этого
                //sheet = null;
                if (ex1 != null)
                {
                    
                    //ex1.ActiveWorkbook.Close(false);
                    ex1.Quit();
                }
                //Marshal.ReleaseComObject(ex1);
            }
            catch
            { }
            try
            {
                //Marshal.ReleaseComObject(sheet2);
                //sheet2 = null;
                if (ex2 != null)
                ex2.Quit();
                //Marshal.ReleaseComObject(ex2);
            }
            catch
            { }
            
           
            
            GC.Collect();
            GC.WaitForFullGCComplete();
            GC.Collect();
            if (thread != null)
            thread.Abort();
        }
    }
}
