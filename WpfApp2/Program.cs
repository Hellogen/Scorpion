using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
/* Пример этойже программы в консоли
namespace ConsoleApp2
{
    internal class Program
    {
        public static void Save(string theText, string path, bool toggle = false)
        {
            //SaveFileDialog Saving = new SaveFileDialog();

            // Saving.DefaultExt = ".txt";
            //  Saving.Filter = "Text documents (.txt)|*.txt";

            using (FileStream file1 = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            {
                StreamWriter writter = new StreamWriter(file1);
                writter.WriteLine(theText);
                writter.Close();
                if (toggle == true)
                {
                    Process.Start(path);
                }
            }




        }
        static void Man()
        {
            string alp = "1234567890";
            XmlDocument xDoc = new XmlDocument();
            string text = "";
            Excel.Application ex1 = new Microsoft.Office.Interop.Excel.Application();
            Excel.Application ex2 = new Microsoft.Office.Interop.Excel.Application();
            string pathtoexcel;
            string pathtoexcel2;
            FileInfo fileInfo = new FileInfo("test222.xlsx"); //динамическая ссылка к файлу excel
            FileInfo fileInfo2 = new FileInfo("test.xlsx");
            if (fileInfo.Exists && fileInfo2.Exists)
            {
                ex1.Visible = true;
                ex1.SheetsInNewWorkbook = 2;

                ex2.Visible = true;
                ex2.SheetsInNewWorkbook = 2;
                
                
                Console.WriteLine("yes"); //подтверждение в чат что файл существует
                pathtoexcel = fileInfo.FullName; // получение полного пути к файлу excel
                pathtoexcel2 = fileInfo2.FullName; //получение полного пути к файлу excel2
                ex1.Workbooks.Open(pathtoexcel,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
                ex2.Workbooks.Open(pathtoexcel2,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
                Excel.Worksheet sheet = (Excel.Worksheet)ex1.Worksheets.get_Item(1);
                Excel.Worksheet sheet2 = (Excel.Worksheet)ex2.Worksheets.get_Item(2);

                
               
                

                for (int i = 6; i < 153; i++ )
                {
                    Excel.Range forYach2 = sheet2.Cells[i, 5] as Excel.Range;
                    string find = "7" + forYach2.Value2.ToString();
                    find = find.Replace("-", "");
                    for(int j = 8; j < 261; j++)
                    {
                         Excel.Range forYach = sheet.Cells[j, 60] as Excel.Range;
                         if (forYach.Value2.ToString() == find)
                        {
                            sheet2.Cells[i, 6] = sheet.Cells[j,63];
                            Console.WriteLine("V" + find + " " + forYach.ToString());
                            Thread.Sleep(10);
                        }
                    }
                }
                
                sheet2.SaveAs("test3.xlsx");
                //Получаем значение из ячейки и преобразуем в строку
                
                
            }
            else
            {
                Console.WriteLine("no");
            }
            //Console.WriteLine(text);
            Console.WriteLine("finish");
            
            Console.ReadKey();
            ex1.Quit();
            
            
        }
    }
}
*/