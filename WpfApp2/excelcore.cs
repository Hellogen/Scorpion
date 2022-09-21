using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Microsoft.Win32;

namespace WpfApp2
{

    public class excelcoreobject
    {
        

        public void start()
        {
            try
            {
                string savepath = "";
                bool next = false;
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                
                try //открывается выбор куда сохранять файл
                {
                    if(saveFileDialog.ShowDialog() == true)
                    {
                        savepath = saveFileDialog.FileName;
                        next = true;
                    }    
                    else
                    { }
                }
                catch
                {

                }
                int list1 = Convert.ToInt32(INFOBANK.List1); //конвертация кол-во листов первого документа
                int list2 = Convert.ToInt32(INFOBANK.LIst2); //конвертация кол-во листов второго документа
                int yachs1 = Convert.ToInt32(INFOBANK.yachs1); //конвертация начало сканирования с ячейки файла 1
                int yachs2 = Convert.ToInt32(INFOBANK.yachs2); //конвертация начало сканирования с ячейки файла 2
                int yachdo1 = Convert.ToInt32(INFOBANK.yachdo1); //Конвертация конец сканирования до ячейки файла 1
                int yachdo2 = Convert.ToInt32(INFOBANK.yachdo2); //конвертация конец сканирования до ячейки файла 2
                int stolbets1 = Convert.ToInt32(INFOBANK.stolbets1); //Конвертация столбец файла 1
                int stolbets2 = Convert.ToInt32(INFOBANK.stolbets2); //Конвертация столбец файла 2
                int schetsstolbca = Convert.ToInt32(INFOBANK.SchetSStolbca); //Конвертация считывание со столбца файла 2
                int sapisvstolbets = Convert.ToInt32(INFOBANK.ZapisVStolbets); //Конвертация считывание со столбца файла 1

                INFOBANK.minimumProgress = Convert.ToDouble(yachs1); //Для прогресс бара Минимум (конвертация)
                INFOBANK.maximumProgress = Convert.ToDouble(yachdo1-1); //Для прогресс бара максимум (конвертация) - 1 т.к цикл идет до числа (если не будет минус еденицы то прогресс бар не дойдет до конца)

                
                string pathtoexcel;
                string pathtoexcel2;
                FileInfo fileInfo = new FileInfo(INFOBANK.pathtoexcel2); //динамическая ссылка к файлу excel 2
                FileInfo fileInfo2 = new FileInfo(INFOBANK.pathtoexcel1); //Динамическая ссылка к файлу excel 1
                FileInfo fileinfo3 = new FileInfo(savepath);
                if (fileinfo3.Exists)
                {
                    fileinfo3.Delete();
                }
                if (fileInfo.Exists && fileInfo2.Exists && next) 
                {
                    INFOBANK.ex1 = new Excel.Application(); // Инициализация Ексель файла
                    INFOBANK.ex2 = new Excel.Application(); // Инициализация Ексель файла
                    //INFOBANK.isrunning = true;
                    Console.WriteLine("yes"); //подтверждение в чат что файл существует
                    pathtoexcel = fileInfo.FullName; // получение полного пути к файлу excel
                    pathtoexcel2 = fileInfo2.FullName; //получение полного пути к файлу excel2



                    INFOBANK.ex1.Workbooks.Open(pathtoexcel,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing); // открытие документа по пути


                    INFOBANK.ex2.Workbooks.Open(pathtoexcel2,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing); // открытие второго документа по пути2
                    INFOBANK.sheet = (Excel.Worksheet)INFOBANK.ex1.Worksheets.get_Item(list2);
                    INFOBANK.sheet2 = (Excel.Worksheet)INFOBANK.ex2.Worksheets.get_Item(list1);
                    Console.WriteLine("yesend"); //подтверждение в чат что файл существует

                    

                    for (int i = yachs1; i <= yachdo1; i++) //6-152
                    {
                        Excel.Range forYach2 = INFOBANK.sheet2.Cells[i, stolbets1] as Excel.Range;
                        string find2 = forYach2.Value2.ToString();
                        //find = find.Replace("-", ""); //фильтры для замены для поиска здесь
                        INFOBANK.progress =  Convert.ToDouble(i);
                        for (int j = yachs2; j <= yachdo2; j++) //j 8-260
                        {
                            Excel.Range forYach = INFOBANK.sheet.Cells[j, stolbets2] as Excel.Range;
                            string find = forYach.Value2.ToString();
                            if (find.Contains(find2)) // если первый документ содержит запись из второго то начинается запись
                            {
                                INFOBANK.sheet2.Cells[i, sapisvstolbets] = INFOBANK.sheet.Cells[j, schetsstolbca];
                                Console.WriteLine("V" + find + " " + forYach.ToString());
                               
                                //Thread.Sleep(10);
                            }
                        }
                    }

                            INFOBANK.ex2.DisplayAlerts = false; // убирает окно замены файла сохраниния 
                    
                            INFOBANK.ex2.ActiveWorkbook.SaveAs(savepath);
                            //INFOBANK.sheet2.SaveAs(savepath); // сохранение листа
                            
                }
                else
                {
                    Console.WriteLine("no"); // документа не существует ИЛИ нет файла сохранения
                    INFOBANK.error = "Ошибка: Документа/документов не существует";
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("hello this is error:" + ex.Message); // ошибка в чат(если пользуетесь visual studio то смотреть данные отладки
                
            }
            finally
            {INFOBANK.worksbutton = true; /*кнопка запуска активна*/ INFOBANK.quit(); /*quit() Обязательно должен быть последним т.к он абортит поток*/}
            
            
        }
    }

    public static class ExcelCoreReact
    {
        public static excelcoreobject excel;
    }
}



