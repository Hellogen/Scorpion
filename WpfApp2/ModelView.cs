using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp2
{
    internal class ModelView : INotifyPropertyChanged //mvvm C#
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
        private double progress = 0; // Прогрессия для прогрессбара
        private bool workingButton = true; //Включение/выключение кнопки
        private bool isRunning = false;
        private double minimum = 0; //Минимум для прогрессбара
        private double maximum = 0; //Максимум для прогрессбара
        private string error = ""; // вывод ошибки
        public string Error
        {
            get { return error; }
            set { error = value; OnPropertyChanged("ERROR"); }
        }
        public double Minimum // вывод минимума в прогресс бар на форму
        {
            get { return minimum; }
            set { minimum = value; OnPropertyChanged("minimum"); }

        }
        public double Maximum // вывод максимума в прогресс бар на форму
        {
            get { return maximum; }
            set { maximum = value; OnPropertyChanged("maximum"); }
        }
        public bool WorkingButton // вывод на форму информации о кнопке включена она или нет 
        {
            get { return workingButton; }
            set { workingButton = value; OnPropertyChanged("workingButton"); }
        }
        public double Progress // вывод на форму информации о прогрессии для прогресс бара
        {
            get { return progress; }
            set { progress = value; OnPropertyChanged("progress"); }
        }
        
        public void InitExcelCore() // инициализация класса
        {
            try
            {
                ExcelCoreReact.excel = new excelcoreobject();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public bool IsRunning
        { get { return isRunning; } set { isRunning = value; OnPropertyChanged("running"); } }
        public ModelView()
        {

            Task.Factory.StartNew(() => // поток обновления и включения функций
            {
                InitExcelCore();
                while (true)
                {
                    if (INFOBANK.start) // если true то запустить поток перебора ексел и отключить кнопки
                    {
                        INFOBANK.error = "";
                        WorkingButton = false;
                        INFOBANK.worksbutton = false;
                        INFOBANK.start = false;
                        
                        INFOBANK.thread = new Thread(new ThreadStart(ExcelCoreReact.excel.start));
                        INFOBANK.thread.Start();
                        
                    }
                    // обновление данных каждую секунду
                    Minimum = INFOBANK.minimumProgress;
                    Maximum = INFOBANK.maximumProgress;
                    Progress = INFOBANK.progress;
                    WorkingButton = INFOBANK.worksbutton;
                    Error = INFOBANK.error;
                    Task.Delay(1000).Wait();
                }

             });
        }
    }
}
