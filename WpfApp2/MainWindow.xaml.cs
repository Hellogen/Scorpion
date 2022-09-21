using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            link1.Text = INFOBANK.pathtoexcel1;
            link2.Text = INFOBANK.pathtoexcel2;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //передача файлов с формы в инфобанк
            INFOBANK.yachs1 = ProvYachS1.Text;
            INFOBANK.yachs2 = ProvYachS2.Text;
            INFOBANK.yachdo1 = ProvYachDo1.Text;
            INFOBANK.yachdo2 = ProvYachDo2.Text;
            INFOBANK.stolbets1 = stolbets1.Text;
            INFOBANK.stolbets2 = stolbets2.Text;
            INFOBANK.ZapisVStolbets = PerezapisStolbets1.Text;
            INFOBANK.SchetSStolbca = zapisStolb2.Text;
            INFOBANK.List1 = list1.Text;
            INFOBANK.LIst2 = list2.Text;
            INFOBANK.pathtoexcel1 = link1.Text;
            INFOBANK.pathtoexcel2 = link2.Text;
            INFOBANK.start = true;
            //progressbar.Minimum = ProvYachS1.Text;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            INFOBANK.quit();
            
           
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();

            try
            {

                if (open.ShowDialog() == true)
                {
                    link1.Text = open.FileName;
                    INFOBANK.pathtoexcel1 = open.FileName;
                }
            }
            catch
            {

            }

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();

            try
            {

                if (open.ShowDialog() == true)
                {
                    link2.Text = open.FileName;
                    INFOBANK.pathtoexcel2 = open.FileName;
                }
            }
            catch
            {

            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            INFOBANK.quit();
            Application.Current.Shutdown();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void TextBlock_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void link1_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void link2_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
