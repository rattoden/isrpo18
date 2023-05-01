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

namespace isrpo18
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ExcelPageBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new ExcelPage());
        }

        private void WordPageBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new WordPage());
        }

        private void DeleteDataBtn_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить данные?", "Внимание!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                using (var isrpoEntities = new IsrpoEntities())
                {
                    isrpoEntities.employees.RemoveRange(isrpoEntities.employees.ToList());
                    isrpoEntities.SaveChanges();
                    IsrpoEntities.GetContext().employees.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList().Clear();
                    foreach (var uslugi in isrpoEntities.employees.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList())
                    {
                        IsrpoEntities.GetContext().employees.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList().Add(uslugi);
                    }
                }
            }
        }
    }
}
