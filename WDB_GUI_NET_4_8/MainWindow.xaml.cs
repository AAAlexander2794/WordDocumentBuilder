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
using WordDocumentBuilder;

namespace WDB_GUI_NET_4_8
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            WordDocumentBuilder.ElectionContracts.Builder builder = new WordDocumentBuilder.ElectionContracts.Builder();
            var dt = builder.Do();
            // Тестово посмотреть
            DataGrid.ItemsSource = dt.DefaultView;
            MessageBox.Show("Готово.");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            WordDocumentBuilder.ElectionContracts.Builder builder = new WordDocumentBuilder.ElectionContracts.Builder();
            var dt = builder.Do("2");
            // Тестово посмотреть
            DataGrid.ItemsSource = dt.DefaultView;
            MessageBox.Show("Готово.");
        }

        private void LoadSettings(object sender, RoutedEventArgs e)
        {
            tbContractsFolderPath.Text = Settings.Default.ContractsFolderPath;
            tbDataFilePath.Text = Settings.Default.DataFilePath;
            tbTemplateFilePath.Text = Settings.Default.TemplateFilePath;
        }

        private void SaveSettings(object sender, RoutedEventArgs e)
        {
            if (tbContractsFolderPath.Text != "") Settings.Default.ContractsFolderPath = tbContractsFolderPath.Text;
            if (tbDataFilePath.Text != "") Settings.Default.DataFilePath = tbDataFilePath.Text;
            if (tbTemplateFilePath.Text != "") Settings.Default.TemplateFilePath = tbTemplateFilePath.Text;
            Settings.Default.Save();
        }
    }
}
