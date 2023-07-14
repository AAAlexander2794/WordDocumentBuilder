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
            var dt = builder.BuildContractsCandidates();
            // Тестово посмотреть
            DataGrid.ItemsSource = dt.DefaultView;
            MessageBox.Show("Готово.");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            WordDocumentBuilder.ElectionContracts.Builder builder = new WordDocumentBuilder.ElectionContracts.Builder();
            var dt = builder.BuildContractsCandidates("1");
            // Тестово посмотреть
            DataGrid.ItemsSource = dt.DefaultView;
            MessageBox.Show("Готово.");
        }

        private void LoadSettings(object sender, RoutedEventArgs e)
        {
            tbCandidatesFilePath.Text = Settings.Default.CandidatesFilePath;
            tbContractsFolderPath.Text = Settings.Default.ContractsFolderPath;
            tbTalonsDefaultFilePath.Text = Settings.Default.TalonsFilePath;
            tbTemplateFilePath_РВ.Text = Settings.Default.TemplateFilePath_РВ;
            tbTemplateFilePath_ТВ.Text = Settings.Default.TemplateFilePath_ТВ;
            //
            tbTalons_Маяк.Text = Settings.Default.TalonsFilePath_Маяк;
            tbTalons_Вести_ФМ.Text = Settings.Default.TalonsFilePath_Вести_ФМ;
            tbTalons_Радио_России.Text = Settings.Default.TalonsFilePath_Радио_России;
            tbTalons_Россия_1.Text = Settings.Default.TalonsFilePath_Россия_1;
            tbTalons_Россия24.Text = Settings.Default.TalonsFilePath_Россия_24;
        }

        private void SaveSettings(object sender, RoutedEventArgs e)
        {
            if (tbCandidatesFilePath.Text != "") Settings.Default.CandidatesFilePath = tbCandidatesFilePath.Text;
            if (tbContractsFolderPath.Text != "") Settings.Default.ContractsFolderPath = tbContractsFolderPath.Text;
            if (tbTalonsDefaultFilePath.Text != "") Settings.Default.TalonsFilePath = tbTalonsDefaultFilePath.Text;
            if (tbTemplateFilePath_РВ.Text != "") Settings.Default.TemplateFilePath_РВ = tbTemplateFilePath_РВ.Text;
            if (tbTemplateFilePath_ТВ.Text != "") Settings.Default.TemplateFilePath_ТВ = tbTemplateFilePath_ТВ.Text;
            //
            if (tbTalons_Маяк.Text != "") Settings.Default.TalonsFilePath_Маяк = tbTalons_Маяк.Text;
            if (tbTalons_Вести_ФМ.Text != "") Settings.Default.TalonsFilePath_Вести_ФМ = tbTalons_Вести_ФМ.Text;
            if (tbTalons_Радио_России.Text != "") Settings.Default.TalonsFilePath_Радио_России = tbTalons_Радио_России.Text;
            if (tbTalons_Россия_1.Text != "") Settings.Default.TalonsFilePath_Россия_1 = tbTalons_Россия_1.Text;
            if (tbTalons_Россия24.Text != "") Settings.Default.TalonsFilePath_Россия_24 = tbTalons_Россия24.Text;

            Settings.Default.Save();
        }
    }
}
