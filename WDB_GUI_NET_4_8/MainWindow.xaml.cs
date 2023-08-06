using System;
using System.Collections.Generic;
using System.Data;
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

        /// <summary>
        /// Договоры кандидатов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            WordDocumentBuilder.ElectionContracts.Builder builder = new WordDocumentBuilder.ElectionContracts.Builder();
            var dt = builder.BuildContractsCandidates("1");
            // Тестово посмотреть
            DataGrid.ItemsSource = dt.DefaultView;
            MessageBox.Show("Готово.");
        }

        /// <summary>
        /// Договоры партий
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            WordDocumentBuilder.ElectionContracts.Builder builder = new WordDocumentBuilder.ElectionContracts.Builder();
            var dt = builder.BuildContractsParties("1");
            // Тестово посмотреть
            DataGrid.ItemsSource = dt.DefaultView;
            MessageBox.Show("Готово.");
        }

        /// <summary>
        /// Протоколы партий
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            WordDocumentBuilder.ElectionContracts.Builder builder = new WordDocumentBuilder.ElectionContracts.Builder();
            DataTable dt = new DataTable();
            //// test
            //dt = ExcelProcessor.ReadExcelSheet(Settings.Default.Parties_TalonsFilePath_Вести_ФМ, sheetNumber: 0);
            //DataGrid.ItemsSource = dt.DefaultView;
            try
            {
                dt = builder.BuildProtocolsParties("1");
                // Тестово посмотреть
                DataGrid.ItemsSource = dt.DefaultView;
                MessageBox.Show("Готово.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            WordDocumentBuilder.ElectionContracts.Builder builder = new WordDocumentBuilder.ElectionContracts.Builder();
            DataTable dt = new DataTable();
            try
            {
                dt = builder.BuildProtocolsCandidates("1");
                // Тестово посмотреть
                DataGrid.ItemsSource = dt.DefaultView;
                MessageBox.Show("Готово.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadSettings(object sender, RoutedEventArgs e)
        {
            //
            tbContractsFolderPath.Text = Settings.Default.ContractsFolderPath;
            // Кандидаты
            tbCandidates_FilePath.Text = Settings.Default.Candidates_FilePath;
            tbCandidates_TemplateFilePath_РВ.Text = Settings.Default.Candidates_TemplateFilePath_РВ;
            tbCandidates_TemplateFilePath_ТВ.Text = Settings.Default.Candidates_TemplateFilePath_ТВ;
            //
            tbCandidates_TalonsDefaultFilePath.Text = Settings.Default.Candidates_TalonsFilePath;
            tbCandidates_Talons_Маяк.Text = Settings.Default.Candidates_TalonsFilePath_Маяк;
            tbCandidates_Talons_Вести_ФМ.Text = Settings.Default.Candidates_TalonsFilePath_Вести_ФМ;
            tbCandidates_Talons_Радио_России.Text = Settings.Default.Candidates_TalonsFilePath_Радио_России;
            tbCandidates_Talons_Россия_1.Text = Settings.Default.Candidates_TalonsFilePath_Россия_1;
            tbCandidates_Talons_Россия24.Text = Settings.Default.Candidates_TalonsFilePath_Россия_24;
            // Партии
            tbParties_FilePath.Text = Settings.Default.Parties_FilePath;
            tbParties_TemplateFilePath_РВ.Text = Settings.Default.Parties_TemplateFilePath_РВ;
            tbParties_TemplateFilePath_ТВ.Text = Settings.Default.Parties_TemplateFilePath_ТВ;
            //
            tbParties_TalonsDefaultFilePath.Text = Settings.Default.Parties_TalonsFilePath;
            tbParties_Talons_Маяк.Text = Settings.Default.Parties_TalonsFilePath_Маяк;
            tbParties_Talons_Вести_ФМ.Text = Settings.Default.Parties_TalonsFilePath_Вести_ФМ;
            tbParties_Talons_Радио_России.Text = Settings.Default.Parties_TalonsFilePath_Радио_России;
            tbParties_Talons_Россия_1.Text = Settings.Default.Parties_TalonsFilePath_Россия_1;
            tbParties_Talons_Россия24.Text = Settings.Default.Parties_TalonsFilePath_Россия_24;
            // Протоколы
            tbProtocols_FolderPath.Text = Settings.Default.Protocols_FolderPath;
            tbProtocols_TemplateFilePath_Candidates.Text = Settings.Default.Protocols_TemplateFilePath_Candidates;
            tbProtocols_TemplateFilePath_Parties.Text = Settings.Default.Protocols_TemplateFilePath_Parties;
            tbProtocols_FilePath.Text = Settings.Default.Protocols_FilePath;
        }

        private void SaveSettings(object sender, RoutedEventArgs e)
        {
            //
            if (tbContractsFolderPath.Text != "") Settings.Default.ContractsFolderPath = tbContractsFolderPath.Text;
            // Кандидаты
            if (tbCandidates_FilePath.Text != "") Settings.Default.Candidates_FilePath = tbCandidates_FilePath.Text;
            if (tbCandidates_TemplateFilePath_РВ.Text != "") Settings.Default.Candidates_TemplateFilePath_РВ = tbCandidates_TemplateFilePath_РВ.Text;
            if (tbCandidates_TemplateFilePath_ТВ.Text != "") Settings.Default.Candidates_TemplateFilePath_ТВ = tbCandidates_TemplateFilePath_ТВ.Text;
            //
            if (tbCandidates_TalonsDefaultFilePath.Text != "") Settings.Default.Candidates_TalonsFilePath = tbCandidates_TalonsDefaultFilePath.Text;
            if (tbCandidates_Talons_Маяк.Text != "") Settings.Default.Candidates_TalonsFilePath_Маяк = tbCandidates_Talons_Маяк.Text;
            if (tbCandidates_Talons_Вести_ФМ.Text != "") Settings.Default.Candidates_TalonsFilePath_Вести_ФМ = tbCandidates_Talons_Вести_ФМ.Text;
            if (tbCandidates_Talons_Радио_России.Text != "") Settings.Default.Candidates_TalonsFilePath_Радио_России = tbCandidates_Talons_Радио_России.Text;
            if (tbCandidates_Talons_Россия_1.Text != "") Settings.Default.Candidates_TalonsFilePath_Россия_1 = tbCandidates_Talons_Россия_1.Text;
            if (tbCandidates_Talons_Россия24.Text != "") Settings.Default.Candidates_TalonsFilePath_Россия_24 = tbCandidates_Talons_Россия24.Text;
            // Партии
            if (tbParties_FilePath.Text != "") Settings.Default.Parties_FilePath = tbParties_FilePath.Text;
            if (tbParties_TemplateFilePath_РВ.Text != "") Settings.Default.Parties_TemplateFilePath_РВ = tbParties_TemplateFilePath_РВ.Text;
            if (tbParties_TemplateFilePath_ТВ.Text != "") Settings.Default.Parties_TemplateFilePath_ТВ = tbParties_TemplateFilePath_ТВ.Text;
            //
            if (tbParties_TalonsDefaultFilePath.Text != "") Settings.Default.Parties_TalonsFilePath = tbParties_TalonsDefaultFilePath.Text;
            if (tbParties_Talons_Маяк.Text != "") Settings.Default.Parties_TalonsFilePath_Маяк = tbParties_Talons_Маяк.Text;
            if (tbParties_Talons_Вести_ФМ.Text != "") Settings.Default.Parties_TalonsFilePath_Вести_ФМ = tbParties_Talons_Вести_ФМ.Text;
            if (tbParties_Talons_Радио_России.Text != "") Settings.Default.Parties_TalonsFilePath_Радио_России = tbParties_Talons_Радио_России.Text;
            if (tbParties_Talons_Россия_1.Text != "") Settings.Default.Parties_TalonsFilePath_Россия_1 = tbParties_Talons_Россия_1.Text;
            if (tbParties_Talons_Россия24.Text != "") Settings.Default.Parties_TalonsFilePath_Россия_24 = tbParties_Talons_Россия24.Text;
            // Протоколы
            if (tbProtocols_FolderPath.Text != "") Settings.Default.Protocols_FolderPath = tbProtocols_FolderPath.Text;
            if (tbProtocols_TemplateFilePath_Candidates.Text != "") Settings.Default.Protocols_TemplateFilePath_Candidates = tbProtocols_TemplateFilePath_Candidates.Text;
            if (tbProtocols_TemplateFilePath_Parties.Text != "") Settings.Default.Protocols_TemplateFilePath_Parties = tbProtocols_TemplateFilePath_Parties.Text;
            if (tbProtocols_FilePath.Text != "") Settings.Default.Protocols_FilePath = tbProtocols_FilePath.Text;
            //
            Settings.Default.Save();
        }

        
    }
}
