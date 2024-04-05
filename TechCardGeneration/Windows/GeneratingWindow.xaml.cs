using System;
using System.Windows;

namespace TechCardGeneration.Windows
{ 
    public partial class GeneratingWindow : Window
    {
        private bool isContinueButtonClicked = false;

        public GeneratingWindow(int studentsNumber, int columnsNumber)
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            TextBoxGenerator.Generation(studentsNumber, StudentTextBoxContainer, 430);
            TextBoxGenerator.Generation(columnsNumber, ColumnNameTextBoxContainer, 300);
            TextBoxGenerator.Generation(columnsNumber, ColumnСoefficientTextBoxContainer, 300);
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            if (isContinueButtonClicked)
            {
                ShowFileCreatingElements();
                return;
            }

            isContinueButtonClicked = true;
            ShowColumnInfo();
        }
        private void FinishButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PathInputTextBox.Text) || string.IsNullOrWhiteSpace(FileNameInputTextBox.Text))
            {
                MessageBox.Show("Внимательно проверьте, чтобы все поля были заполнены.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        private void ShowColumnInfo() 
        {
            StudentLabelSP.Visibility = Visibility.Collapsed;
            StudentSP.Visibility = Visibility.Collapsed;

            ColumnNameLabelSP.Visibility = Visibility.Visible;
            ColumnNameSP.Visibility = Visibility.Visible;
            ColumnCoefficientLabelSP.Visibility = Visibility.Visible;
            ColumnCoefficientSP.Visibility = Visibility.Visible;
        }

        private void ShowFileCreatingElements() 
        {
            ColumnNameLabelSP.Visibility = Visibility.Collapsed;
            ColumnNameSP.Visibility = Visibility.Collapsed;
            ColumnCoefficientLabelSP.Visibility = Visibility.Collapsed;
            ColumnCoefficientSP.Visibility = Visibility.Collapsed;
            GeneratingWindowControlSP.Visibility = Visibility.Collapsed;

            CreatingLabelsWindowSP.Visibility = Visibility.Visible;
            CreatingTextBoxesWindowSP.Visibility = Visibility.Visible;
            CreatingWindowControlSP.Visibility = Visibility.Visible;
        }
    }
}