using System;
using System.Windows;


namespace TechCardGeneration.Windows
{
    public partial class InitialWindow : Window
    {
        private bool isContinueButtonClicked = false;

        public InitialWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        // События клика кнопок
        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            ShowCreatingElements();
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(StudentsInputTextBox.Text) || string.IsNullOrWhiteSpace(ColumnsInputTextBox.Text))
            {
                MessageBox.Show("Внимательно проверьте, чтобы все поля были заполнены.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            int studentsNumber = int.Parse(StudentsInputTextBox.Text);
            int columnsNumber = int.Parse(ColumnsInputTextBox.Text);

            if (studentsNumber <= 0 || columnsNumber <= 0)
            {
                MessageBox.Show("Количество студентов или колонок не должно равняться 0!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            TextBoxGenerator.Generation(studentsNumber, StudentTextBoxContainer, 430);
            TextBoxGenerator.Generation(columnsNumber, ColumnNameTextBoxContainer, 300);
            TextBoxGenerator.Generation(columnsNumber, ColumnСoefficientTextBoxContainer, 300);

            ShowStudentInfo();
        }

        private void ContinueGeneratingButton_Click(object sender, RoutedEventArgs e)
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

        // Проверка на ввод чисел
        private void StudentsInputTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            NumberInputCheck(e);
        }

        private void ColumnsInputTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            NumberInputCheck(e);
        }

        // Метод для проверки на ввод числа
        private void NumberInputCheck(System.Windows.Input.TextCompositionEventArgs e)
        {
            if (!char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
            }
        }

        // Показ элементов интерфейса
        private void ShowCreatingElements()
        {
            InitialWindowMainSP.Visibility = Visibility.Collapsed;
            InitialWindowControlSP.Visibility = Visibility.Collapsed;

            CreatingLabelsWindowMainSP.Visibility = Visibility.Visible;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Visible;
            CreatingWindowControlSP.Visibility = Visibility.Visible;
        }

        private void ShowStudentInfo()
        {
            CreatingLabelsWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingWindowControlSP.Visibility = Visibility.Collapsed;

            StudentLabelSP.Visibility = Visibility.Visible;
            StudentSP.Visibility = Visibility.Visible;
            GeneratingWindowControlSP.Visibility = Visibility.Visible;
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
            FinishingWindowControlSP.Visibility = Visibility.Visible;
        }
    }
}
