using System;
using System.Windows;


namespace TechCardGeneration.Windows
{
    public partial class InitialWindow : Window
    {
        public InitialWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

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

            GeneratingWindow generatingWindow = new GeneratingWindow(studentsNumber, columnsNumber);
            generatingWindow.Show();
            this.Close();
        }

        private void StudentsInputTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            NumberInputCheck(e);
        }

        private void ColumnsInputTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            NumberInputCheck(e);
        }

        private void ShowCreatingElements()
        {
            InitialWindowMainSP.Visibility = Visibility.Collapsed;
            InitialWindowControlSP.Visibility = Visibility.Collapsed;

            CreatingLabelsWindowMainSP.Visibility = Visibility.Visible;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Visible;
            CreatingWindowControlSP.Visibility = Visibility.Visible;
        }

        private void NumberInputCheck(System.Windows.Input.TextCompositionEventArgs e) 
        {
            if (!char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
            }
        }
    }
}
