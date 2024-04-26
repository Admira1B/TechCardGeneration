using System;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace TechCardGeneration.Windows
{
    public partial class InitialWindow : Window
    {
        private bool isContinueButtonClicked = false;

        private string[] students = null;
        private string[] columnsNames = null;
        private double[] columnsCoefficients = null;

        public InitialWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        // События клика кнопок.
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

            students = new string[studentsNumber];
            columnsNames = new string[columnsNumber];
            columnsCoefficients = new double[columnsNumber];

            ShowStudentInfo();
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            BackToCreatingElements();
            isContinueButtonClicked = false;
            StudentTextBoxContainer.Children.Clear();
            ColumnNameTextBoxContainer.Children.Clear();
            ColumnСoefficientTextBoxContainer.Children.Clear();
        }

        private void ContinueGeneratingButton_Click(object sender, RoutedEventArgs e)
        {

            if (isContinueButtonClicked)
            {
                try
                {
                    AddDatasToArray(ColumnNameTextBoxContainer, columnsNames);
                }
                catch (Exception)
                {
                    MessageBox.Show("Внимательно проверьте, чтобы все поля названий колонок были заполнены.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                try
                {
                    AddDatasToArray(ColumnСoefficientTextBoxContainer, columnsCoefficients);
                }
                catch (Exception)
                {
                    MessageBox.Show("Внимательно проверьте, чтобы все коэффиценты были введены через запятую и не было пустых полей.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                ShowFileCreatingElements();
                return;
            }

            isContinueButtonClicked = true;
            try
            {
                AddDatasToArray(StudentTextBoxContainer, students);
            }
            catch (Exception)
            {
                MessageBox.Show("Внимательно проверьте, чтобы все поля имен студентов были заполнены.", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                isContinueButtonClicked = false;
                return;
            }
            
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

        // Проверка на ввод чисел.
        private void StudentsInputTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            NumberInputCheck(e);
        }

        private void ColumnsInputTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            NumberInputCheck(e);
        }

        // Метод для проверки на ввод числа.
        private void NumberInputCheck(System.Windows.Input.TextCompositionEventArgs e)
        {
            if (!char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
            }
        }

        // Метод для сохранения данных в массивы.
        private void AddDatasToArray(StackPanel textBoxContainer, string[] strArray)
        {
            int count = 0;

            if (textBoxContainer != null)
            {
                foreach (var item in textBoxContainer.Children)
                {
                    if (item is StackPanel)
                    {
                        StackPanel newSP = item as StackPanel;
                        foreach (var itemSP in newSP.Children)
                        {
                            if (itemSP is TextBox)
                            {
                                TextBox textBox = itemSP as TextBox;
                                if (string.IsNullOrWhiteSpace(textBox.Text)) throw new Exception(); 

                                strArray[count] = textBox.Text;
                                count++;
                            }
                        }
                    }
                }
            }
        }
        private void AddDatasToArray(StackPanel textBoxContainer, double[] dblArray)
        {
            int count = 0;

            if (textBoxContainer != null)
            {
                foreach (var item in textBoxContainer.Children)
                {
                    if (item is StackPanel)
                    {
                        StackPanel newSP = item as StackPanel;
                        foreach (var itemSP in newSP.Children)
                        {
                            if (itemSP is TextBox)
                            {
                                TextBox textBox = itemSP as TextBox;
                                dblArray[count] = double.Parse(textBox.Text);
                                count++;
                            }
                        }
                    }
                }
            }
        }

        // Показ элементов интерфейса.
        private void ShowCreatingElements()
        {
            InitialWindowMainSP.Visibility = Visibility.Collapsed;
            InitialWindowControlSP.Visibility = Visibility.Collapsed;

            CreatingLabelsWindowMainSP.Visibility = Visibility.Visible;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Visible;
            CreatingWindowControlSP.Visibility = Visibility.Visible;
        }

        private void BackToCreatingElements() 
        {
            BackSP.Visibility = Visibility.Collapsed;
            StudentLabelSP.Visibility = Visibility.Collapsed;
            StudentSP.Visibility = Visibility.Collapsed;
            GeneratingWindowControlSP.Visibility = Visibility.Collapsed;
            ColumnNameLabelSP.Visibility = Visibility.Collapsed;
            ColumnNameSP.Visibility = Visibility.Collapsed;
            ColumnCoefficientLabelSP.Visibility = Visibility.Collapsed;
            ColumnCoefficientSP.Visibility = Visibility.Collapsed;

            CreatingLabelsWindowMainSP.Visibility = Visibility.Visible;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Visible;
            CreatingWindowControlSP.Visibility = Visibility.Visible;
        }

        private void ShowStudentInfo()
        {
            CreatingLabelsWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingWindowControlSP.Visibility = Visibility.Collapsed;

            BackSP.Visibility= Visibility.Visible;
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
            BackSP.Visibility = Visibility.Collapsed;
            ColumnNameLabelSP.Visibility = Visibility.Collapsed;
            ColumnNameSP.Visibility = Visibility.Collapsed;
            ColumnCoefficientLabelSP.Visibility = Visibility.Collapsed;
            ColumnCoefficientSP.Visibility = Visibility.Collapsed;
            GeneratingWindowControlSP.Visibility = Visibility.Collapsed;

            PathExamples.Visibility = Visibility.Visible;
            CreatingLabelsWindowSP.Visibility = Visibility.Visible;
            CreatingTextBoxesWindowSP.Visibility = Visibility.Visible;
            FinishingWindowControlSP.Visibility = Visibility.Visible;
        }


        // Создание EXEL файла.
        private void CreateExelFile(string fileName, string path, string[] students, string[] columnsName, double[] columnsCoefficients)
        {

        }
    }
}
