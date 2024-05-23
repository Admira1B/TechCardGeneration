using System;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

            TextBoxGenerator.Generation(columnsNumber, ColumnNameTextBoxContainer, 300);
            TextBoxGenerator.Generation(columnsNumber, ColumnСoefficientTextBoxContainer, 300);
            students = new string[studentsNumber];
            columnsNames = new string[columnsNumber];
            columnsCoefficients = new double[columnsNumber];

            if (!(bool)NotFillStudentsNames.IsChecked)
            {
                TextBoxGenerator.Generation(studentsNumber, StudentTextBoxContainer, 430);

                for (int i = 0; i < students.Length; i++)
                {
                    students[i] = "";
                }

                ShowStudentInfo();
                return;
            }

            isContinueButtonClicked = true;
            ShowColumnInfo();
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

                    double sumOfCoefficients = 0;

                    for (int i = 0; i < columnsCoefficients.Length; i++)
                    {
                        sumOfCoefficients += (columnsCoefficients[i] * 100);
                    }

                    if (sumOfCoefficients != 100)
                    {
                        MessageBox.Show($"Внимательно проверьте, чтоб сумма коэффицентов была равна 1. Текущая сумма коэффицентов {sumOfCoefficients / 100}", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

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

        private void DirButton_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel файлы (.xlsx)|*.xlsx";
            saveFileDialog.Title = "Сохранить эксель файл";

            if (saveFileDialog.ShowDialog() == true)
            {
                string fullFilePath = saveFileDialog.FileName;
                string fullFileName = saveFileDialog.SafeFileName;
                string filePath = fullFilePath.Replace(fullFileName, "");
                string fileName = fullFileName.Substring(0, fullFileName.Length - 5);

                PathInputTextBox.Text = filePath;
                FileNameInputTextBox.Text = fileName;
            }
        }

        private void FinishButton_Click(object sender, RoutedEventArgs e)
        {
            InitialWindow newWindow = new InitialWindow();

            if (string.IsNullOrWhiteSpace(PathInputTextBox.Text) || string.IsNullOrWhiteSpace(FileNameInputTextBox.Text))
            {
                MessageBox.Show("Внимательно проверьте, чтобы все поля были заполнены.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string fileName = FileNameInputTextBox.Text;
            string directoryPath = PathInputTextBox.Text;

            if (Directory.Exists(directoryPath))
            {
                if (File.Exists(directoryPath + fileName + ".xlsx"))
                {
                    MessageBoxResult result = MessageBox.Show("Файл с таким названием уже существует, хотите ли вы его заменить?", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (result == MessageBoxResult.Yes)
                    {
                        try
                        {
                            FileCreate(fileName, directoryPath);

                            newWindow.Show();
                            this.Close();
                            return;
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Файл с таким названием открыт на вашем компьютере, закройте его и повторите попытку.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }

                    MessageBox.Show("Введите новое название для файла или измение директорию.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                FileCreate(fileName, directoryPath);

                newWindow.Show();
                this.Close();
                return;
            }

            MessageBox.Show("Директория по указанному пути не существует.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
            CheckBoxContainerStudentsNamesSP.Visibility = Visibility.Visible;
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
            CheckBoxContainerStudentsNamesSP.Visibility = Visibility.Visible;
        }

        private void ShowStudentInfo()
        {
            CreatingLabelsWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingWindowControlSP.Visibility = Visibility.Collapsed;
            CheckBoxContainerStudentsNamesSP.Visibility = Visibility.Collapsed;

            BackSP.Visibility = Visibility.Visible;
            StudentLabelSP.Visibility = Visibility.Visible;
            StudentSP.Visibility = Visibility.Visible;
            GeneratingWindowControlSP.Visibility = Visibility.Visible;
        }

        private void ShowColumnInfo()
        {
            StudentLabelSP.Visibility = Visibility.Collapsed;
            StudentSP.Visibility = Visibility.Collapsed;
            CreatingLabelsWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingTextBoxesWindowMainSP.Visibility = Visibility.Collapsed;
            CreatingWindowControlSP.Visibility = Visibility.Collapsed;
            CheckBoxContainerStudentsNamesSP.Visibility = Visibility.Collapsed;

            BackSP.Visibility = Visibility.Visible;
            ColumnNameLabelSP.Visibility = Visibility.Visible;
            ColumnNameSP.Visibility = Visibility.Visible;
            ColumnCoefficientLabelSP.Visibility = Visibility.Visible;
            ColumnCoefficientSP.Visibility = Visibility.Visible;
            GeneratingWindowControlSP.Visibility = Visibility.Visible;
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
            DirButton.Visibility = Visibility.Visible;
            FinishingWindowControlSP.Visibility = Visibility.Visible;
        }

        // Создание файла.
        private void FileCreate(string fileName, string directoryPath)
        {
            CreateExelFile(fileName, directoryPath, students, columnsNames, columnsCoefficients);

            if (File.Exists(directoryPath + fileName + ".xlsx"))
            {
                MessageBox.Show("Файл был успешно создан.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            MessageBox.Show("Попробуйте ввести нвовый путь или перезапустить программу.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void CreateExelFile(string fileName, string path, string[] students, string[] columnsNames, double[] columnsCoefficients)
        {
            int startRow = 2;
            int startColumn = 4;
            int totalColumn = startColumn + columnsNames.Length;
            int examColumn = totalColumn + 1;
            int totalExamCol = examColumn + 1;
            int coefficientStartRow = 3;
            int studentsStartRow = 5;
            int studentStartCol = 2;
            int studentColToMerge = 3;

            using (var package = new ExcelPackage())
            {

                // Генерация листов.
                var titlePage = package.Workbook.Worksheets.Add("Титульник");
                var controlpoint = package.Workbook.Worksheets.Add("Контрольные точки");

                // Генерация полей титульного листа.
                titlePage.Cells["A1"].Value = "Технологическая карта";
                titlePage.Cells["A1:E1"].Merge = true;

                titlePage.Cells["A2"].Value = "Дисциплина \"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\"";
                titlePage.Cells["A2"].Style.Fill.SetBackground(Color.FromArgb(252, 184, 184));
                titlePage.Cells["A2:E2"].Merge = true;

                titlePage.Cells["A4"].Value = "Специальность: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!";
                titlePage.Cells["A4"].Style.Fill.SetBackground(Color.FromArgb(252, 184, 184));
                titlePage.Cells["A4:E4"].Merge = true;

                titlePage.Cells["A5"].Value = "Группа";
                titlePage.Cells["B5"].Value = "!!!!";
                titlePage.Cells["B5"].Style.Fill.SetBackground(Color.FromArgb(252, 184, 184));
                titlePage.Cells["A6"].Value = "Семестр";
                titlePage.Cells["B6"].Value = "!!!!";
                titlePage.Cells["B6"].Style.Fill.SetBackground(Color.FromArgb(252, 184, 184));

                titlePage.Cells["A8"].Value = "Виды учебной деятельности";
                titlePage.Cells["A8"].Style.Fill.SetBackground(Color.FromArgb(234, 255, 231));
                titlePage.Cells["A8:C8"].Merge = true;
                titlePage.Cells["D8"].Value = "Весовой коэффициент";
                titlePage.Cells["D8"].Style.Fill.SetBackground(Color.FromArgb(234, 255, 231));
                titlePage.Cells["D8:E8"].Merge = true;

                int titlePageStartRow = 9;
                int titlePageStartCol = 1;
                int titlePageCoefCol = 4;
                for (int i = 0; i < columnsNames.Length; i++)
                {
                    // Заполнение видов учебной деятельности.
                    titlePage.Cells[titlePageStartRow + i, titlePageStartCol].Value = columnsNames[i];
                    titlePage.Cells[$"A{titlePageStartRow + i}"].Style.Fill.SetBackground(Color.FromArgb(245, 211, 181));
                    titlePage.Cells[$"A{titlePageStartRow + i}:C{titlePageStartRow + i}"].Merge = true;

                    // Заполнение весовых коэффициентов
                    titlePage.Cells[titlePageStartRow + i, titlePageCoefCol].Value = columnsCoefficients[i];
                    titlePage.Cells[$"D{titlePageStartRow + i}"].Style.Fill.SetBackground(Color.FromArgb(248, 250, 202));
                    titlePage.Cells[$"D{titlePageStartRow + i}:E{titlePageStartRow + i}"].Merge = true;
                }

                titlePage.Cells[$"A{titlePageStartRow + columnsNames.Length}"].Value = "Итого";
                titlePage.Cells[$"A{titlePageStartRow + columnsNames.Length}"].Style.Fill.SetBackground(Color.FromArgb(252, 183, 119));
                titlePage.Cells[$"D{titlePageStartRow + columnsNames.Length}"].Formula = $"SUM(D{titlePageStartRow}:D{titlePageStartRow + columnsNames.Length - 1})";
                titlePage.Cells[$"D{titlePageStartRow + columnsNames.Length}"].Style.Fill.SetBackground(Color.FromArgb(245, 245, 113));
                titlePage.Cells[$"A{titlePageStartRow + columnsNames.Length}:C{titlePageStartRow + columnsNames.Length}"].Merge = true;
                titlePage.Cells[$"D{titlePageStartRow + columnsNames.Length}:E{titlePageStartRow + columnsNames.Length}"].Merge = true;
                titlePage.Cells[$"F{titlePageStartRow + columnsNames.Length}"].Value = "← Должно быть равно 1!";

                // Генерация обязательных полей страницы контрольной точки.
                controlpoint.Cells["B3:C3"].Merge = true;
                controlpoint.Cells["B4:C4"].Merge = true;

                controlpoint.Cells["B3"].Value = "Вес(значимость)";
                controlpoint.Cells["B3"].Style.Fill.SetBackground(Color.FromArgb(202, 206, 250));
                controlpoint.Cells["B4"].Value = "Дата";
                controlpoint.Cells["B4"].Style.Fill.SetBackground(Color.FromArgb(231, 202, 252));

                controlpoint.Cells[startRow, totalColumn].Value = "Итого";
                controlpoint.Cells[startRow, totalColumn].Style.Fill.SetBackground(Color.FromArgb(245, 211, 181));
                controlpoint.Cells[startRow + 2, totalColumn].Style.Fill.SetBackground(Color.FromArgb(245, 211, 181));
                controlpoint.Cells[startRow, totalColumn, startRow + 1, totalColumn].Merge = true;

                controlpoint.Cells[startRow, examColumn].Value = "Экзамен";
                controlpoint.Cells[startRow + 1, examColumn].Value = 0.2;
                controlpoint.Cells[startRow, examColumn].Style.Fill.SetBackground(Color.FromArgb(234, 255, 231));
                controlpoint.Cells[startRow + 1, examColumn].Style.Fill.SetBackground(Color.FromArgb(234, 255, 231));
                controlpoint.Cells[startRow + 2, examColumn].Style.Fill.SetBackground(Color.FromArgb(234, 255, 231));

                controlpoint.Cells[startRow, totalExamCol].Value = "Итого с экз.";
                controlpoint.Cells[startRow, totalExamCol].Style.Fill.SetBackground(Color.FromArgb(252, 184, 184));
                controlpoint.Cells[startRow + 2, totalExamCol].Style.Fill.SetBackground(Color.FromArgb(252, 184, 184));
                controlpoint.Cells[startRow, totalExamCol, startRow + 1, totalExamCol].Merge = true;

                // Заполнение списка студентов.
                for (int i = 0; i < students.Length; i++)
                {
                    controlpoint.Cells[$"A{studentsStartRow + i}"].Value = $"{i + 1}";
                    controlpoint.Cells[$"B{studentsStartRow + i}"].Value = students[i];
                    controlpoint.Cells[$"B{studentsStartRow + i}"].Style.Fill.SetBackground(Color.FromArgb(216, 242, 242));

                    controlpoint.Cells[studentsStartRow + i, studentStartCol, studentsStartRow + i, studentColToMerge].Merge = true;
                }

                // Заполнение списка названий колонок и коэффиецентов, покраска дат.
                for (int i = 0; i < columnsNames.Length; i++)
                {
                    controlpoint.Cells[startRow, startColumn + i].Value = columnsNames[i];
                    controlpoint.Cells[startRow, startColumn + i].Style.Fill.SetBackground(Color.FromArgb(248, 250, 202));

                    controlpoint.Cells[startRow + 1, startColumn + i].Value = columnsCoefficients[i];
                    controlpoint.Cells[startRow + 1, startColumn + i].Style.Fill.SetBackground(Color.FromArgb(164, 171, 252));

                    controlpoint.Cells[startRow + 2, startColumn + i].Style.Fill.SetBackground(Color.FromArgb(207, 167, 255));
                }

                // Заполнение столбца Итого.
                for (int i = 0; i < students.Length; i++)
                {
                    string coefficientFirstCellAddress = controlpoint.Cells[coefficientStartRow, 4].Address;
                    string coefficientLastCellAddress = controlpoint.Cells[coefficientStartRow, totalColumn - 1].Address;

                    string studentsFirstCellAddress = controlpoint.Cells[studentsStartRow + i, 4].Address;
                    string studentsLastCellAddress = controlpoint.Cells[studentsStartRow + i, totalColumn - 1].Address;

                    controlpoint.Cells[studentsStartRow + i, totalColumn].Formula = $"SUMPRODUCT({coefficientFirstCellAddress}:" +
                        $"{coefficientLastCellAddress}, {studentsFirstCellAddress}:{studentsLastCellAddress})";

                    controlpoint.Cells[studentsStartRow + i, totalColumn].Style.Fill.SetBackground(Color.FromArgb(245, 211, 181));

                    controlpoint.Cells[studentsStartRow + i, examColumn].Style.Fill.SetBackground(Color.FromArgb(234, 255, 231));

                    controlpoint.Cells[studentsStartRow + i, totalExamCol].Style.Fill.SetBackground(Color.FromArgb(252, 184, 184));
                    controlpoint.Cells[studentsStartRow + i, totalExamCol + 1].Style.Fill.SetBackground(Color.FromArgb(250, 145, 145));
                }

                // Заполнение столбца Итого с экз.
                for (int i = 0; i < students.Length; i++)
                {
                    string examCoefficientAddress = controlpoint.Cells[3, examColumn].Address;
                    string examColAddress = controlpoint.Cells[studentsStartRow + i, examColumn].Address;

                    string totalAddress = controlpoint.Cells[studentsStartRow + i, totalColumn].Address;

                    controlpoint.Cells[studentsStartRow + i, totalExamCol].Formula = $"{examCoefficientAddress} * {examColAddress} + {totalAddress}";
                }

                // Заполнение столбца текстовых оценок.
                for (int i = 0; i < students.Length; i++)
                {
                    string totalExamAddress = controlpoint.Cells[studentsStartRow + i, totalExamCol].Address;

                    controlpoint.Cells[studentsStartRow + i, totalExamCol + 1].Formula = $"IF({totalExamAddress}>=85, \"отл.\", " +
                        $"IF({totalExamAddress}>=70, \"хор.\", IF({totalExamAddress}>=50, \"удовл.\", \"неудовл.\")))";
                }

                using (ExcelRange range = controlpoint.Cells[1, 1, controlpoint.Dimension.End.Row, controlpoint.Dimension.End.Column])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.None;
                    range.Style.Border.Left.Style = ExcelBorderStyle.None;
                    range.Style.Border.Right.Style = ExcelBorderStyle.None;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.None;
                }

                // Сохранение файла
                package.SaveAs(new FileInfo(path + fileName + ".xlsx"));
            }
        }
    }
}
