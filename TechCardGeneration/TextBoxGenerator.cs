using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;

namespace TechCardGeneration.Windows
{
    public static class TextBoxGenerator
    {
        public static void Generation(int numberOfTextBoxes, StackPanel TextBoxContainer, int width)
        {
            Style mainTextBoxStyle = (Style)Application.Current.FindResource("TextBoxMainStyle");
            Style robotoCFont = (Style)Application.Current.FindResource("RobotoCFont");

            for (int i = 0; i < numberOfTextBoxes; i++)
            {
                double marginRight = 0;

                if (i <= 8)
                {
                    marginRight = 10.5;
                }

                StackPanel stackPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                };

                Label label = new Label
                {
                    Content = i + 1,
                    Style = robotoCFont,
                    Margin = new Thickness(0, 7, marginRight, 0),
                    Foreground = new SolidColorBrush(Color.FromRgb(75, 135, 190)),
                    FontSize = 22,
                };

                TextBox textBox = new TextBox
                {
                    Margin = new Thickness(0, 10, 0, 0),
                    Width = width,
                    Height = 35,
                    Style = mainTextBoxStyle,
                    Padding = new Thickness(2, 0, 0, 0),
                    FontSize = 22
                };

                stackPanel.Children.Add(label);
                stackPanel.Children.Add(textBox);

                TextBoxContainer.Children.Add(stackPanel);
            }
        }
    }
}
