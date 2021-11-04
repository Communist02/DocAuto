using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DocAuto
{
    class Field : ListBoxItem
    {
        public String title;
        public String value;

        public Field(string title, string value)
        {
            this.title = title;
            this.value = value;

            HorizontalContentAlignment = HorizontalAlignment.Stretch;
            Border border = new Border() { MinWidth = 100, Margin = new Thickness(10, 6, 10, 6), Padding = new Thickness(10), CornerRadius = new CornerRadius(9), Background = new SolidColorBrush(Color.FromRgb(238, 238, 238)) };
            StackPanel stackPanel = new StackPanel() { Orientation = Orientation.Vertical };
            TextBlock titleBlock = new TextBlock() { Text = title };
            TextBox textBox = new TextBox() { Text = value, MinWidth = 100};
            textBox.TextChanged += new TextChangedEventHandler(TextChanged);
            stackPanel.Children.Add(titleBlock);
            stackPanel.Children.Add(textBox);
            border.Child = stackPanel;
            AddChild(border);

            void TextChanged(object sender, TextChangedEventArgs e)
            {
                value = textBox.Text;
                MainWindow.TextChanged(title, value);
            }
        }
    }
}
