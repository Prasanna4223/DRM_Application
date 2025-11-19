using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DRM_Management
{
    public static class ComboBoxExtensions
    {
        public static readonly DependencyProperty PlaceholderTextProperty =
            DependencyProperty.RegisterAttached(
                "PlaceholderText",
                typeof(string),
                typeof(ComboBoxExtensions),
                new PropertyMetadata(string.Empty, OnPlaceholderChanged));

        public static string GetPlaceholderText(DependencyObject obj) =>
            (string)obj.GetValue(PlaceholderTextProperty);

        public static void SetPlaceholderText(DependencyObject obj, string value) =>
            obj.SetValue(PlaceholderTextProperty, value);

        private static void OnPlaceholderChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ComboBox combo)
            {
                combo.Loaded -= ComboBox_Loaded;
                combo.Loaded += ComboBox_Loaded;
            }
        }

        private static void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            var combo = sender as ComboBox;
            combo.ApplyTemplate();
            var textBox = combo.Template.FindName("PART_EditableTextBox", combo) as TextBox;
            if (textBox == null) return;

            string placeholder = GetPlaceholderText(combo);

            void Update()
            {
                if (string.IsNullOrWhiteSpace(textBox.Text) || textBox.Text == placeholder)
                {
                    textBox.Text = placeholder;
                    textBox.Foreground = Brushes.Gray;
                }
                else
                {
                    textBox.Foreground = Brushes.Black;
                }
            }

            textBox.GotFocus += (s, a) =>
            {
                if (textBox.Text == placeholder) textBox.Text = "";
                textBox.Foreground = Brushes.Black;
            };

            textBox.LostFocus += (s, a) => Update();

            Update();
        }
    }
}