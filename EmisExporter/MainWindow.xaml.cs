using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;


namespace EmisExporter
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            string year = (yearSelect.SelectedItem as ComboBoxItem).Content.ToString();
            new Exporter().export(year, Progress);
            Progress.Visibility = Visibility.Visible;
        }

        private void yearSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ContinueButton.IsEnabled = true;
        }
    }

    public static class helpers
    {
        public static Func<A, R> Memoize<A, R>(this Func<A, R> f)
        {
            var map = new Dictionary<A, R>();
            return a =>
            {
                R value;
                if (map.TryGetValue(a, out value))
                    return value;
                value = f(a);
                map.Add(a, value);
                return value;
            };
        }

        public static string GetCellAddress(int col, int row)
        {
            StringBuilder sb = new StringBuilder();
            do
            {
                col--;
                sb.Insert(0, (char)('A' + (col % 26)));
                col /= 26;
            } while (col > 0);
            sb.Append(row);
            return sb.ToString();
        }
    }
}