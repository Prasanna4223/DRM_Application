//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Windows;

//namespace DRM_Management
//{
//    public partial class EmailSettingsWindow : Window
//    {
//        static string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
//        private static readonly string emailFilePath= System.IO.Path.Combine(appDirectory, "Emails.txt");



//        private List<string> emails = new List<string>();

//        public EmailSettingsWindow()
//        {
//            InitializeComponent();
//            LoadEmails();
//        }

//        private void LoadEmails()
//        {
//            if (File.Exists(emailFilePath))
//                emails = File.ReadAllLines(emailFilePath).Distinct().ToList();

//            RefreshGrid();
//        }

//        private void RefreshGrid()
//        {
//            EmailGrid.ItemsSource = emails.Select((e, i) => new { Index = i + 1, Email = e }).ToList();
//        }

//        private void AddEmail_Click(object sender, RoutedEventArgs e)
//        {
//            string email = NewEmailBox.Text.Trim();

//            if (string.IsNullOrWhiteSpace(email))
//            {
//                MessageBox.Show("Please enter an email address.", "Warning");
//                return;
//            }

//            if (emails.Contains(email, StringComparer.OrdinalIgnoreCase))
//            {
//                MessageBox.Show("Email already exists!", "Duplicate");
//                return;
//            }

//            emails.Add(email);
//            NewEmailBox.Clear();
//            RefreshGrid();
//        }

//        private void UpdateEmail_Click(object sender, RoutedEventArgs e)
//        {
//            if (EmailGrid.SelectedItem == null)
//            {
//                MessageBox.Show("Select an email to update.", "Info");
//                return;
//            }

//            string newEmail = NewEmailBox.Text.Trim();
//            if (string.IsNullOrWhiteSpace(newEmail))
//            {
//                MessageBox.Show("Please enter a new email address.", "Warning");
//                return;
//            }

//            int index = EmailGrid.SelectedIndex;
//            emails[index] = newEmail;
//            RefreshGrid();
//        }

//        private void DeleteEmail_Click(object sender, RoutedEventArgs e)
//        {
//            if (EmailGrid.SelectedItem == null)
//            {
//                MessageBox.Show("Select an email to delete.", "Info");
//                return;
//            }

//            int index = EmailGrid.SelectedIndex;
//            emails.RemoveAt(index);
//            RefreshGrid();
//        }

//        private void SaveEmails_Click(object sender, RoutedEventArgs e)
//        {
//            File.WriteAllLines(emailFilePath, emails);
//            MessageBox.Show("Emails saved successfully.", "Success");
//        }
//    }
//}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace DRM_Management
{
    public partial class EmailSettingsWindow : Window
    {
        private static readonly string AppDirectory = AppDomain.CurrentDomain.BaseDirectory;
        private static readonly string EmailFilePath = Path.Combine(AppDirectory, "Emails.txt");

        private readonly List<string> _emails = new();

        public EmailSettingsWindow()
        {
            InitializeComponent();
            LoadEmails();
        }

        #region --- File I/O -------------------------------------------------
        private void LoadEmails()
        {
            if (File.Exists(EmailFilePath))
            {
                _emails.AddRange(
                    File.ReadAllLines(EmailFilePath)
                        .Select(l => l.Trim())
                        .Where(l => !string.IsNullOrWhiteSpace(l))
                        .Distinct(StringComparer.OrdinalIgnoreCase));
            }

            RefreshGrid();
        }

        private void SaveEmails()
        {
            File.WriteAllLines(EmailFilePath, _emails);
            MessageBox.Show("Emails saved successfully.", "Success",
                            MessageBoxButton.OK, MessageBoxImage.Information);
        }
        #endregion

        #region --- Grid ----------------------------------------------------
        private void RefreshGrid()
        {
            EmailGrid.ItemsSource = _emails
                .Select((e, i) => new { Index = i + 1, Email = e })
                .ToList();
        }
        #endregion

        #region --- Button Handlers -----------------------------------------
        private void AddEmail_Click(object sender, RoutedEventArgs e)
        {
            string email = NewEmailBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(email) || email == "Enter email...")
            {
                MessageBox.Show("Please enter a valid email address.", "Warning",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_emails.Contains(email, StringComparer.OrdinalIgnoreCase))
            {
                MessageBox.Show("Email already exists!", "Duplicate",
                                MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            _emails.Add(email);
            NewEmailBox.Text = "Enter email...";
            RefreshGrid();
        }

        private void UpdateEmail_Click(object sender, RoutedEventArgs e)
        {
            if (EmailGrid.SelectedItem is not { } selected) return;

            // Cast to the anonymous type shape
            var sel = (dynamic)selected;   // <-- still works because we only read properties
            // OR use a tiny helper record (recommended, see below)

            string newEmail = NewEmailBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(newEmail) || newEmail == "Enter email...")
            {
                MessageBox.Show("Please enter a new email address.", "Warning",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int idx = sel.Index - 1;
            _emails[idx] = newEmail;
            RefreshGrid();
        }

        private void DeleteEmail_Click(object sender, RoutedEventArgs e)
        {
            if (EmailGrid.SelectedItem is not { } selected) return;

            var sel = (dynamic)selected;
            int idx = sel.Index - 1;
            _emails.RemoveAt(idx);
            RefreshGrid();
        }

        private void SaveEmails_Click(object sender, RoutedEventArgs e) => SaveEmails();
        #endregion

        #region --- UI Helpers -----------------------------------------------
        private void NewEmailBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (NewEmailBox.Text == "Enter email...")
                NewEmailBox.Text = "";
        }

        private void NewEmailBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(NewEmailBox.Text))
                NewEmailBox.Text = "Enter email...";
        }

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

        #endregion

        private void NewEmailBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}
