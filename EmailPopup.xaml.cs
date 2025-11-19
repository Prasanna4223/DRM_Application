//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Shapes;
//using Outlook = Microsoft.Office.Interop.Outlook;

//namespace DRM_Management
//{
//    /// <summary>
//    /// Interaction logic for EmailPopup.xaml
//    /// </summary>
//    public partial class EmailPopup : Window
//    {
//        public EmailPopup()
//        {
//            InitializeComponent();
//            LoadEmailDropdowns();
//        }


//        private void LoadEmailDropdowns()
//        {
//            // Get the application running directory
//            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;

//            // Define the path for emails.txt
//            string emailFilePath = System.IO.Path.Combine(appDirectory, "Emails.txt");



//            if (File.Exists(emailFilePath))
//            {
//                var emails = File.ReadAllLines(emailFilePath).Distinct().ToList();

//                ToEmailDropdown.ItemsSource = emails;
//                CcEmailDropdown.ItemsSource = emails;
//            }
//        }

//        private void ToEmailDropdown_SelectionChanged(object sender, SelectionChangedEventArgs e)
//        {

//        }

//        private void CancelButton_Click(object sender, RoutedEventArgs e)
//        {
//            ClearInputs();
//        }
//        private void ClearInputs()
//        {
//            ToEmailDropdown.Text = string.Empty;
//            CcEmailDropdown.Text = string.Empty;    
//            SubjectInput.Text = string.Empty;
//            ContentInput.Text = string.Empty;
//        }
//        private void SendButton_Click(object sender, RoutedEventArgs e)
//        {

//            try
//            {
//                // Path to your Excel file
//                string excelFilePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DailyReport.xlsx");

//                if (!File.Exists(excelFilePath))
//                {
//                    MessageBox.Show("Excel file not found! Save it first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
//                    return;
//                }

//                // Create Outlook application instance
//                Outlook.Application outlookApp = new Outlook.Application();

//                // Create a new mail item
//                Outlook.MailItem mail = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

//                mail.Subject = "Daily Work Activity Reporting";
//                mail.Body = $"Dear Sir,\r\n                Please find the attached excel for daily work activity update for {DateTime.Now}.\r\n\r\nRegards\r\nPrasanna\r\n";
//                mail.To = "";
//                mail.Attachments.Add(excelFilePath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);

//                // Display the email (popup)
//                mail.Display(true);
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show("Error sending email: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
//            }
//        }

//        private void CcEmailDropdown_SelectionChanged(object sender, SelectionChangedEventArgs e)
//        {

//        }

//        private void ToEmailDropdown_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
//        {

//        }
//        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
//        {
//            if (e.ChangedButton == MouseButton.Left) DragMove();
//        }

//        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
//    }
//}
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office.CustomUI;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace DRM_Management
{
    /// <summary>
    /// Interaction logic for EmailPopup.xaml
    /// </summary>
    public partial class EmailPopup : Window
    {
        private Application _outlookApp;
        private bool _isOutlookCreated = false;

        public EmailPopup()
        {
            InitializeComponent();
            this.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) DragMove(); };
            LoadEmailDropdowns();
            SetDefaultContent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.MouseLeftButtonDown += (s, ev) =>
            {
                if (ev.ButtonState == MouseButtonState.Pressed) DragMove();
            };
        }
        private void LoadEmailDropdowns()
        {
            try
            {
                string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string emailFilePath = Path.Combine(appDirectory, "Emails.txt");

                if (!File.Exists(emailFilePath)) return;

                var emails = File.ReadAllLines(emailFilePath)
                    .Select(e => e.Trim())
                    .Where(e => !string.IsNullOrWhiteSpace(e) && IsValidEmail(e))
                    .Distinct()
                    .OrderBy(e => e)
                    .ToList();

                ToEmailDropdown.ItemsSource = emails;
                CcEmailDropdown.ItemsSource = emails;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Failed to load email list: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private void SetDefaultContent()
        {
            SubjectInput.Text = "Daily Work Activity Reporting";
            ContentInput.Text =
                "Dear Sir,\r\n\r\n" +
                $"Please find the attached Excel report for daily work activity update for ({DateTime.Now.ToString("dd'/'MM'/'yyyy", System.Globalization.CultureInfo.InvariantCulture)}).\r\n\r\n" +
                "Regards,\r\n" +
                "Prasanna";
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            string toEmail = ToEmailDropdown.Text.Trim();
            string ccEmail = CcEmailDropdown.Text.Trim();

            // Validation
            if (string.IsNullOrEmpty(toEmail))
            {
                ShowWarning("Please enter a recipient email in 'To' field.");
                return;
            }

            if (!IsValidEmail(toEmail))
            {
                ShowWarning("Please enter a valid email address in 'To' field.");
                return;
            }

            if (!string.IsNullOrEmpty(ccEmail) && !ccEmail.Split(';').All(IsValidEmail))
            {
                ShowWarning("One or more CC emails are invalid.");
                return;
            }

            string excelPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "DailyReport.xlsx");

            if (!File.Exists(excelPath))
            {
                ShowError("Excel file not found on Desktop:\n" + excelPath);
                return;
            }

            try
            {
                // Reuse existing Outlook or create new
                _outlookApp = GetOutlookInstance();
                if (_outlookApp == null)
                {
                    ShowError("Microsoft Outlook is not installed or not accessible.");
                    return;
                }

                MailItem mail = _outlookApp.CreateItem(OlItemType.olMailItem) as MailItem;
                if (mail == null)
                {
                    ShowError("Failed to create email item.");
                    return;
                }

                mail.To = toEmail;
                if (!string.IsNullOrEmpty(ccEmail)) mail.CC = ccEmail;
                mail.Subject = SubjectInput.Text.Trim();
                mail.Body = ContentInput.Text.Trim() +
                           $"\r\n\r\nReport generated on: {DateTime.Now:dddd, MMMM d, yyyy 'at' h:mm tt}";

                // Attach file
                mail.Attachments.Add(excelPath, OlAttachmentType.olByValue, 1, "DailyReport.xlsx");

                // Show email (user can edit before sending)
                mail.Display(false);

                // Success: close window
                this.Close();
            }
            catch (System.Exception ex)
            {
                ShowError($"Failed to create email:\n{ex.Message}");
            }
            finally
            {
                // Optional: Release COM objects if not reusing Outlook
                // ReleaseComObject(mail);
            }
        }

        private Application GetOutlookInstance()
        {
            try
            {
                // Try to get running Outlook instance (works in .NET 6+)
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType != null)
                {
                    _outlookApp = (Application)Activator.CreateInstance(outlookType);
                    _isOutlookCreated = true; // We created it
                    return _outlookApp;
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to get Outlook: " + ex.Message);
                // Fall back to new instance
            }

            try
            {
                // Fallback: create new instance
                _outlookApp = new Application();
                _isOutlookCreated = true;
                return _outlookApp;
            }
            catch
            {
                return null;
            }
        }
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Are you sure you want to cancel sending the email?",
                "Confirm Cancel",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                this.Close();
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // Clean up Outlook if we created it
            if (_isOutlookCreated && _outlookApp != null)
            {
                try
                {
                    // Do NOT quit Outlook if user has it open
                    // Just release reference
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_outlookApp);
                }
                catch { }
                _outlookApp = null;
            }
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(message, "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void CcEmailDropdown_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        private void ToEmailDropdown_SelectionChanged_1(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            System.Windows.Controls.ComboBox comboBox = sender as System.Windows.Controls.ComboBox;

            if (comboBox != null && comboBox.SelectedItem != null)
            {
                string selectedText = comboBox.SelectedItem.ToString();

                // This ensures the selected text appears in the dropdown box
                comboBox.Text = selectedText;
                comboBox.Tag = selectedText;
            }
        }

        private void SubjectInput_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }

        // Optional: Full COM cleanup (use only if needed)
        // private void ReleaseComObject(object obj)
        // {
        //     try
        //     {
        //         if (obj != null)
        //             System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //     }
        //     catch { }
        //     finally
        //     {
        //         obj = null;
        //     }
        //     GC.Collect();
        //     GC.WaitForPendingFinalizers();
        // }
    }
}