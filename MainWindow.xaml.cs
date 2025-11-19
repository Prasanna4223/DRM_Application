//using System.Text;
//using System.Windows;
//using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Navigation;
//using System.Windows.Shapes;
//using ClosedXML.Excel;
//using System;
//using System.IO;
//using Outlook = Microsoft.Office.Interop.Outlook;

//namespace DRM_Management
//{
//    public partial class MainWindow : Window
//    {
//        private readonly string excelFilePath = System.IO.Path.Combine(
//    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
//    "DRM_Prasanna.xlsx");


//        public MainWindow()
//        {
//            InitializeComponent();
//            DateInput.SelectedDate = DateTime.Today;
//        }

//        private void StatusInput_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
//        {
//            if (StatusInput.SelectedItem is System.Windows.Controls.ComboBoxItem selected)
//            {
//                string status = selected.Content.ToString();
//                switch (status)
//                {
//                    case "Completed":
//                        ActivityInput.Background = new SolidColorBrush(Color.FromRgb(204, 255, 204));
//                        ActivityInput.Foreground = new SolidColorBrush(Color.FromRgb(0, 100, 0));
//                        break;
//                    case "In Progress":
//                        ActivityInput.Background = new SolidColorBrush(Color.FromRgb(255, 235, 156));
//                        ActivityInput.Foreground = new SolidColorBrush(Color.FromRgb(139, 0, 0));
//                        break;
//                    case "Pending":
//                        ActivityInput.Background = new SolidColorBrush(Colors.Orange);
//                        ActivityInput.Foreground = new SolidColorBrush(Colors.Red);
//                        break;
//                    default:
//                        ActivityInput.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
//                        ActivityInput.ClearValue(System.Windows.Controls.Control.ForegroundProperty);
//                        break;
//                }
//            }
//        }

//        private void SaveButton_Click(object sender, RoutedEventArgs e)
//        {
//            if (string.IsNullOrWhiteSpace(ActivityInput.Text))
//            {
//                MessageBox.Show("Please enter an activity.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
//                return;
//            }

//            string date = DateInput.SelectedDate?.ToShortDateString() ?? DateTime.Today.ToShortDateString();
//            string activity = ActivityInput.Text;
//            string status = (StatusInput.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content.ToString() ?? "Pending";
//            string comments = CommentsInput.Text;

//            SaveToExcel(date, activity, status, comments);

//            MessageBox.Show("Data saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

//            ClearInputs();
//        }
//        private void SendEmailButton_Click(object sender, RoutedEventArgs e)
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

//                mail.Subject = "Daily Report";
//                mail.Body = "Please find the attached Daily Report Excel file.";
//                mail.To = ""; // You can prefill a recipient or leave empty
//                mail.Attachments.Add(excelFilePath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);

//                // Display the email (popup)
//                mail.Display(true);
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show("Error sending email: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
//            }
//        }


//        private void CancelButton_Click(object sender, RoutedEventArgs e)
//        {
//            ClearInputs();
//        }

//        private void ClearInputs()
//        {
//            DateInput.SelectedDate = DateTime.Today;
//            ActivityInput.Clear();
//            CommentsInput.Clear();
//            StatusInput.SelectedIndex = -1;
//            ActivityInput.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
//            ActivityInput.ClearValue(System.Windows.Controls.Control.ForegroundProperty);
//        }

//        private void SaveToExcel(string date, string activity, string status, string comments)
//        {
//            XLWorkbook workbook;
//            IXLWorksheet worksheet;

//            if (File.Exists(excelFilePath))
//            {
//                workbook = new XLWorkbook(excelFilePath);
//                string monthName = DateTime.Now.ToString("MMMM"); 
//                worksheet = workbook.Worksheet(monthName);

//            }
//            else
//            {
//                workbook = new XLWorkbook();
//                string monthName = DateTime.Now.ToString("MMMM"); 
//                worksheet = workbook.AddWorksheet(monthName);

//                // Set header row
//                worksheet.Cell(1, 1).Value = "Date";
//                worksheet.Cell(1, 2).Value = "Activity";
//                worksheet.Cell(1, 3).Value = "Status";
//                worksheet.Cell(1, 4).Value = "Comments";

//                // Apply header formatting: yellow background + bold
//                var headerRange = worksheet.Range(1, 1, 1, 4); // Row 1, columns 1-4
//                headerRange.Style.Fill.BackgroundColor = XLColor.Yellow;
//                headerRange.Style.Font.Bold = true;
//            }

//            // Determine the next empty even row
//            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
//            int newRow = lastRow + 2; // leave one empty row
//            if (newRow % 2 != 0) newRow++; // ensure even row

//            // Fill the data
//            worksheet.Cell(newRow, 1).Value = date;
//            worksheet.Cell(newRow, 2).Value = activity;
//            worksheet.Cell(newRow, 3).Value = status;
//            worksheet.Cell(newRow, 4).Value = comments;

//            // Apply Excel formatting based on status (color coding)
//            var statusCell = worksheet.Cell(newRow, 3);
//            switch (status)
//            {
//                case "Completed":
//                    statusCell.Style.Fill.BackgroundColor = XLColor.LightGreen;
//                    statusCell.Style.Font.FontColor = XLColor.DarkGreen;
//                    break;
//                case "In Progress":
//                    statusCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 235, 156); // RGB color
//                    statusCell.Style.Font.FontColor = XLColor.DarkRed;
//                    break;
//                case "Pending":
//                    statusCell.Style.Fill.BackgroundColor = XLColor.Orange;
//                    statusCell.Style.Font.FontColor = XLColor.Red;
//                    break;
//            }

//            // Autofit columns for neat look
//            worksheet.Columns().AdjustToContents();

//            workbook.SaveAs(excelFilePath);
//           // MessageBox.Show($"File saved to:\n{excelFilePath}", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
//        }


//    }
//}


using ClosedXML.Excel;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DRM_Management
{
    public partial class MainWindow : Window
    {
        private readonly string excelFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "DailyReport.xlsx");
        private const string PlaceholderText = "Enter activity...";
        private const string ContentholderText = "Enter comments...";
        public MainWindow()
        {
            InitializeComponent();
            DateInput.SelectedDate = DateTime.Today;
        }

        private void StatusInput_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StatusInput.SelectedItem is ComboBoxItem selected)
            {
                string status = selected.Content.ToString();
                switch (status)
                {
                    case "Completed":
                        ActivityInput.Background = new SolidColorBrush(Color.FromRgb(204, 255, 204));
                        ActivityInput.Foreground = new SolidColorBrush(Color.FromRgb(0, 100, 0));
                        break;
                    case "In Progress":
                        ActivityInput.Background = new SolidColorBrush(Color.FromRgb(255, 235, 156));
                        ActivityInput.Foreground = new SolidColorBrush(Color.FromRgb(139, 0, 0));
                        break;
                    case "Pending":
                        ActivityInput.Background = new SolidColorBrush(Colors.Orange);
                        ActivityInput.Foreground = new SolidColorBrush(Colors.Red);
                        break;
                    default:
                        ActivityInput.ClearValue(Control.BackgroundProperty);
                        ActivityInput.ClearValue(Control.ForegroundProperty);
                        break;
                }
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ActivityInput.Text))
            {
                MessageBox.Show("Please enter an activity.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string date = DateInput.SelectedDate?.ToShortDateString() ?? DateTime.Today.ToShortDateString();
            string activity = ActivityInput.Text;
            string status = (StatusInput.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Pending";
            string comments = CommentsInput.Text;

            SaveToExcel(date, activity, status, comments);

            MessageBox.Show("Data saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

            ClearInputs();
        }

        private void SendEmailButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show("Excel file not found! Save it first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.Explorer explorer = outlookApp.ActiveExplorer();

                if (explorer == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select an email in Outlook Inbox first.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Outlook.MailItem selectedMail = explorer.Selection[1] as Outlook.MailItem;
                if (selectedMail == null)
                {
                    MessageBox.Show("Selected item is not an email.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Outlook.MailItem replyAll = selectedMail.ReplyAll();
                replyAll.Attachments.Add(excelFilePath);

                replyAll.Display();

                MessageBox.Show("Reply All email prepared successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating Reply All email: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            ClearInputs();
        }

        private void ClearInputs()
        {
            DateInput.SelectedDate = DateTime.Today;
            ActivityInput.Clear();
            CommentsInput.Clear();
            StatusInput.SelectedIndex = -1;
            ActivityInput.ClearValue(Control.BackgroundProperty);
            ActivityInput.ClearValue(Control.ForegroundProperty);
        }

        private void SaveToExcel(string date, string activity, string status, string comments)
        {
            XLWorkbook workbook;
            IXLWorksheet worksheet;
            string monthName = DateTime.Now.ToString("MMMM");

            if (File.Exists(excelFilePath))
            {
                workbook = new XLWorkbook(excelFilePath);
                worksheet = workbook.Worksheets.Contains(monthName)
                    ? workbook.Worksheet(monthName)
                    : workbook.AddWorksheet(monthName);

                // Create headers if empty
                if (worksheet.Cell(1, 1).IsEmpty())
                {
                    CreateExcelHeaders(worksheet);
                }
            }
            else
            {
                workbook = new XLWorkbook();
                worksheet = workbook.AddWorksheet(monthName);
                CreateExcelHeaders(worksheet);
            }

            // Determine next empty even row (leave a gap)
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            int newRow = lastRow + 2;
            if (newRow % 2 != 0) newRow++;

            // Fill data
            worksheet.Cell(newRow, 1).Value = date;
            worksheet.Cell(newRow, 2).Value = activity;
            worksheet.Cell(newRow, 3).Value = status;
            worksheet.Cell(newRow, 4).Value = comments;

            // Apply status formatting
            var statusCell = worksheet.Cell(newRow, 3);
            switch (status)
            {
                case "Completed":
                    statusCell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                    statusCell.Style.Font.FontColor = XLColor.DarkGreen;
                    break;
                case "In Progress":
                    statusCell.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 235, 156);
                    statusCell.Style.Font.FontColor = XLColor.DarkRed;
                    break;
                case "Pending":
                    statusCell.Style.Fill.BackgroundColor = XLColor.Orange;
                    statusCell.Style.Font.FontColor = XLColor.Red;
                    break;
            }

            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(excelFilePath);
        }

        private void CreateExcelHeaders(IXLWorksheet worksheet)
        {
            worksheet.Cell(1, 1).Value = "Date";
            worksheet.Cell(1, 2).Value = "Activity";
            worksheet.Cell(1, 3).Value = "Status";
            worksheet.Cell(1, 4).Value = "Comments";

            var headerRange = worksheet.Range(1, 1, 1, 4);
            headerRange.Style.Fill.BackgroundColor = XLColor.Yellow;
            headerRange.Style.Font.Bold = true;
        }

        //private void SendEmailButton_Click1(object sender, RoutedEventArgs e)
        //{
        //    try
        //    {
        //        // Path to your Excel file
        //        string excelFilePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DailyReport.xlsx");

        //        if (!File.Exists(excelFilePath))
        //        {
        //            MessageBox.Show("Excel file not found! Save it first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        //            return;
        //        }

        //        // Create Outlook application instance
        //        Outlook.Application outlookApp = new Outlook.Application();

        //        // Create a new mail item
        //        Outlook.MailItem mail = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

        //        mail.Subject = "Daily Report";
        //        mail.Body = "Please find the attached Daily Report Excel file.";
        //        mail.To = "";
        //        mail.Attachments.Add(excelFilePath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);

        //        // Display the email (popup)
        //        mail.Display(true);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error sending email: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        //    }
        //}
        private void SendEmailButton_Click1(object sender, RoutedEventArgs e)
        {
            // Create and show the EmailPopup window
            EmailPopup emailPopup = new EmailPopup();
            emailPopup.Owner = this;
            emailPopup.ShowDialog();
        }

        private void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            EmailSettingsWindow settingsWindow = new EmailSettingsWindow();
            settingsWindow.Owner = this;
            settingsWindow.Show();
        }

        private void RotatingImage_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var rotateTransform = new RotateTransform();
            RotatingImage.RenderTransform = rotateTransform;

            // Set the rotation center to the center of the image
            rotateTransform.CenterX = RotatingImage.ActualWidth / 2;
            rotateTransform.CenterY = RotatingImage.ActualHeight / 2;

            // Create the rotation animation
            var rotateAnimation = new DoubleAnimation
            {
                From = 0,
                To = 180,
                Duration = TimeSpan.FromSeconds(0.5),  // Duration of the rotation
                AutoReverse = false  // Ensures it rotates once fully
            };

            // Apply animation to the image rotation
            rotateTransform.BeginAnimation(RotateTransform.AngleProperty, rotateAnimation);
        }

        // Triggered when the mouse leaves the image (reset rotation to 0 degrees)
        private void RotatingImage_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var rotateTransform = RotatingImage.RenderTransform as RotateTransform;
            if (rotateTransform != null)
            {
                // Reset rotation to 0
                var resetAnimation = new DoubleAnimation
                {
                    From = 180,
                    To = 0,
                    Duration = TimeSpan.FromSeconds(0.5), // Resetting duration
                    AutoReverse = false
                };

                // Apply the reset animation
                rotateTransform.BeginAnimation(RotateTransform.AngleProperty, resetAnimation);
            }
        }

        private void ActivityInput_GotFocus(object sender, RoutedEventArgs e)
        {
            if (ActivityInput.Text == PlaceholderText)
            {
                ActivityInput.Text = string.Empty;
                ActivityInput.Foreground = Brushes.Black;
            }
        }

        private void ActivityInput_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ActivityInput.Text))
            {
                ActivityInput.Text = PlaceholderText;
                ActivityInput.Foreground = Brushes.Gray;
            }
        }

        private void CommentsInput_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(CommentsInput.Text))
            {
                CommentsInput.Text = ContentholderText;
                CommentsInput.Foreground = Brushes.Gray;
            }
        }

        private void CommentsInput_GotFocus(object sender, RoutedEventArgs e)
        {
            if (CommentsInput.Text == ContentholderText)
            {
                CommentsInput.Text = string.Empty;
                CommentsInput.Foreground = Brushes.Black;
            }
        }
    }
}
