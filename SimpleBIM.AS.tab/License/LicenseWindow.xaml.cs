using System;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace SimpleBIM.AS.tab.License
{
    public partial class LicenseWindow : Window
    {
        private readonly LicenseManager _licenseManager;
        private bool _isOnline;

        public LicenseWindow()
        {
            InitializeComponent();
            _licenseManager = LicenseManager.Instance;
            _licenseManager.LicenseStatusChanged += OnLicenseStatusChanged;

            CheckConnection();
            LoadLicenseInfo();

            this.Loaded += Window_Loaded;
        }

        private void CheckConnection()
        {
            try
            {
                _isOnline = NetworkInterface.GetIsNetworkAvailable();
                UpdateConnectionStatus(_isOnline);
            }
            catch
            {
                _isOnline = false;
                UpdateConnectionStatus(false);
            }
        }

        private void UpdateConnectionStatus(bool isOnline)
        {
            Dispatcher.Invoke(() =>
            {
                ConnectionIndicator.Fill = isOnline ? Brushes.Green : Brushes.Red;
                ConnectionStatusText.Text = isOnline ? "Online" : "Offline";
                ConnectionStatusText.Foreground = isOnline ? Brushes.Green : Brushes.Red;
            });
        }

        private void LoadLicenseInfo()
        {
            var license = _licenseManager.CurrentLicense;

            if (license != null && _licenseManager.ValidateOffline())
            {
                // Hiển thị thông tin license
                LicenseInfoBorder.Visibility = Visibility.Visible;
                StatusMessage.Visibility = Visibility.Collapsed;

                LicenseStatusText.Text = license.IsActive ? "Active" : "Inactive";
                LicenseStatusText.Foreground = license.IsActive ?
                    new SolidColorBrush(Color.FromRgb(40, 167, 69)) :
                    new SolidColorBrush(Color.FromRgb(220, 53, 69));

                ExpiryDateText.Text = $"Expires: {license.ExpiredAt.ToLocalTime():dd/MM/yyyy HH:mm}";
                MachineText.Text = $"Machine: {Environment.MachineName}";

                ValidateButton.Content = "Revalidate";
                LicenseKeyTextBox.Text = license.Key;
            }
            else
            {
                // Không có license hoặc license không hợp lệ
                LicenseInfoBorder.Visibility = Visibility.Collapsed;
                StatusMessage.Visibility = Visibility.Visible;

                if (license == null)
                {
                    StatusMessage.Text = "Chức năng của add in hiện không khả dụng";
                }
                else
                {
                    StatusMessage.Text = "License cần được xác thực lại online";
                }

                ValidateButton.Content = "Validate Key";
            }
        }

        private async void ValidateButton_Click(object sender, RoutedEventArgs e)
        {
            var key = LicenseKeyTextBox.Text?.Trim();

            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("Please enter a license key", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ValidateButton.IsEnabled = false;
            ValidateButton.Content = "Validating...";

            try
            {
                // Lấy phiên bản Revit từ App
                string revitVersion = "2023"; // Bạn có thể truyền từ App.cs

                var result = await _licenseManager.ValidateOnlineAsync(key, revitVersion);

                if (result.valid)
                {
                    MessageBox.Show($"License validated successfully!\nExpires: {result.expired_at?.ToLocalTime():dd/MM/yyyy}\nNote: {result.note}",
                        "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    LoadLicenseInfo();
                }
                else
                {
                    MessageBox.Show("Invalid license key or activation failed",
                        "Validation Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Validation error: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ValidateButton.IsEnabled = true;
                ValidateButton.Content = "Validate Key";
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to clear the license?",
                "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _licenseManager.ClearLicense();
                LoadLicenseInfo();
                LicenseKeyTextBox.Clear();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OnLicenseStatusChanged(object sender, LicenseManager.LicenseStatusChangedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                if (e.IsValid)
                {
                    MessageBox.Show(e.Message, "License Updated",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    LoadLicenseInfo();
                }
                else
                {
                    MessageBox.Show(e.Message, "License Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            });
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Kiểm tra connection mỗi 30 giây
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(30);
            timer.Tick += (s, args) => CheckConnection();
            timer.Start();
        }
    }
}