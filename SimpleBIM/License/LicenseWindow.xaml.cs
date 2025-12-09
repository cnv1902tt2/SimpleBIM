using System;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using static SimpleBIM.License.LicenseManager;

namespace SimpleBIM.License
{
    public partial class LicenseWindow : Window
    {
        private readonly LicenseManager _licenseManager;
        private bool _isOnline;
        private bool _isValidating = false;

        public LicenseWindow()
        {
            InitializeComponent();

            _licenseManager = LicenseManager.Instance;
            _licenseManager.LicenseStatusChanged += OnLicenseStatusChanged;

            CheckConnection();
            LoadLicenseInfo();

            Loaded += Window_Loaded;
        }

        private void CheckConnection()
        {
            _isOnline = NetworkInterface.GetIsNetworkAvailable();
            UpdateConnectionStatus(_isOnline);
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

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var lic = _licenseManager.CurrentLicense;
                if (!_isOnline || lic == null || string.IsNullOrEmpty(lic.Key))
                    return;

                ValidateButton.IsEnabled = false;
                ValidateButton.Content = "Đang kiểm tra...";

                ValidateResponse onlineResult = null;
                try
                {
                    string revitVersion = App.GetRevitVersion();
                    onlineResult = await _licenseManager.ValidateOnlineAsync(lic.Key, revitVersion);
                }
                catch
                {
                    // silent: nếu lỗi mạng, chúng ta vẫn hiển thị trạng thái offline phía dưới
                }
                finally
                {
                    // luôn bật nút lại và cập nhật content theo trạng thái offline hiện tại (im lặng)
                    ValidateButton.IsEnabled = true;
                    ValidateButton.Content = _licenseManager.ValidateOffline() ? "Revalidate" : "Validate Key";
                }

                // Nếu server trả về kết quả online và nó báo không hợp lệ → show 1 message duy nhất
                if (onlineResult != null && (!onlineResult.valid || !onlineResult.is_active))
                {
                    var note = onlineResult.note ?? "License đã bị thu hồi hoặc không còn hợp lệ";
                    StatusMessage.Text = note;
                    StatusMessage.Visibility = Visibility.Visible;
                    LoadLicenseInfo();
                    return;
                }

                // bình thường, chỉ load UI (nếu onlineResult == null là lỗi mạng, LoadLicenseInfo sẽ dùng offline info)
                LoadLicenseInfo();
            }
            catch { }
        }


        // -----------------------------
        // UI CHỈ HIỂN THỊ – KHÔNG XỬ LÝ LOGIC
        // -----------------------------
        private void LoadLicenseInfo()
        {
            var lic = _licenseManager.CurrentLicense;

            if (lic == null)
            {
                ShowNoLicenseUI();
                return;
            }

            ShowLicenseInfoBox(lic);

            if (_licenseManager.ValidateOffline())
            {
                StatusMessage.Visibility = Visibility.Collapsed;
                ValidateButton.Content = "Revalidate";
            }
            else
            {
                ShowOfflineErrorMessage(lic);
                ValidateButton.Content = "Validate Key";
            }
        }
        private void ShowOfflineErrorMessage(LicenseInfo lic)
        {
            StatusMessage.Visibility = Visibility.Visible;
            StatusMessage.Foreground = Brushes.Red;

            if (lic.ExpiredAt < DateTime.UtcNow)
                StatusMessage.Text = $"License đã hết hạn vào {lic.ExpiredAt.ToLocalTime():dd/MM/yyyy}\nLiên hệ với tác giả để gia hạn";
            else if (!lic.IsActive)
                StatusMessage.Text = $"License đã bị thu hồi\nVui lòng liên hệ tác giả để mở khóa.";
            else if (lic.MachineHash != _licenseManager.GetMachineHash())
                StatusMessage.Text = "License đã được kích hoạt ở máy khác\nTác giả nghiêm cấm các hành vi gian lận.";
            else
                StatusMessage.Text = "License không hợp lệ\nLiên hệ tác giả nhận license dùng thử.";
        }


        private void ShowLicenseInfoBox(LicenseInfo lic)
        {
            LicenseInfoBorder.Visibility = Visibility.Visible;

            LicenseStatusText.Text = lic.IsActive ? "Active" : "Inactive";
            LicenseStatusText.Foreground = lic.IsActive ?
                new SolidColorBrush(Color.FromRgb(40, 167, 69)) :
                Brushes.Orange;

            ExpiryDateText.Text = $"Expires: {lic.ExpiredAt.ToLocalTime():dd/MM/yyyy HH:mm}";
            MachineText.Text = $"Machine: {Environment.MachineName}";

            LicenseKeyTextBox.Text = lic.Key;
            LicenseKeyTextBox.IsEnabled = false;
        }


        // --- UI trạng thái không có license ---
        private void ShowNoLicenseUI()
        {
            LicenseInfoBorder.Visibility = Visibility.Collapsed;
            StatusMessage.Visibility = Visibility.Visible;

            StatusMessage.Text = "Chức năng của add-in hiện không khả dụng";
            StatusMessage.Foreground = Brushes.Red;

            LicenseKeyTextBox.Clear();
            LicenseKeyTextBox.IsEnabled = true;
            ValidateButton.Content = "Validate Key";
        }

        // -----------------------------
        // NÚT VALIDATE KEY
        // -----------------------------
        private async void ValidateButton_Click(object sender, RoutedEventArgs e)
        {
            var key = LicenseKeyTextBox.Text?.Trim();
            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("Vui lòng nhập license key", "Thiếu thông tin",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!_isOnline)
            {
                MessageBox.Show("Không có kết nối Internet!", "Lỗi kết nối",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            ValidateButton.IsEnabled = false;
            ValidateButton.Content = "Đang xác thực...";
            _isValidating = true;

            try
            {
                string revitVersion = App.GetRevitVersion();
                var result = await _licenseManager.ValidateOnlineAsync(key, revitVersion);

                if (result.valid && result.is_active)
                {
                    MessageBox.Show(
                        $"License kích hoạt thành công!\nHạn sử dụng: {result.expired_at?.ToLocalTime():dd/MM/yyyy HH:mm}",
                        "Thành công",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );
                }
                else
                {
                    MessageBox.Show(
                        $"Xác thực thất bại!\n{result.note ?? "License không hợp lệ"}",
                        "Lỗi",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                }

                LoadLicenseInfo();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi kết nối:\n{ex.Message}",
                    "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isValidating = false;
                ValidateButton.IsEnabled = true;
                ValidateButton.Content = "Validate Key";
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(
                    "Bạn có chắc chắn muốn xóa license?",
                    "Xác nhận",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question
                ) == MessageBoxResult.Yes)
            {
                _licenseManager.ClearLicense();
                LoadLicenseInfo();

                MessageBox.Show("Đã xóa license", "Thông báo",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // -----------------------------
        // NHẬN EVENT TỪ LICENSE MANAGER
        // -----------------------------
        private void OnLicenseStatusChanged(object sender, LicenseManager.LicenseStatusChangedEventArgs e)
        {
            if (_isValidating) return;

            Dispatcher.Invoke(() =>
            {
                StatusMessage.Text = e.Message;
                StatusMessage.Visibility = Visibility.Visible;
                StatusMessage.Foreground = Brushes.Red;

                // Cập nhật UI nhẹ: chỉ gọi LoadLicenseInfo nếu cửa sổ đang hiển thị và không đang xử lý validate
                if (this.IsLoaded && !_isValidating)
                {
                    LoadLicenseInfo();
                }
            });
        }

    }
}
