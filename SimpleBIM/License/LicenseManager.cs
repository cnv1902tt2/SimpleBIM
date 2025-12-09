using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
namespace SimpleBIM.License
{
    public class LicenseManager
    {
        private static readonly string LicenseFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SimpleBIM", "license.json");
        private static readonly string ApiBaseUrl = "https://apikeymanagement.onrender.com/keys/validate";
        private static readonly HttpClient httpClient = new HttpClient();
        public LicenseInfo CurrentLicense { get; private set; }
        public event EventHandler<LicenseStatusChangedEventArgs> LicenseStatusChanged;
        private static readonly byte[] Key = new byte[]
                {
            0x2B, 0x7E, 0x15, 0x16, 0x28, 0xAE, 0xD2, 0xA6,
            0xAB, 0xF7, 0x15, 0x88, 0x09, 0xCF, 0x4F, 0x3C,
            0x76, 0x2E, 0x71, 0x60, 0xF3, 0x8B, 0x4D, 0xA5,
            0x6A, 0x78, 0x4D, 0x90, 0x41, 0xD3, 0xC4, 0x6F
                };
        private static readonly byte[] IV = new byte[]
        {
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07,
            0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F
        };
        public class LicenseInfo
        {
            public string Key { get; set; }
            public DateTime ExpiredAt { get; set; }
            public DateTime LastOnlineCheck { get; set; }
            public DateTime LastOfflineCheck { get; set; }
            public string MachineHash { get; set; }
            public bool IsActive { get; set; }
            public string Note { get; set; }
        }
        public class LicenseStatusChangedEventArgs : EventArgs
        {
            public bool IsValid { get; set; }
            public string Message { get; set; }
        }
        public class ValidateRequest
        {
            public string key_value { get; set; }
            public string machine_name { get; set; }
            public string os_version { get; set; }
            public string revit_version { get; set; }
            public string cpu_info { get; set; }
            public string ip_address { get; set; }
            public string machine_hash { get; set; }
        }
        public class ValidateResponse
        {
            public bool valid { get; set; }
            public DateTime? expired_at { get; set; }
            public bool is_active { get; set; }
            public string note { get; set; }
            public string machine_hash { get; set; }
        }
        private static LicenseManager _instance;
        public static LicenseManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new LicenseManager();
                }
                return _instance;
            }
        }
        private LicenseManager()
        {
            LoadLicense();
        }
        public static string GetLicenseFilePath() => LicenseFilePath;
        private void LoadLicense()
        {
            try
            {
                if (File.Exists(LicenseFilePath))
                {
                    var encryptedContent = File.ReadAllText(LicenseFilePath);
                    var json = DecryptString(encryptedContent);
                    CurrentLicense = JsonConvert.DeserializeObject<LicenseInfo>(json);
                    System.Diagnostics.Debug.WriteLine($"License loaded: Key={CurrentLicense?.Key}, IsActive={CurrentLicense?.IsActive}");
                }
                else
                {
                    CurrentLicense = null;
                    System.Diagnostics.Debug.WriteLine("No license file found");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Load license error: {ex.Message}");
                CurrentLicense = null;
            }
        }
        private void SaveLicense()
        {
            try
            {
                var directory = Path.GetDirectoryName(LicenseFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                var json = JsonConvert.SerializeObject(CurrentLicense);
                var encrypted = EncryptString(json);
                File.WriteAllText(LicenseFilePath, encrypted);
                System.Diagnostics.Debug.WriteLine($"License saved: Key={CurrentLicense?.Key}, IsActive={CurrentLicense?.IsActive}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Save error: {ex.Message}");
            }
        }
        public string GetMachineHash()
        {
            try
            {
                var builder = new StringBuilder();
                using (var searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var processorId = obj["ProcessorId"];
                        builder.Append(processorId != null ? processorId.ToString() : "");
                        break;
                    }
                }
                using (var searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var serialNumber = obj["SerialNumber"];
                        builder.Append(serialNumber != null ? serialNumber.ToString() : "");
                        break;
                    }
                }
                var mac = NetworkInterface.GetAllNetworkInterfaces()
                                .Where(nic => nic.OperationalStatus == OperationalStatus.Up &&
                                nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                                .FirstOrDefault()?.GetPhysicalAddress();
                builder.Append(mac != null ? mac.ToString() : "");
                // Compute hash
                using (var sha256 = SHA256.Create())
                {
                    var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(builder.ToString()));
                    return BitConverter.ToString(hash).Replace("-", "").ToLower();
                }
            }
            catch
            {
                return Environment.MachineName;
            }
        }
        public async Task<ValidateResponse> ValidateOnlineAsync(string key, string revitVersion = "Unknown")
        {
            try
            {
                var currentMachineHash = GetMachineHash();

                var request = new ValidateRequest
                {
                    key_value = key,
                    machine_name = Environment.MachineName,
                    os_version = Environment.OSVersion.VersionString,
                    revit_version = revitVersion,
                    cpu_info = GetCpuInfo(),
                    ip_address = GetLocalIPAddress(),
                    machine_hash = currentMachineHash
                };

                var jsonContent = JsonConvert.SerializeObject(request);
                var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                var response = await httpClient.PostAsync(ApiBaseUrl, httpContent);

                if (response.IsSuccessStatusCode)
                {
                    var responseString = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<ValidateResponse>(responseString);

                    if (result.valid && result.is_active)
                    {
                        if (!string.IsNullOrEmpty(result.machine_hash) && result.machine_hash != currentMachineHash)
                        {
                            CurrentLicense = null;
                            ClearLicense();
                            LicenseStatusChanged?.Invoke(this, new LicenseStatusChangedEventArgs
                            {
                                IsValid = false,
                                Message = "License đã được kích hoạt trên máy khác\n Đây là hành vi cấm khi sử dụng add in!"
                            });

                            return new ValidateResponse
                            {
                                valid = false,
                                note = "License đã được kích hoạt trên máy khác"
                            };
                        }
                        var newExpire = DateTime.UtcNow.AddSeconds(15);
                        CurrentLicense = new LicenseInfo
                        {
                            Key = key,
                            ExpiredAt = result.expired_at ?? DateTime.MaxValue,
                            LastOnlineCheck = DateTime.UtcNow,
                            LastOfflineCheck = DateTime.UtcNow,
                            MachineHash = currentMachineHash,
                            IsActive = true,
                            Note = result.note ?? "License hợp lệ"
                        };

                        SaveLicense();
                        return result;
                    }
                    if (!result.is_active)
                    {
                        CurrentLicense = new LicenseInfo
                        {
                            Key = key,
                            ExpiredAt = result.expired_at ?? DateTime.UtcNow,
                            LastOnlineCheck = DateTime.UtcNow,
                            LastOfflineCheck = DateTime.UtcNow,
                            MachineHash = currentMachineHash,
                            IsActive = false,
                            Note = result.note ?? "License đã bị thu hồi"
                        };

                        SaveLicense();

                        LicenseStatusChanged?.Invoke(this, new LicenseStatusChangedEventArgs
                        {
                            IsValid = false,
                            Message = result.note ?? "\nVui lòng liên hệ với tác giả để mở khóa"
                        });
                        return result;
                    }
                    
                    CurrentLicense = new LicenseInfo
                    {
                        Key = key,
                        ExpiredAt = result.expired_at ?? DateTime.UtcNow,
                        LastOnlineCheck = DateTime.UtcNow,
                        LastOfflineCheck = DateTime.UtcNow,
                        MachineHash = currentMachineHash,
                        IsActive = false,
                        Note = result.note ?? "License đã hết thời hạn sử dụng"
                    };

                    SaveLicense();

                    LicenseStatusChanged?.Invoke(this, new LicenseStatusChangedEventArgs
                    {
                        IsValid = false,
                        Message = result.note ?? "\nVui lòng liên hệ với tác giả để cập nhật lại license mới"
                    });
                    return result;
                }
                return new ValidateResponse { valid = false, note = "Lỗi kết nối server" };
            }
            catch (Exception ex)
            {
                return new ValidateResponse { valid = false, note = $"Lỗi: {ex.Message}" };
            }
        }
        public bool ValidateOffline()
        {
            try
            {
                if (CurrentLicense == null)
                    return false;

                if (CurrentLicense.ExpiredAt < DateTime.UtcNow)
                    return false;

                var currentHash = GetMachineHash();
                if (CurrentLicense.MachineHash != currentHash)
                    return false;

                if (!CurrentLicense.IsActive)
                    return false;

                CurrentLicense.LastOfflineCheck = DateTime.UtcNow;
                SaveLicense();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void ClearLicense()
        {
            CurrentLicense = null;
            try
            {
                if (File.Exists(LicenseFilePath))
                {
                    File.Delete(LicenseFilePath);
                    System.Diagnostics.Debug.WriteLine("License cleared and file deleted");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Clear license error: {ex.Message}");
            }
        }
        private string GetCpuInfo()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT Name, NumberOfCores FROM Win32_Processor"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var name = obj["Name"];
                        var cores = obj["NumberOfCores"];
                        return $"{name} ({cores} cores)";
                    }
                }
            }
            catch
            {
                return "Unknown";
            }
            return "Unknown";
        }
        private static string EncryptString(string plainText)
        {
            using (var aes = Aes.Create())
            {
                aes.Key = Key;
                aes.IV = IV;
                using (var encryptor = aes.CreateEncryptor())
                {
                    var plainBytes = Encoding.UTF8.GetBytes(plainText);
                    var encrypted = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length);
                    using (var hmac = new HMACSHA256(Key))
                    {
                        var hash = hmac.ComputeHash(encrypted);
                        var final = new byte[hash.Length + encrypted.Length];
                        Buffer.BlockCopy(hash, 0, final, 0, hash.Length);
                        Buffer.BlockCopy(encrypted, 0, final, hash.Length, encrypted.Length);
                        return Convert.ToBase64String(final);
                    }
                }
            }
        }
        private static string DecryptString(string cipherText)
        {
            var allBytes = Convert.FromBase64String(cipherText);
            var receivedHmac = new byte[32];
            var encryptedBytes = new byte[allBytes.Length - 32];
            Buffer.BlockCopy(allBytes, 0, receivedHmac, 0, 32);
            Buffer.BlockCopy(allBytes, 32, encryptedBytes, 0, encryptedBytes.Length);
            using (var hmac = new HMACSHA256(Key))
            {
                var calculatedHmac = hmac.ComputeHash(encryptedBytes);
                if (!receivedHmac.SequenceEqual(calculatedHmac))
                    throw new CryptographicException("License file tampered!");
            }
            using (var aes = Aes.Create())
            {
                aes.Key = Key;
                aes.IV = IV;
                using (var decryptor = aes.CreateDecryptor())
                {
                    var decryptedBytes = decryptor.TransformFinalBlock(encryptedBytes, 0, encryptedBytes.Length);
                    return Encoding.UTF8.GetString(decryptedBytes);
                }
            }
        }
        private string GetLocalIPAddress()
        {
            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        return ip.ToString();
                    }
                }
            }
            catch { }
            return "Unknown";
        }
    }
}