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


namespace SimpleBIM.AS.tab.License
{
    public class LicenseManager
    {
        private static readonly string LicenseFilePath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "SimpleBIM", "license.json");

        private static readonly string ApiBaseUrl = "https://apikeymanagement.onrender.com/keys/validate";

        private static readonly HttpClient httpClient = new HttpClient();

        public LicenseInfo CurrentLicense { get; private set; }
        public event EventHandler<LicenseStatusChangedEventArgs> LicenseStatusChanged;

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

        private void LoadLicense()
        {
            try
            {
                if (File.Exists(LicenseFilePath))
                {
                    var json = DecryptString(File.ReadAllText(LicenseFilePath));
                    CurrentLicense = JsonConvert.DeserializeObject<LicenseInfo>(json);
                }
            }
            catch
            {
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
                File.WriteAllText(LicenseFilePath, EncryptString(json));
            }
            catch
            {
                // Ignore save errors
            }
        }

        public string GetMachineHash()
        {
            try
            {
                var builder = new StringBuilder();

                // CPU Info
                using (var searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var processorId = obj["ProcessorId"];
                        builder.Append(processorId != null ? processorId.ToString() : "");
                        break;
                    }
                }

                // Motherboard Serial
                using (var searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var serialNumber = obj["SerialNumber"];
                        builder.Append(serialNumber != null ? serialNumber.ToString() : "");
                        break;
                    }
                }

                // MAC Address
                var mac = NetworkInterface.GetAllNetworkInterfaces()
                    .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
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
                return Environment.MachineName; // Fallback
            }
        }

        public async Task<ValidateResponse> ValidateOnlineAsync(string key, string revitVersion = "Unknown")
        {
            try
            {
                var request = new ValidateRequest
                {
                    key_value = key,
                    machine_name = Environment.MachineName,
                    os_version = Environment.OSVersion.VersionString,
                    revit_version = revitVersion,
                    cpu_info = GetCpuInfo(),
                    ip_address = GetLocalIPAddress(),
                    machine_hash = GetMachineHash()
                };

                var jsonContent = JsonConvert.SerializeObject(request);
                var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync(ApiBaseUrl, httpContent);

                if (response.IsSuccessStatusCode)
                {
                    var responseString = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<ValidateResponse>(responseString);

                    if (result.valid)
                    {
                        CurrentLicense = new LicenseInfo
                        {
                            Key = key,
                            ExpiredAt = result.expired_at ?? DateTime.MaxValue,
                            LastOnlineCheck = DateTime.UtcNow,
                            LastOfflineCheck = DateTime.UtcNow,
                            MachineHash = GetMachineHash(),
                            IsActive = result.is_active,
                            Note = result.note
                        };

                        SaveLicense();

                        if (LicenseStatusChanged != null)
                        {
                            LicenseStatusChanged(this, new LicenseStatusChangedEventArgs
                            {
                                IsValid = true,
                                Message = "License validated successfully"
                            });
                        }
                    }
                    else
                    {
                        if (CurrentLicense != null)
                        {
                            CurrentLicense.IsActive = false;                           // khóa luôn
                            CurrentLicense.ExpiredAt = DateTime.UtcNow.AddDays(-1);    // cho hết hạn ngay
                            CurrentLicense.LastOnlineCheck = DateTime.UtcNow;
                            CurrentLicense.Note = result.note ?? "License đã bị thu hồi hoặc không hợp lệ";
                            SaveLicense(); // GHI ĐÈ FILE JSON NGAY LẬP TỨC
                        }

                        // Phát sự kiện để form cập nhật lại UI (hiện đỏ, hiện thông báo, v.v.)
                        LicenseStatusChanged?.Invoke(this, new LicenseStatusChangedEventArgs
                        {
                            IsValid = false,
                            Message = "License không hợp lệ hoặc đã bị thu hồi"
                        });
                    }

                        return result;
                }

                return new ValidateResponse { valid = false };
            }
            catch
            {
                return new ValidateResponse { valid = false };
            }
        }

        public bool ValidateOffline()
        {
            try
            {
                if (CurrentLicense == null)
                    return false;

                // Kiểm tra nếu đã quá 1 ngày kể từ lần check online cuối
                //if ((DateTime.UtcNow - CurrentLicense.LastOnlineCheck).TotalDays > 1)
                //{
                //    // Yêu cầu check online nếu quá 1 ngày
                //    return false;
                //}

                // Kiểm tra expired date
                if (CurrentLicense.ExpiredAt < DateTime.UtcNow)
                    return false;

                // Kiểm tra machine hash (chống copy license)
                if (CurrentLicense.MachineHash != GetMachineHash())
                    return false;

                // Kiểm tra trạng thái active
                if (!CurrentLicense.IsActive)
                    return false;

                // Cập nhật lần check offline
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
                    File.Delete(LicenseFilePath);
            }
            catch { }

            if (LicenseStatusChanged != null)
            {
                LicenseStatusChanged(this, new LicenseStatusChangedEventArgs
                {
                    IsValid = false,
                    Message = "License cleared"
                });
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

        private static readonly byte[] Key = Convert.FromBase64String("u7K9pL2mQ8vN4xR5tW3yZ1aB6cD8eF0g"); // 32 bytes
        private static readonly byte[] IV = Convert.FromBase64String("hJ3kL9mN2pQ5rT8u");     // 16 bytes

        private static string EncryptString(string plainText)
        {
            var aes = Aes.Create();
            aes.Key = Key;
            aes.IV = IV;
            var encryptor = aes.CreateEncryptor();
            var plainBytes = Encoding.UTF8.GetBytes(plainText);
            var encrypted = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length);
            var hmac = new HMACSHA256(Key);
            var hash = hmac.ComputeHash(encrypted);
            var final = new byte[hash.Length + encrypted.Length];
            Buffer.BlockCopy(hash, 0, final, 0, hash.Length);
            Buffer.BlockCopy(encrypted, 0, final, hash.Length, encrypted.Length);
            return Convert.ToBase64String(final);
        }

        private static string DecryptString(string cipherText)
        {
            try
            {
                var allBytes = Convert.FromBase64String(cipherText);

                // Tách HMAC (32 byte đầu) và dữ liệu đã mã hóa
                var receivedHmac = new byte[32];
                var encryptedBytes = new byte[allBytes.Length - 32];
                Buffer.BlockCopy(allBytes, 0, receivedHmac, 0, 32);
                Buffer.BlockCopy(allBytes, 32, encryptedBytes, 0, encryptedBytes.Length);

                // Kiểm tra tính toàn vẹn
                using (var hmac = new HMACSHA256(Key))
                {
                    var calculatedHmac = hmac.ComputeHash(encryptedBytes);
                    if (!receivedHmac.SequenceEqual(calculatedHmac))
                        throw new CryptographicException("License file đã bị sửa đổi!");
                }

                // Giải mã AES
                var aes = Aes.Create();
                aes.Key = Key;
                aes.IV = IV;
                var decryptor = aes.CreateDecryptor();
                var decryptedBytes = decryptor.TransformFinalBlock(encryptedBytes, 0, encryptedBytes.Length);

                // CHỖ QUAN TRỌNG: Dùng GetString bình thường
                return Encoding.UTF8.GetString(decryptedBytes);
            }
            catch
            {
                throw new CryptographicException("Không thể đọc file license. Có thể bị hỏng hoặc sai key.");
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