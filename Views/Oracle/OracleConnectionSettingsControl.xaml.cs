using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using WpfMessageBoxResult = System.Windows.MessageBoxResult;
using MediaColor = System.Windows.Media.Color;
using WpfRadioButton = System.Windows.Controls.RadioButton;
using WpfButton = System.Windows.Controls.Button;
using WpfUserControl = System.Windows.Controls.UserControl;

namespace ExcelSupport.Views
{
    public partial class OracleConnectionSettingsControl : WpfUserControl
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(OracleConnectionSettingsControl),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        public event EventHandler? ProfilesChanged;

        private readonly ObservableCollection<OracleConnectionProfile> _profiles = new ObservableCollection<OracleConnectionProfile>();
        private readonly ObservableCollection<OracleUserCredential> _settingUsers = new ObservableCollection<OracleUserCredential>();

        public OracleConnectionSettingsControl()
        {
            InitializeComponent();
            lstProfiles.ItemsSource = _profiles;
            dgSettingUsers.ItemsSource = _settingUsers;
        }

        public void ReloadProfiles(string? selectProfileId = null)
        {
            try
            {
                string? targetId = selectProfileId ?? (lstProfiles.SelectedItem as OracleConnectionProfile)?.Id;
                var list = OracleConnectionManager.GetProfiles();
                _profiles.Clear();
                foreach (var p in list)
                {
                    _profiles.Add(p);
                }

                lstProfiles.SelectedItem = _profiles.FirstOrDefault(p => p.Id == targetId) ?? _profiles.FirstOrDefault();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OracleConnectionSettingsControl.ReloadProfiles] Error: {ex.Message}");
            }
        }

        private void LstProfiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                txtSettingProfileName.Text = p.Name;
                txtSettingHost.Text = p.Host;
                txtSettingPort.Text = p.Port.ToString();
                txtSettingService.Text = p.ServiceNameOrSid;
                rbSettingService.IsChecked = p.ServiceType == OracleServiceNameType.ServiceName;
                rbSettingSid.IsChecked = p.ServiceType == OracleServiceNameType.SID;
                txtSettingStatus.Text = string.Empty;

                p.EnsureDefaultUsers();
                _settingUsers.Clear();
                foreach (var u in p.Users)
                {
                    _settingUsers.Add(u);
                }

                var defUser = _settingUsers.FirstOrDefault(u => u.IsDefault) ?? _settingUsers.FirstOrDefault();
                dgSettingUsers.SelectedItem = defUser;
            }
        }

        private void BtnUserAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var newUser = new OracleUserCredential
                {
                    Username = $"USER_{_settingUsers.Count + 1}",
                    Password = "",
                    RoleOrDescription = "",
                    IsDefault = _settingUsers.Count == 0
                };
                _settingUsers.Add(newUser);
                dgSettingUsers.SelectedItem = newUser;
                txtSettingStatus.Text = string.Format(LocalizationService.Get("Oracle_MsgUserAdded") ?? "Đã thêm tài khoản '{0}' vào danh sách.", newUser.Username);
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
            catch (Exception ex)
            {
                txtSettingStatus.Text = ex.Message;
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
            }
        }

        private void RbUserDefault_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfRadioButton rb && rb.DataContext is OracleUserCredential selected)
            {
                dgSettingUsers.SelectedItem = selected;
                foreach (var u in _settingUsers)
                {
                    u.IsDefault = (u == selected);
                }
                txtSettingStatus.Text = string.Format(LocalizationService.Get("Oracle_MsgUserDefaultSet") ?? "Đã đặt User '{0}' làm mặc định.", selected.Username);
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
        }

        private void BtnRowDeleteUser_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.DataContext is OracleUserCredential item)
            {
                var parentWindow = Window.GetWindow(this);
                if (_settingUsers.Count <= 1)
                {
                    WpfMessageBox.Show(parentWindow,
                        LocalizationService.Get("Oracle_MsgMinUserRequired") ?? "Mỗi cấu hình kết nối phải có ít nhất 1 tài khoản User.",
                        LocalizationService.Get("Common_Warning") ?? "Cảnh báo",
                        WpfMessageBoxButton.OK,
                        WpfMessageBoxImage.Warning);
                    return;
                }

                _settingUsers.Remove(item);
                if (!_settingUsers.Any(u => u.IsDefault) && _settingUsers.Count > 0)
                {
                    _settingUsers[0].IsDefault = true;
                }
                txtSettingStatus.Text = string.Format(LocalizationService.Get("Oracle_MsgUserDeleted") ?? "Đã xóa User '{0}'.", item.Username);
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
        }

        private void BtnNewProfile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var newProfile = new OracleConnectionProfile
                {
                    Name = $"Connection {_profiles.Count + 1}",
                    Host = "localhost",
                    Port = 1521,
                    ServiceNameOrSid = "ORCL",
                    ServiceType = OracleServiceNameType.ServiceName
                };
                newProfile.EnsureDefaultUsers();

                OracleConnectionManager.AddOrUpdateProfile(newProfile);
                ReloadProfiles(newProfile.Id);
                ProfilesChanged?.Invoke(this, EventArgs.Empty);

                txtSettingStatus.Text = $"Đã tạo mới cấu hình '{newProfile.Name}'.";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
            catch (Exception ex)
            {
                txtSettingStatus.Text = ex.Message;
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
            }
        }

        private void BtnSetDefaultProfile_Click(object sender, RoutedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                OracleConnectionManager.SetDefaultProfile(p.Id);
                ReloadProfiles(p.Id);
                ProfilesChanged?.Invoke(this, EventArgs.Empty);
                txtSettingStatus.Text = $"Đã đặt '{p.Name}' làm kết nối mặc định.";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
            else
            {
                var parentWindow = Window.GetWindow(this);
                WpfMessageBox.Show(parentWindow,
                    LocalizationService.Get("Oracle_MsgSelectProfileFirst") ?? "Vui lòng chọn một cấu hình kết nối trong danh sách để đặt làm mặc định.",
                    LocalizationService.Get("Common_Notice") ?? "Chưa chọn kết nối",
                    WpfMessageBoxButton.OK,
                    WpfMessageBoxImage.Information);
            }
        }

        private void BtnCloneProfile_Click(object sender, RoutedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                var clone = p.Clone();
                OracleConnectionManager.AddOrUpdateProfile(clone);
                ReloadProfiles(clone.Id);
                ProfilesChanged?.Invoke(this, EventArgs.Empty);
                txtSettingStatus.Text = $"Đã nhân bản cấu hình '{p.Name}'.";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
        }

        private void BtnDeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                var parentWindow = Window.GetWindow(this);
                var confirm = WpfMessageBox.Show(parentWindow,
                    $"Bạn có chắc chắn muốn xóa cấu hình kết nối '{p.Name}' không?",
                    LocalizationService.Get("Common_Confirm") ?? "Xác nhận xóa",
                    WpfMessageBoxButton.YesNo,
                    WpfMessageBoxImage.Question);

                if (confirm == WpfMessageBoxResult.Yes)
                {
                    OracleConnectionManager.DeleteProfile(p.Id);
                    ReloadProfiles();
                    ProfilesChanged?.Invoke(this, EventArgs.Empty);
                    txtSettingStatus.Text = $"Đã xóa cấu hình '{p.Name}'.";
                    txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
                }
            }
        }

        private void BtnSettingSave_Click(object sender, RoutedEventArgs e)
        {
            if (lstProfiles.SelectedItem is OracleConnectionProfile p)
            {
                int.TryParse(txtSettingPort.Text, out int port);
                if (port <= 0) port = 1521;

                p.Name = string.IsNullOrWhiteSpace(txtSettingProfileName.Text) ? "Connection" : txtSettingProfileName.Text.Trim();
                p.Host = txtSettingHost.Text.Trim();
                p.Port = port;
                p.ServiceNameOrSid = txtSettingService.Text.Trim();
                p.ServiceType = (rbSettingSid.IsChecked == true) ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName;

                p.Users = _settingUsers.ToList();
                if (!p.Users.Any(u => u.IsDefault) && p.Users.Count > 0)
                {
                    p.Users[0].IsDefault = true;
                }
                var def = p.GetEffectiveUser();
                p.Username = def?.Username ?? "";
                p.Password = def?.Password ?? "";

                OracleConnectionManager.AddOrUpdateProfile(p);
                ReloadProfiles(p.Id);
                ProfilesChanged?.Invoke(this, EventArgs.Empty);

                txtSettingStatus.Text = "✅ Đã lưu cấu hình thành công!";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
        }

        private async void BtnSettingTest_Click(object sender, RoutedEventArgs e)
        {
            btnSettingTest.IsEnabled = false;

            int.TryParse(txtSettingPort.Text, out int port);
            if (port <= 0) port = 1521;

            var activeUser = (dgSettingUsers.SelectedItem as OracleUserCredential)
                             ?? _settingUsers.FirstOrDefault(u => u.IsDefault)
                             ?? _settingUsers.FirstOrDefault();

            string testUsername = activeUser?.Username?.Trim() ?? "";
            string testPassword = activeUser?.Password ?? "";

            txtSettingStatus.Text = string.Format(LocalizationService.Get("Oracle_MsgTestingConnUser") ?? "⏳ Đang kiểm tra kết nối với User '{0}'...", testUsername);
            txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(217, 119, 6));

            var config = new OracleConnectionConfig
            {
                Host = txtSettingHost.Text.Trim(),
                Port = port,
                ServiceNameOrSid = txtSettingService.Text.Trim(),
                ServiceType = (rbSettingSid.IsChecked == true) ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName,
                Username = testUsername,
                Password = testPassword
            };

            var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(config);
            btnSettingTest.IsEnabled = true;

            if (success)
            {
                txtSettingStatus.Text = $"✅ Kết nối thành công! {version}";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
            else
            {
                txtSettingStatus.Text = $"❌ Kết nối thất bại: {msg}";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
            }
        }

        private async void BtnSettingTestAll_Click(object sender, RoutedEventArgs e)
        {
            var parentWindow = Window.GetWindow(this);
            if (_settingUsers.Count == 0)
            {
                WpfMessageBox.Show(parentWindow,
                    LocalizationService.Get("Oracle_MsgNoUsersToTest") ?? "Chưa có tài khoản User nào trong cấu hình để kiểm tra.",
                    LocalizationService.Get("Common_Warning") ?? "Cảnh báo",
                    WpfMessageBoxButton.OK,
                    WpfMessageBoxImage.Warning);
                return;
            }

            int.TryParse(txtSettingPort.Text, out int port);
            if (port <= 0) port = 1521;

            string host = txtSettingHost.Text.Trim();
            string service = txtSettingService.Text.Trim();
            var serviceType = (rbSettingSid.IsChecked == true) ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName;
            string profileName = string.IsNullOrWhiteSpace(txtSettingProfileName.Text) ? "Profile" : txtSettingProfileName.Text.Trim();

            btnSettingTest.IsEnabled = false;
            btnSettingTestAll.IsEnabled = false;
            btnSettingSave.IsEnabled = false;

            txtSettingStatus.Text = string.Format(LocalizationService.Get("Oracle_MsgTestingAllConn") ?? "⏳ Đang kiểm tra kết nối cho toàn bộ {0} tài khoản User...", _settingUsers.Count);
            txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(217, 119, 6));

            try
            {
                var usersCopy = _settingUsers.ToList();
                var tasks = usersCopy.Select(async u =>
                {
                    var config = new OracleConnectionConfig
                    {
                        Host = host,
                        Port = port,
                        ServiceNameOrSid = service,
                        ServiceType = serviceType,
                        Username = u.Username?.Trim() ?? "",
                        Password = u.Password ?? ""
                    };
                    var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(config);
                    return (User: u, Success: success, Message: msg, Version: version);
                });

                var results = await System.Threading.Tasks.Task.WhenAll(tasks);
                var failed = results.Where(r => !r.Success).ToList();
                var passed = results.Where(r => r.Success).ToList();

                if (failed.Count == 0)
                {
                    txtSettingStatus.Text = string.Format(LocalizationService.Get("Oracle_MsgTestAllSuccess") ?? "✅ Tất cả {0}/{1} User kết nối thành công!", passed.Count, results.Length);
                    txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));

                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine(string.Format(LocalizationService.Get("Oracle_MsgAllConnSuccessHeader") ?? "Tất cả {0} tài khoản User đều kết nối thành công tới Database!", results.Length));
                    sb.AppendLine($"Host: {host}:{port} ({service})");
                    sb.AppendLine();
                    foreach (var r in results)
                    {
                        string roleDesc = string.IsNullOrWhiteSpace(r.User.RoleOrDescription) ? "" : $" - {r.User.RoleOrDescription}";
                        string defBadge = r.User.IsDefault ? " ⭐" : "";
                        sb.AppendLine($"✅ [{r.User.Username}{defBadge}]{roleDesc}: {r.Version}");
                    }

                    WpfMessageBox.Show(parentWindow,
                        sb.ToString(),
                        LocalizationService.Get("Oracle_TitleTestAllResults") ?? "Kết Quả Kiểm Tra Kết Nối",
                        WpfMessageBoxButton.OK,
                        WpfMessageBoxImage.Information);
                }
                else
                {
                    txtSettingStatus.Text = string.Format(LocalizationService.Get("Oracle_MsgTestAllFailed") ?? "❌ Có {0}/{1} kết nối thất bại! Xem chi tiết trong thông báo.", failed.Count, results.Length);
                    txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));

                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine(string.Format(LocalizationService.Get("Oracle_MsgTestAllFailedHeader") ?? "Phát hiện {0}/{1} tài khoản kết nối THẤT BẠI trong cấu hình '{2}':", failed.Count, results.Length, profileName));
                    sb.AppendLine($"Host: {host}:{port} ({service})");
                    sb.AppendLine();
                    sb.AppendLine("=== CÁC TÀI KHOẢN BỊ LỖI ===");
                    foreach (var r in failed)
                    {
                        string roleDesc = string.IsNullOrWhiteSpace(r.User.RoleOrDescription) ? "" : $" ({r.User.RoleOrDescription})";
                        string defBadge = r.User.IsDefault ? " ⭐" : "";
                        sb.AppendLine($"❌ [{r.User.Username}{defBadge}]{roleDesc}:");
                        sb.AppendLine($"   ➥ {r.Message}");
                        sb.AppendLine();
                    }

                    if (passed.Count > 0)
                    {
                        sb.AppendLine("=== CÁC TÀI KHOẢN THÀNH CÔNG ===");
                        foreach (var r in passed)
                        {
                            string roleDesc = string.IsNullOrWhiteSpace(r.User.RoleOrDescription) ? "" : $" ({r.User.RoleOrDescription})";
                            string defBadge = r.User.IsDefault ? " ⭐" : "";
                            sb.AppendLine($"✅ [{r.User.Username}{defBadge}]{roleDesc}: {r.Version}");
                        }
                    }

                    WpfMessageBox.Show(parentWindow,
                        sb.ToString(),
                        LocalizationService.Get("Oracle_TitleTestAllResults") ?? "Kết Quả Kiểm Tra Kết Nối",
                        WpfMessageBoxButton.OK,
                        WpfMessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                txtSettingStatus.Text = $"❌ Lỗi kiểm tra: {ex.Message}";
                txtSettingStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
            }
            finally
            {
                btnSettingTest.IsEnabled = true;
                btnSettingTestAll.IsEnabled = true;
                btnSettingSave.IsEnabled = true;
            }
        }
    }
}
