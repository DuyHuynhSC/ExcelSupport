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
    }
}
