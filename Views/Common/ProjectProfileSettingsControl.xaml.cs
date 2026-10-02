using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using ExcelSupport.Helpers;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;
using WpfMessageBoxResult = System.Windows.MessageBoxResult;
using MediaColor = System.Windows.Media.Color;
using MediaBrush = System.Windows.Media.Brush;
using WpfRadioButton = System.Windows.Controls.RadioButton;
using WpfButton = System.Windows.Controls.Button;
using WpfUserControl = System.Windows.Controls.UserControl;

namespace ExcelSupport.Views
{
    public partial class ProjectProfileSettingsControl : WpfUserControl
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(ProjectProfileSettingsControl),
                new PropertyMetadata(false, OnIsDarkThemeChanged));

        private static void OnIsDarkThemeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ProjectProfileSettingsControl ctrl && ctrl._selectedProfile != null)
            {
                ctrl.UpdateActiveStatusBanner(ctrl._selectedProfile);
                ctrl.UpdateFolderStatusLabels();
            }
        }

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        private MediaBrush SuccessBrush => IsDarkTheme
            ? new SolidColorBrush(MediaColor.FromRgb(74, 222, 128)) // #4ADE80
            : new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));  // #16A34A

        private MediaBrush ErrorBrush => IsDarkTheme
            ? new SolidColorBrush(MediaColor.FromRgb(248, 113, 113)) // #F87171
            : new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));  // #DC2626

        private MediaBrush MutedBrush => IsDarkTheme
            ? new SolidColorBrush(MediaColor.FromRgb(148, 163, 184)) // #94A3B8
            : new SolidColorBrush(MediaColor.FromRgb(100, 116, 139)); // #64748B

        private MediaBrush InfoBrush => IsDarkTheme
            ? new SolidColorBrush(MediaColor.FromRgb(96, 165, 250))  // #60A5FA
            : new SolidColorBrush(MediaColor.FromRgb(37, 99, 235));  // #2563EB

        private readonly ObservableCollection<ProjectProfile> _profiles = new ObservableCollection<ProjectProfile>();
        private readonly ObservableCollection<OracleUserCredential> _settingUsers = new ObservableCollection<OracleUserCredential>();
        private ProjectProfile? _selectedProfile;
        private string? _activeProfileId;
        private bool _isBinding = false;

        public ProjectProfileSettingsControl()
        {
            InitializeComponent();
            lstProfiles.ItemsSource = _profiles;
            dgSettingUsers.ItemsSource = _settingUsers;
            Loaded += (s, e) => ReloadProfiles();
        }

        public void ReloadProfiles(string? selectProfileId = null)
        {
            _isBinding = true;
            try
            {
                var config = ProjectProfileManager.CurrentConfig;
                _activeProfileId = config.ActiveProfileId;

                string? targetId = selectProfileId ?? (lstProfiles.SelectedItem as ProjectProfile)?.Id ?? _activeProfileId;

                _profiles.Clear();
                foreach (var p in config.Profiles)
                {
                    // Fallback if name was somehow empty
                    if (string.IsNullOrWhiteSpace(p.Name))
                    {
                        p.Name = $"{LocalizationService.Get("SpecProfile_NewProjectPrefix", "Dự Án Mới")} {_profiles.Count + 1}";
                    }
                    p.IsActive = (p.Id == _activeProfileId);
                    p.DatabaseConnection.EnsureDefaultUsers();
                    p.DatabaseConnection.Id = p.Id;
                    p.DatabaseConnection.Name = p.Name;
                    p.DatabaseConnection.IsDefault = (p.Id == _activeProfileId);
                    _profiles.Add(p);
                }

                _selectedProfile = _profiles.FirstOrDefault(p => p.Id == targetId) ?? _profiles.FirstOrDefault();
                lstProfiles.SelectedItem = _selectedProfile;
                BindProfileToForm(_selectedProfile);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProjectProfileSettingsControl.ReloadProfiles] Error: {ex.Message}");
            }
            finally
            {
                _isBinding = false;
            }
        }

        private void LstProfiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isBinding) return;

            // Only save previously selected profile if it had a valid name and was edited
            if (e.RemovedItems.Count > 0 && e.RemovedItems[0] is ProjectProfile prevProfile && _profiles.Contains(prevProfile))
            {
                if (!string.IsNullOrWhiteSpace(txtName.Text))
                {
                    SaveFormToModel(prevProfile);
                }
            }

            _selectedProfile = lstProfiles.SelectedItem as ProjectProfile;
            BindProfileToForm(_selectedProfile);
        }

        private void TxtName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isBinding || _selectedProfile == null) return;
            string newName = txtName.Text;
            _selectedProfile.Name = newName;
            _selectedProfile.DatabaseConnection.Name = newName;
        }

        private void BindProfileToForm(ProjectProfile? p)
        {
            if (p == null) return;

            _isBinding = true;
            try
            {
                // 1. Documents Tab
                txtName.Text = p.Name;
                txtRootFolder.Text = p.RootFolder;
                txtTkctViFolder.Text = p.DetailedDesignFolderVi;
                txtTkctJaFolder.Text = p.DetailedDesignFolderJa;
                txtTkcbViFolder.Text = p.BasicDesignFolderVi;
                txtTkcbJaFolder.Text = p.BasicDesignFolderJa;
                txtTestSpecFolder.Text = p.TestSpecFolder;
                chkDoNotSearchSubfolders.IsChecked = p.DoNotSearchSubfolders;
                chkOpenReadOnly.IsChecked = p.OpenReadOnlyDefault;
                txtFileExtensions.Text = p.FileExtensions;

                // 2. Database Tab
                var db = p.DatabaseConnection;
                db.EnsureDefaultUsers();
                txtSettingHost.Text = db.Host;
                txtSettingPort.Text = db.Port.ToString();
                txtSettingService.Text = db.ServiceNameOrSid;
                rbSettingService.IsChecked = db.ServiceType == OracleServiceNameType.ServiceName;
                rbSettingSid.IsChecked = db.ServiceType == OracleServiceNameType.SID;
                txtSettingDbStatus.Text = string.Empty;

                _settingUsers.Clear();
                foreach (var u in db.Users)
                {
                    _settingUsers.Add(u);
                }

                var defUser = _settingUsers.FirstOrDefault(u => u.IsDefault) ?? _settingUsers.FirstOrDefault();
                dgSettingUsers.SelectedItem = defUser;

                // 3. Active Status Banner
                UpdateActiveStatusBanner(p);
                UpdateFolderStatusLabels();
                txtStatusNote.Text = string.Empty;
            }
            finally
            {
                _isBinding = false;
            }
        }

        private void SaveFormToModel(ProjectProfile p)
        {
            if (p == null) return;

            // 1. Documents
            p.Name = txtName.Text.Trim();
            p.RootFolder = txtRootFolder.Text.Trim();
            p.DetailedDesignFolderVi = txtTkctViFolder.Text.Trim();
            p.DetailedDesignFolderJa = txtTkctJaFolder.Text.Trim();
            p.BasicDesignFolderVi = txtTkcbViFolder.Text.Trim();
            p.BasicDesignFolderJa = txtTkcbJaFolder.Text.Trim();
            p.TestSpecFolder = txtTestSpecFolder.Text.Trim();
            p.DoNotSearchSubfolders = chkDoNotSearchSubfolders.IsChecked == true;
            p.OpenReadOnlyDefault = chkOpenReadOnly.IsChecked == true;
            p.FileExtensions = txtFileExtensions.Text.Trim();
            p.LastModified = DateTime.Now;

            // 2. Database
            var db = p.DatabaseConnection;
            db.Name = p.Name;
            db.Host = txtSettingHost.Text.Trim();
            if (int.TryParse(txtSettingPort.Text.Trim(), out int port))
            {
                db.Port = port;
            }
            db.ServiceNameOrSid = txtSettingService.Text.Trim();
            db.ServiceType = rbSettingSid.IsChecked == true ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName;
            db.Users = _settingUsers.ToList();

            var defUser = _settingUsers.FirstOrDefault(u => u.IsDefault);
            if (defUser != null)
            {
                db.SelectedUserId = defUser.Id;
            }
        }

        private void UpdateActiveStatusBanner(ProjectProfile p)
        {
            bool isActive = (p.Id == _activeProfileId);
            p.IsActive = isActive;
            p.DatabaseConnection.IsDefault = isActive;

            if (isActive)
            {
                lblActiveStatus.Text = LocalizationService.Get("SpecProfile_StatusActive", "Đang kích hoạt (Active)");
                lblActiveStatus.Foreground = SuccessBrush;
                iconActiveStatus.Text = "🟢";
                btnSetActive.IsEnabled = false;
                btnSetActive.Content = LocalizationService.Get("SpecProfile_BtnIsActive", "✓ Đang Kích Hoạt");
            }
            else
            {
                lblActiveStatus.Text = LocalizationService.Get("SpecProfile_StatusInactive", "Không kích hoạt");
                lblActiveStatus.Foreground = MutedBrush;
                iconActiveStatus.Text = "⚪";
                btnSetActive.IsEnabled = true;
                btnSetActive.Content = LocalizationService.Get("SpecProfile_BtnSetActive", "⭐ Đặt Làm Dự Án Đang Kích Hoạt");
            }
        }

        private void UpdateFolderStatusLabels()
        {
            if (_selectedProfile == null) return;
            string root = txtRootFolder.Text.Trim();

            UpdateSingleLabel(lblTkctViStatus, txtTkctViFolder.Text.Trim(), root);
            UpdateSingleLabel(lblTkctJaStatus, txtTkctJaFolder.Text.Trim(), root);
            UpdateSingleLabel(lblTkcbViStatus, txtTkcbViFolder.Text.Trim(), root);
            UpdateSingleLabel(lblTkcbJaStatus, txtTkcbJaFolder.Text.Trim(), root);
            UpdateSingleLabel(lblTestSpecStatus, txtTestSpecFolder.Text.Trim(), root);
        }

        private void UpdateSingleLabel(TextBlock lbl, string folderPath, string root)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                lbl.Text = LocalizationService.Get("SpecProfile_StatusNotConfigured", "— Chưa cấu hình");
                lbl.Foreground = MutedBrush;
                return;
            }

            string fullPath = Path.IsPathRooted(folderPath) ? folderPath : Path.Combine(root, folderPath);
            if (Directory.Exists(fullPath))
            {
                lbl.Text = LocalizationService.Get("SpecProfile_StatusValid", "✓ Hợp lệ");
                lbl.Foreground = SuccessBrush;
            }
            else
            {
                lbl.Text = LocalizationService.Get("SpecProfile_StatusNotFound", "⚠ Không tìm thấy");
                lbl.Foreground = ErrorBrush;
            }
        }

        private void OnFolderFieldChanged(object sender, TextChangedEventArgs e)
        {
            if (_isBinding) return;
            if (_selectedProfile != null && sender == txtRootFolder)
            {
                _selectedProfile.RootFolder = txtRootFolder.Text.Trim();
            }
            UpdateFolderStatusLabels();
        }

        #region Master Project Actions (Add, Clone, Delete, SetActive, Save)

        private void BtnAddProject_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile != null && !string.IsNullOrWhiteSpace(txtName.Text))
            {
                SaveFormToModel(_selectedProfile);
            }

            var newProf = new ProjectProfile
            {
                Name = $"{LocalizationService.Get("SpecProfile_NewProjectPrefix", "Dự Án Mới")} {_profiles.Count + 1}",
                RootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                DetailedDesignFolderVi = "Detailed_Design/VI",
                DetailedDesignFolderJa = "Detailed_Design/JA",
                BasicDesignFolderVi = "Basic_Design/VI",
                BasicDesignFolderJa = "Basic_Design/JA",
                TestSpecFolder = "Test_Specification",
                DoNotSearchSubfolders = false,
                FileExtensions = ".xlsx;.xlsm;.xls;.docx;.pdf;.pptx",
                OpenReadOnlyDefault = false
            };
            newProf.DatabaseConnection.Name = newProf.Name;

            _profiles.Add(newProf);
            if (string.IsNullOrEmpty(_activeProfileId))
            {
                _activeProfileId = newProf.Id;
            }

            lstProfiles.SelectedItem = newProf;
            txtName.Focus();
            txtName.SelectAll();
        }

        private void BtnCloneProject_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;
            SaveFormToModel(_selectedProfile);

            var cloned = _selectedProfile.Clone();
            _profiles.Add(cloned);
            lstProfiles.SelectedItem = cloned;
        }

        private void BtnDeleteProject_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;

            var win = Window.GetWindow(this);
            if (_profiles.Count <= 1)
            {
                WpfMessageBox.Show(
                    win,
                    LocalizationService.Get("SpecProfile_MinProfilePrompt", "Cần duy trì ít nhất 1 Profile dự án trong hệ thống."),
                    LocalizationService.Get("Common_Notice", "Thông Báo"),
                    WpfMessageBoxButton.OK,
                    WpfMessageBoxImage.Warning);
                win?.Activate();
                return;
            }

            var result = WpfMessageBox.Show(
                win,
                LocalizationService.Get("SpecProfile_DeleteConfirmPrompt", _selectedProfile.Name),
                LocalizationService.Get("SpecProfile_DeleteConfirmTitle", "Xác Nhận Xóa"),
                WpfMessageBoxButton.YesNo,
                WpfMessageBoxImage.Question);
            win?.Activate();

            if (result == WpfMessageBoxResult.Yes)
            {
                string deletedId = _selectedProfile.Id;
                _profiles.Remove(_selectedProfile);

                if (_activeProfileId == deletedId)
                {
                    _activeProfileId = _profiles.FirstOrDefault()?.Id;
                }

                _selectedProfile = _profiles.FirstOrDefault();
                lstProfiles.SelectedItem = _selectedProfile;
                BindProfileToForm(_selectedProfile);
            }
        }

        private void BtnSetActive_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;
            SaveFormToModel(_selectedProfile);

            _activeProfileId = _selectedProfile.Id;
            foreach (var p in _profiles)
            {
                p.IsActive = (p.Id == _activeProfileId);
                p.DatabaseConnection.IsDefault = (p.Id == _activeProfileId);
            }

            // Force refresh UI list item templates
            lstProfiles.Items.Refresh();
            UpdateActiveStatusBanner(_selectedProfile);

            // Persist active profile
            ProjectProfileManager.SetActiveProfile(_activeProfileId);

            txtStatusNote.Text = LocalizationService.Get("SpecProfile_MsgActiveChanged", _selectedProfile.Name);
            txtStatusNote.Foreground = SuccessBrush;
        }

        private void BtnSaveProject_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_selectedProfile != null)
                {
                    SaveFormToModel(_selectedProfile);
                }

                // Save all profiles to ProjectProfileManager (which syncs to OracleConnectionManager)
                ProjectProfileManager.SaveAllProfiles(_profiles.ToList(), _activeProfileId);

                txtStatusNote.Text = LocalizationService.Get("SpecProfile_SaveSuccess", "✓ Đã lưu thành công cấu hình Profile dự án!");
                txtStatusNote.Foreground = SuccessBrush;
                lstProfiles.Items.Refresh();
            }
            catch (Exception ex)
            {
                txtStatusNote.Text = ex.Message;
                txtStatusNote.Foreground = ErrorBrush;
            }
        }

        #endregion

        #region Folder Browsing Handlers

        private void OnBrowseRootFolderClick(object sender, RoutedEventArgs e)
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = LocalizationService.Get("SpecProfile_BrowseRootDesc", "Chọn thư mục gốc của dự án:"),
                SelectedPath = Directory.Exists(txtRootFolder.Text) ? txtRootFolder.Text : string.Empty
            };
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtRootFolder.Text = fbd.SelectedPath;
            }
        }

        private void OnBrowseTkctViFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseFolderForField(txtTkctViFolder, LocalizationService.Get("SpecProfile_BrowseTkctViDesc", "Chọn thư mục TKCT (Tiếng Việt):"));
        }

        private void OnBrowseTkctJaFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseFolderForField(txtTkctJaFolder, LocalizationService.Get("SpecProfile_BrowseTkctJaDesc", "Chọn thư mục TKCT (Tiếng Nhật):"));
        }

        private void OnBrowseTkcbViFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseFolderForField(txtTkcbViFolder, LocalizationService.Get("SpecProfile_BrowseTkcbViDesc", "Chọn thư mục TKCB (Tiếng Việt):"));
        }

        private void OnBrowseTkcbJaFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseFolderForField(txtTkcbJaFolder, LocalizationService.Get("SpecProfile_BrowseTkcbJaDesc", "Chọn thư mục TKCB (Tiếng Nhật):"));
        }

        private void OnBrowseTestSpecFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseFolderForField(txtTestSpecFolder, LocalizationService.Get("SpecProfile_BrowseTestDesc", "Chọn thư mục Chỉ Thị Test (Test Specification):"));
        }

        private void BrowseFolderForField(System.Windows.Controls.TextBox targetBox, string description)
        {
            string root = txtRootFolder.Text.Trim();
            string initialPath = targetBox.Text.Trim();

            if (!Path.IsPathRooted(initialPath) && !string.IsNullOrWhiteSpace(root) && Directory.Exists(root))
            {
                string combined = Path.Combine(root, initialPath);
                if (Directory.Exists(combined))
                {
                    initialPath = combined;
                }
                else
                {
                    initialPath = root;
                }
            }

            using var fbd = new FolderBrowserDialog
            {
                Description = description,
                SelectedPath = Directory.Exists(initialPath) ? initialPath : (Directory.Exists(root) ? root : string.Empty)
            };

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string chosen = fbd.SelectedPath;
                // If subfolder of root, store relative path for cleaner portable config
                if (!string.IsNullOrWhiteSpace(root) && chosen.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                {
                    string rel = chosen.Substring(root.Length).TrimStart('\\', '/');
                    targetBox.Text = string.IsNullOrEmpty(rel) ? "." : rel;
                }
                else
                {
                    targetBox.Text = chosen;
                }
            }
        }

        #endregion

        #region Database User & Test Handlers

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
                txtSettingDbStatus.Text = LocalizationService.Get("Oracle_MsgUserAdded", newUser.Username);
                txtSettingDbStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(22, 163, 74));
            }
            catch (Exception ex)
            {
                txtSettingDbStatus.Text = ex.Message;
                txtSettingDbStatus.Foreground = new SolidColorBrush(MediaColor.FromRgb(220, 38, 38));
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
                dgSettingUsers.Items.Refresh();
            }
        }

        private void BtnRowDeleteUser_Click(object sender, RoutedEventArgs e)
        {
            if (sender is WpfButton btn && btn.DataContext is OracleUserCredential userToDelete)
            {
                if (_settingUsers.Count <= 1)
                {
                    WpfMessageBox.Show(
                        Window.GetWindow(this),
                        LocalizationService.Get("Oracle_MsgMinUserRequired", "Mỗi profile cần ít nhất 1 tài khoản kết nối."),
                        LocalizationService.Get("Common_Notice", "Thông Báo"),
                        WpfMessageBoxButton.OK,
                        WpfMessageBoxImage.Warning);
                    return;
                }

                _settingUsers.Remove(userToDelete);
                if (userToDelete.IsDefault && _settingUsers.Count > 0)
                {
                    _settingUsers[0].IsDefault = true;
                }
                dgSettingUsers.Items.Refresh();
            }
        }

        private async void BtnSettingTest_Click(object sender, RoutedEventArgs e)
        {
            var win = Window.GetWindow(this);
            try
            {
                string host = txtSettingHost.Text.Trim();
                if (string.IsNullOrEmpty(host))
                {
                    WpfMessageBox.Show(win, LocalizationService.Get("Oracle_MsgEnterHost", "Vui lòng nhập Host của Oracle Server."), LocalizationService.Get("Common_Notice", "Thông Báo"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    win?.Activate();
                    return;
                }

                if (!int.TryParse(txtSettingPort.Text.Trim(), out int port))
                {
                    port = 1521;
                }

                string service = txtSettingService.Text.Trim();
                if (string.IsNullOrEmpty(service))
                {
                    WpfMessageBox.Show(win, LocalizationService.Get("Oracle_MsgEnterService", "Vui lòng nhập Service Name hoặc SID."), LocalizationService.Get("Common_Notice", "Thông Báo"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    win?.Activate();
                    return;
                }

                var user = dgSettingUsers.SelectedItem as OracleUserCredential ?? _settingUsers.FirstOrDefault(u => u.IsDefault) ?? _settingUsers.FirstOrDefault();
                if (user == null || string.IsNullOrWhiteSpace(user.Username))
                {
                    WpfMessageBox.Show(win, LocalizationService.Get("Oracle_MsgEnterUser", "Vui lòng nhập Username để kiểm tra kết nối."), LocalizationService.Get("Common_Notice", "Thông Báo"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    win?.Activate();
                    return;
                }

                var config = new OracleConnectionConfig
                {
                    Host = host,
                    Port = port,
                    ServiceNameOrSid = service,
                    ServiceType = rbSettingSid.IsChecked == true ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName,
                    Username = user.Username,
                    Password = user.Password
                };

                btnSettingTest.IsEnabled = false;
                txtSettingDbStatus.Text = LocalizationService.Get("Oracle_MsgConnecting", "Đang thử kết nối...");
                txtSettingDbStatus.Foreground = InfoBrush;

                var (success, msg, version) = await OracleDataCompareService.TestConnectionAsync(config);

                if (success)
                {
                    txtSettingDbStatus.Text = $"✓ Kết nối thành công! {version}";
                    txtSettingDbStatus.Foreground = SuccessBrush;
                    string msgSuccess = LocalizationService.Get("Oracle_MsgConnectedUser", user.Username) + $"\n{version}";
                    WpfMessageBox.Show(win, msgSuccess, LocalizationService.Get("Common_Success", "Thành Công"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                    win?.Activate();
                }
                else
                {
                    txtSettingDbStatus.Text = LocalizationService.Get("Oracle_MsgConnectFailed", "✗ Kết nối thất bại!");
                    txtSettingDbStatus.Foreground = ErrorBrush;
                    WpfMessageBox.Show(win, msg, LocalizationService.Get("Oracle_TitleConnectError", "Lỗi Kết Nối Oracle"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                    win?.Activate();
                }
            }
            catch (Exception ex)
            {
                txtSettingDbStatus.Text = "✗ " + ex.Message;
                txtSettingDbStatus.Foreground = ErrorBrush;
                WpfMessageBox.Show(win, ex.Message, LocalizationService.Get("Oracle_TitleConnectError", "Lỗi Kết Nối Oracle"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                win?.Activate();
            }
            finally
            {
                btnSettingTest.IsEnabled = true;
            }
        }

        private async void BtnSettingTestAll_Click(object sender, RoutedEventArgs e)
        {
            var win = Window.GetWindow(this);
            try
            {
                string host = txtSettingHost.Text.Trim();
                if (string.IsNullOrEmpty(host))
                {
                    WpfMessageBox.Show(win, LocalizationService.Get("Oracle_MsgEnterHost", "Vui lòng nhập Host của Oracle Server."), LocalizationService.Get("Common_Notice", "Thông Báo"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    win?.Activate();
                    return;
                }

                if (!int.TryParse(txtSettingPort.Text.Trim(), out int port))
                {
                    port = 1521;
                }

                string service = txtSettingService.Text.Trim();
                if (string.IsNullOrEmpty(service))
                {
                    WpfMessageBox.Show(win, LocalizationService.Get("Oracle_MsgEnterService", "Vui lòng nhập Service Name hoặc SID."), LocalizationService.Get("Common_Notice", "Thông Báo"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    win?.Activate();
                    return;
                }

                var usersToTest = _settingUsers.Where(u => !string.IsNullOrWhiteSpace(u.Username)).ToList();
                if (usersToTest.Count == 0)
                {
                    WpfMessageBox.Show(win, LocalizationService.Get("Oracle_MsgNoUsersToTest", "Chưa có tài khoản kết nối nào được nhập để kiểm tra."), LocalizationService.Get("Common_Notice", "Thông Báo"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    win?.Activate();
                    return;
                }

                btnSettingTestAll.IsEnabled = false;
                btnSettingTest.IsEnabled = false;
                txtSettingDbStatus.Text = LocalizationService.Get("Oracle_MsgTestingAllConn", usersToTest.Count);
                txtSettingDbStatus.Foreground = InfoBrush;

                var serviceType = rbSettingSid.IsChecked == true ? OracleServiceNameType.SID : OracleServiceNameType.ServiceName;

                var tasks = usersToTest.Select(async u =>
                {
                    var cfg = new OracleConnectionConfig
                    {
                        Host = host,
                        Port = port,
                        ServiceNameOrSid = service,
                        ServiceType = serviceType,
                        Username = u.Username,
                        Password = u.Password
                    };
                    var (ok, message, ver) = await OracleDataCompareService.TestConnectionAsync(cfg);
                    return new { User = u, Success = ok, Message = message, Version = ver };
                });

                var testResults = await Task.WhenAll(tasks);

                var failed = testResults.Where(r => !r.Success).ToList();
                var succeeded = testResults.Where(r => r.Success).ToList();

                var sb = new StringBuilder();

                if (failed.Count == 0)
                {
                    txtSettingDbStatus.Text = LocalizationService.Get("Oracle_MsgTestAllSuccess", testResults.Length, testResults.Length);
                    txtSettingDbStatus.Foreground = SuccessBrush;

                    sb.AppendLine(LocalizationService.Get("Oracle_MsgAllConnSuccessHeader", testResults.Length));
                    sb.AppendLine($"Host: {host}:{port} ({service})");
                    sb.AppendLine();
                    foreach (var s in succeeded)
                    {
                        string role = !string.IsNullOrWhiteSpace(s.User.RoleOrDescription) ? $" ({s.User.RoleOrDescription})" : "";
                        sb.AppendLine($"  ✓ {s.User.Username}{role} — {s.Version}");
                    }

                    WpfMessageBox.Show(win, sb.ToString(), LocalizationService.Get("Oracle_TitleTestAllResults", "Kết quả kiểm tra kết nối"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
                    win?.Activate();
                }
                else
                {
                    txtSettingDbStatus.Text = LocalizationService.Get("Oracle_MsgTestAllFailed", failed.Count, testResults.Length);
                    txtSettingDbStatus.Foreground = ErrorBrush;

                    string profName = _selectedProfile?.Name ?? "";
                    sb.AppendLine(LocalizationService.Get("Oracle_MsgTestAllFailedHeader", failed.Count, testResults.Length, profName));
                    sb.AppendLine($"Host: {host}:{port} ({service})");
                    sb.AppendLine();
                    foreach (var f in failed)
                    {
                        string role = !string.IsNullOrWhiteSpace(f.User.RoleOrDescription) ? $" ({f.User.RoleOrDescription})" : "";
                        sb.AppendLine($"❌ [Tài khoản: {f.User.Username}{role}]");
                        sb.AppendLine($"   → Lỗi: {f.Message}");
                        sb.AppendLine();
                    }

                    if (succeeded.Count > 0)
                    {
                        sb.AppendLine("────────────────────────");
                        sb.AppendLine(LocalizationService.Get("Oracle_MsgSomeConnSuccessHeader", succeeded.Count));
                        foreach (var s in succeeded)
                        {
                            string role = !string.IsNullOrWhiteSpace(s.User.RoleOrDescription) ? $" ({s.User.RoleOrDescription})" : "";
                            sb.AppendLine($"  ✓ {s.User.Username}{role} — {s.Version}");
                        }
                    }

                    WpfMessageBox.Show(win, sb.ToString(), LocalizationService.Get("Oracle_TitleTestAllResults", "Kết quả kiểm tra kết nối"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Warning);
                    win?.Activate();
                }
            }
            catch (Exception ex)
            {
                txtSettingDbStatus.Text = "✗ " + ex.Message;
                txtSettingDbStatus.Foreground = ErrorBrush;
                WpfMessageBox.Show(win, ex.Message, LocalizationService.Get("Oracle_TitleConnectError", "Lỗi Kết Nối Oracle"), WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
                win?.Activate();
            }
            finally
            {
                btnSettingTestAll.IsEnabled = true;
                btnSettingTest.IsEnabled = true;
            }
        }

        #endregion
    }
}
