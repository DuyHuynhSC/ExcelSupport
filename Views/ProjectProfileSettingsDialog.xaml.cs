using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using ExcelSupport.Host;
using ExcelSupport.Models;
using ExcelSupport.Services;
using WpfMessageBox = System.Windows.MessageBox;
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;
using FormsDialogResult = System.Windows.Forms.DialogResult;
using WpfColor = System.Windows.Media.Color;

namespace ExcelSupport.Views
{
    public partial class ProjectProfileSettingsDialog : Window, INotifyPropertyChanged
    {
        public static readonly DependencyProperty IsDarkThemeProperty =
            DependencyProperty.Register(
                nameof(IsDarkTheme),
                typeof(bool),
                typeof(ProjectProfileSettingsDialog),
                new PropertyMetadata(false));

        public bool IsDarkTheme
        {
            get => (bool)GetValue(IsDarkThemeProperty);
            set => SetValue(IsDarkThemeProperty, value);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private static ProjectProfileSettingsDialog? _currentInstance;

        private List<ProjectProfile> _profiles = new List<ProjectProfile>();
        private string? _activeProfileId;
        private ProjectProfile? _selectedProfile;
        private bool _isUpdatingForm = false;

        public ProjectProfileSettingsDialog(bool isDarkTheme = false)
        {
            InitializeComponent();
            IsDarkTheme = isDarkTheme;

            Loaded += OnLoaded;
            Closed += (s, e) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                {
                    _currentInstance = null;
                }
            };
        }

        public static void ShowWindow(bool isDarkTheme = false)
        {
            try
            {
                if (_currentInstance != null && _currentInstance.IsLoaded)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new ProjectProfileSettingsDialog(isDarkTheme);

                try
                {
                    var addIn = AddInEvents.Instance;
                    if (addIn?.ExcelAppInstance != null)
                    {
                        var helper = new WindowInteropHelper(_currentInstance);
                        helper.Owner = new IntPtr(addIn.ExcelAppInstance.Hwnd);
                    }
                }
                catch { }

                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(_currentInstance);
                _currentInstance.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProjectProfileSettingsDialog] ShowWindow error: {ex.Message}");
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            var config = ProjectProfileManager.CurrentConfig;
            _activeProfileId = config.ActiveProfileId;

            // Deep clone danh sách để cho phép chỉnh sửa và hủy bỏ
            _profiles = config.Profiles.Select(p => new ProjectProfile
            {
                Id = p.Id,
                Name = p.Name,
                RootFolder = p.RootFolder,
                DetailedDesignFolder = p.DetailedDesignFolder,
                BasicDesignFolder = p.BasicDesignFolder,
                TestSpecFolder = p.TestSpecFolder,
                OpenReadOnlyDefault = p.OpenReadOnlyDefault,
                FileExtensions = p.FileExtensions,
                LastModified = p.LastModified
            }).ToList();

            RefreshProfileList();

            // Chọn active profile hoặc profile đầu tiên
            var target = _profiles.FirstOrDefault(p => p.Id == _activeProfileId) ?? _profiles.FirstOrDefault();
            if (target != null)
            {
                lstProfiles.SelectedItem = target;
            }
        }

        private void RefreshProfileList()
        {
            lstProfiles.ItemsSource = null;
            lstProfiles.ItemsSource = _profiles;
        }

        private void OnProfileSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Lưu form cũ vào model đang chọn trước đó
            SaveCurrentFormToModel();

            _selectedProfile = lstProfiles.SelectedItem as ProjectProfile;
            if (_selectedProfile == null) return;

            _isUpdatingForm = true;
            try
            {
                txtName.Text = _selectedProfile.Name;
                txtRootFolder.Text = _selectedProfile.RootFolder;
                txtTkctFolder.Text = _selectedProfile.DetailedDesignFolder;
                txtTkcbFolder.Text = _selectedProfile.BasicDesignFolder;
                txtTestFolder.Text = _selectedProfile.TestSpecFolder;
                txtExtensions.Text = _selectedProfile.FileExtensions;
                chkReadOnlyDefault.IsChecked = _selectedProfile.OpenReadOnlyDefault;

                UpdateActiveStatusUI();
                ValidateAllFolders();
            }
            finally
            {
                _isUpdatingForm = false;
            }
        }

        private void SaveCurrentFormToModel()
        {
            if (_selectedProfile == null || _isUpdatingForm) return;

            _selectedProfile.Name = txtName.Text.Trim();
            _selectedProfile.RootFolder = txtRootFolder.Text.Trim();
            _selectedProfile.DetailedDesignFolder = txtTkctFolder.Text.Trim();
            _selectedProfile.BasicDesignFolder = txtTkcbFolder.Text.Trim();
            _selectedProfile.TestSpecFolder = txtTestFolder.Text.Trim();
            _selectedProfile.FileExtensions = txtExtensions.Text.Trim();
            _selectedProfile.OpenReadOnlyDefault = chkReadOnlyDefault.IsChecked == true;
        }

        private void UpdateActiveStatusUI()
        {
            if (_selectedProfile == null) return;

            bool isActive = _selectedProfile.Id == _activeProfileId;
            if (isActive)
            {
                lblActiveStatus.Text = LocalizationService.Get("SpecProfile_StatusActive", "Đang kích hoạt (Active)");
                lblActiveStatus.Foreground = new SolidColorBrush(WpfColor.FromRgb(21, 128, 61)); // Green
                btnSetActive.IsEnabled = false;
                btnSetActive.Content = LocalizationService.Get("SpecProfile_BtnIsActive", "✓ Đang Kích Hoạt");
            }
            else
            {
                lblActiveStatus.Text = LocalizationService.Get("SpecProfile_StatusInactive", "Không kích hoạt");
                lblActiveStatus.Foreground = new SolidColorBrush(WpfColor.FromRgb(100, 116, 139)); // Slate
                btnSetActive.IsEnabled = true;
                btnSetActive.Content = LocalizationService.Get("SpecProfile_BtnSetActive", "⭐ Đặt Làm Dự Án Đang Kích Hoạt");
            }
        }

        private void ValidateAllFolders()
        {
            if (_selectedProfile == null) return;

            ValidateFolderField(lblTkctStatus, _selectedProfile.GetTargetFolder(SpecDocumentType.DetailedDesign));
            ValidateFolderField(lblTkcbStatus, _selectedProfile.GetTargetFolder(SpecDocumentType.BasicDesign));
            ValidateFolderField(lblTestStatus, _selectedProfile.GetTargetFolder(SpecDocumentType.TestSpec));
        }

        private void ValidateFolderField(TextBlock lblStatus, string targetPath)
        {
            if (string.IsNullOrWhiteSpace(targetPath))
            {
                lblStatus.Text = LocalizationService.Get("SpecProfile_StatusNotConfigured", "— Chưa cấu hình");
                lblStatus.Foreground = new SolidColorBrush(WpfColor.FromRgb(148, 163, 184));
            }
            else if (Directory.Exists(targetPath))
            {
                lblStatus.Text = LocalizationService.Get("SpecProfile_StatusValid", "✓ Hợp lệ");
                lblStatus.Foreground = new SolidColorBrush(WpfColor.FromRgb(21, 128, 61));
            }
            else
            {
                lblStatus.Text = LocalizationService.Get("SpecProfile_StatusNotFound", "⚠ Không tìm thấy");
                lblStatus.Foreground = new SolidColorBrush(WpfColor.FromRgb(217, 119, 6));
            }
        }

        private void OnFolderFieldChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingForm || _selectedProfile == null) return;
            SaveCurrentFormToModel();
            ValidateAllFolders();
        }

        private void OnSetActiveClick(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;

            _activeProfileId = _selectedProfile.Id;
            UpdateActiveStatusUI();
            RefreshProfileList();
            lstProfiles.SelectedItem = _selectedProfile;
        }

        private void OnAddProfileClick(object sender, RoutedEventArgs e)
        {
            SaveCurrentFormToModel();

            var newProf = new ProjectProfile
            {
                Name = $"{LocalizationService.Get("SpecProfile_NewProjectPrefix", "Dự Án Mới")} {_profiles.Count + 1}",
                RootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                DetailedDesignFolder = "Detailed_Design",
                BasicDesignFolder = "Basic_Design",
                TestSpecFolder = "Test_Specification",
                FileExtensions = ".xlsx;.xlsm;.xls;.docx;.pdf;.pptx",
                OpenReadOnlyDefault = false
            };

            _profiles.Add(newProf);
            if (string.IsNullOrEmpty(_activeProfileId))
            {
                _activeProfileId = newProf.Id;
            }

            RefreshProfileList();
            lstProfiles.SelectedItem = newProf;
            txtName.Focus();
            txtName.SelectAll();
        }

        private void OnCloneProfileClick(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;
            SaveCurrentFormToModel();

            var cloned = _selectedProfile.Clone();
            _profiles.Add(cloned);

            RefreshProfileList();
            lstProfiles.SelectedItem = cloned;
        }

        private void OnDeleteProfileClick(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;

            if (_profiles.Count <= 1)
            {
                WpfMessageBox.Show(
                    this,
                    LocalizationService.Get("SpecProfile_MinProfilePrompt", "Cần duy trì ít nhất 1 Profile dự án trong hệ thống."),
                    LocalizationService.Get("Common_Notice", "Thông Báo"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var result = WpfMessageBox.Show(
                this,
                string.Format(LocalizationService.Get("SpecProfile_DeleteConfirmPrompt", "Bạn có chắc chắn muốn xóa profile dự án '{0}'?"), _selectedProfile.Name),
                LocalizationService.Get("SpecProfile_DeleteConfirmTitle", "Xác Nhận Xóa"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                string deletedId = _selectedProfile.Id;
                _profiles.Remove(_selectedProfile);

                if (_activeProfileId == deletedId)
                {
                    _activeProfileId = _profiles.FirstOrDefault()?.Id;
                }

                RefreshProfileList();
                lstProfiles.SelectedItem = _profiles.FirstOrDefault();
            }
        }

        private void OnBrowseRootFolderClick(object sender, RoutedEventArgs e)
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = LocalizationService.Get("SpecProfile_BrowseRootDesc", "Chọn thư mục gốc của dự án:"),
                SelectedPath = Directory.Exists(txtRootFolder.Text) ? txtRootFolder.Text : string.Empty
            };

            if (fbd.ShowDialog() == FormsDialogResult.OK)
            {
                txtRootFolder.Text = fbd.SelectedPath;
            }
        }

        private void OnBrowseTkctFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseSubFolder(txtTkctFolder, LocalizationService.Get("SpecProfile_BrowseTkctDesc", "Chọn thư mục Thiết Kế Chi Tiết (TKCT / Detailed Design):"));
        }

        private void OnBrowseTkcbFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseSubFolder(txtTkcbFolder, LocalizationService.Get("SpecProfile_BrowseTkcbDesc", "Chọn thư mục Thiết Kế Cơ Bản (TKCB / Basic Design):"));
        }

        private void OnBrowseTestFolderClick(object sender, RoutedEventArgs e)
        {
            BrowseSubFolder(txtTestFolder, LocalizationService.Get("SpecProfile_BrowseTestDesc", "Chọn thư mục Chỉ Thị Test (Test Specification):"));
        }

        private void BrowseSubFolder(System.Windows.Controls.TextBox targetBox, string description)
        {
            string initialPath = string.Empty;
            if (_selectedProfile != null)
            {
                string resolved = _selectedProfile.RootFolder;
                if (Directory.Exists(resolved)) initialPath = resolved;
            }

            using var fbd = new FolderBrowserDialog
            {
                Description = description,
                SelectedPath = initialPath
            };

            if (fbd.ShowDialog() == FormsDialogResult.OK)
            {
                string selectedPath = fbd.SelectedPath;
                string root = txtRootFolder.Text.Trim();

                // Nếu nằm trong root folder, lưu dưới dạng đường dẫn tương đối cho gọn
                if (!string.IsNullOrWhiteSpace(root) && selectedPath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                {
                    string rel = selectedPath.Substring(root.Length).TrimStart('\\', '/');
                    targetBox.Text = string.IsNullOrWhiteSpace(rel) ? "." : rel;
                }
                else
                {
                    targetBox.Text = selectedPath;
                }
            }
        }

        private void OnTestSearchClick(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;
            SaveCurrentFormToModel();

            string keyword = txtTestKeyword.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                lblTestResult.Text = LocalizationService.Get("SpecProfile_TestEnterKeywordPrompt", "⚠ Vui lòng nhập từ khóa test (ví dụ: SCR, M_USER, 001)...");
                lblTestResult.Foreground = new SolidColorBrush(WpfColor.FromRgb(217, 119, 6));
                return;
            }

            var docType = cboTestDocType.SelectedIndex switch
            {
                1 => SpecDocumentType.BasicDesign,
                2 => SpecDocumentType.TestSpec,
                _ => SpecDocumentType.DetailedDesign
            };

            string targetFolder = _selectedProfile.GetTargetFolder(docType);
            if (!Directory.Exists(targetFolder))
            {
                lblTestResult.Text = string.Format(LocalizationService.Get("SpecProfile_TestFolderNotExistPrompt", "⚠ Thư mục không tồn tại: {0}"), targetFolder);
                lblTestResult.Foreground = new SolidColorBrush(WpfColor.FromRgb(239, 68, 68));
                return;
            }

            var hits = ProjectDocumentLauncherService.SearchFiles(_selectedProfile, docType, keyword);
            if (hits.Count == 0)
            {
                lblTestResult.Text = string.Format(LocalizationService.Get("SpecProfile_TestNoMatchesPrompt", "Không tìm thấy file nào khớp với từ khóa '{0}' trong thư mục {1}."), keyword, Path.GetFileName(targetFolder));
                lblTestResult.Foreground = new SolidColorBrush(WpfColor.FromRgb(100, 116, 139));
            }
            else
            {
                var latest = hits.FirstOrDefault(h => h.IsLatest);
                string sampleVer = latest != null ? $" ({LocalizationService.Get("SpecProfile_TestLatestVer", "Bản mới nhất")}: {latest.FileName})" : "";
                lblTestResult.Text = string.Format(LocalizationService.Get("SpecProfile_TestMatchesFoundPrompt", "✓ Tìm thấy {0} file khớp!{1}"), hits.Count, sampleVer);
                lblTestResult.Foreground = new SolidColorBrush(WpfColor.FromRgb(21, 128, 61));
            }
        }

        private void OnSaveClick(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveCurrentFormToModel();

                // Lưu toàn bộ danh sách profiles và active profile vào config một cách an toàn
                ProjectProfileManager.SaveAllProfiles(_profiles, _activeProfileId);

                WpfMessageBox.Show(
                    this,
                    LocalizationService.Get("SpecProfile_SaveSuccess", "Đã lưu thành công cấu hình Profile dự án!"),
                    LocalizationService.Get("Common_Success", "Thành Công"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                Close();
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show(
                    this,
                    $"Lỗi: {ex.Message}",
                    "Lỗi",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
