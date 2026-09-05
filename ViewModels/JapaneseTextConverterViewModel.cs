using System;
using System.Threading.Tasks;
using System.Windows.Input;
using ExcelSupport.Services;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMessageBoxButton = System.Windows.MessageBoxButton;
using WpfMessageBoxImage = System.Windows.MessageBoxImage;

namespace ExcelSupport.ViewModels
{
    public class JapaneseTextConverterViewModel : ViewModelBase
    {
        private readonly ExcelApp _excelApp;

        private bool _toHankaku = true;
        private bool _convertAlpha = true;
        private bool _convertNumbers = true;
        private bool _convertKatakana = true;
        private bool _convertPunctuation = true;
        private bool _convertSpace = true;
        private ConversionScope _scope = ConversionScope.Selection;

        private string _sampleInput = "１２３　ＡＢＣ　アイウエオ　テスト（株）";
        private string _sampleOutput = string.Empty;

        private bool _isProcessing;
        private string _progressStatus = string.Empty;
        private int _progressPercent;
        private string _resultSummary = string.Empty;

        public bool ToHankaku
        {
            get => _toHankaku;
            set
            {
                if (SetProperty(ref _toHankaku, value))
                {
                    OnPropertyChanged(nameof(ToZenkaku));
                    UpdateSampleOutput();
                }
            }
        }

        public bool ToZenkaku
        {
            get => !_toHankaku;
            set => ToHankaku = !value;
        }

        public bool ConvertAlpha
        {
            get => _convertAlpha;
            set { if (SetProperty(ref _convertAlpha, value)) UpdateSampleOutput(); }
        }

        public bool ConvertNumbers
        {
            get => _convertNumbers;
            set { if (SetProperty(ref _convertNumbers, value)) UpdateSampleOutput(); }
        }

        public bool ConvertKatakana
        {
            get => _convertKatakana;
            set { if (SetProperty(ref _convertKatakana, value)) UpdateSampleOutput(); }
        }

        public bool ConvertPunctuation
        {
            get => _convertPunctuation;
            set { if (SetProperty(ref _convertPunctuation, value)) UpdateSampleOutput(); }
        }

        public bool ConvertSpace
        {
            get => _convertSpace;
            set { if (SetProperty(ref _convertSpace, value)) UpdateSampleOutput(); }
        }

        public ConversionScope Scope
        {
            get => _scope;
            set => SetProperty(ref _scope, value);
        }

        public bool IsScopeSelection
        {
            get => _scope == ConversionScope.Selection;
            set { if (value) Scope = ConversionScope.Selection; }
        }

        public bool IsScopeActiveSheet
        {
            get => _scope == ConversionScope.ActiveSheet;
            set { if (value) Scope = ConversionScope.ActiveSheet; }
        }

        public bool IsScopeActiveWorkbook
        {
            get => _scope == ConversionScope.ActiveWorkbook;
            set { if (value) Scope = ConversionScope.ActiveWorkbook; }
        }

        public string SampleInput
        {
            get => _sampleInput;
            set
            {
                if (SetProperty(ref _sampleInput, value))
                {
                    UpdateSampleOutput();
                }
            }
        }

        public string SampleOutput
        {
            get => _sampleOutput;
            set => SetProperty(ref _sampleOutput, value);
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set => SetProperty(ref _isProcessing, value);
        }

        public string ProgressStatus
        {
            get => _progressStatus;
            set => SetProperty(ref _progressStatus, value);
        }

        public int ProgressPercent
        {
            get => _progressPercent;
            set => SetProperty(ref _progressPercent, value);
        }

        public string ResultSummary
        {
            get => _resultSummary;
            set => SetProperty(ref _resultSummary, value);
        }

        public ICommand ConvertCommand { get; }
        public ICommand ResetSampleCommand { get; }

        public JapaneseTextConverterViewModel(ExcelApp excelApp)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));

            ConvertCommand = new RelayCommand(async _ => await ExecuteConversionAsync(), _ => !IsProcessing);
            ResetSampleCommand = new RelayCommand(_ =>
            {
                SampleInput = ToHankaku ? "１２３　ＡＢＣ　アイウエオ　テスト（株）" : "123 ABC ｱｲｳｴｵ ﾃｽﾄ(株)";
            });

            UpdateSampleOutput();
        }

        private void UpdateSampleOutput()
        {
            var options = new JapaneseConversionOptions
            {
                ToHankaku = ToHankaku,
                ConvertAlpha = ConvertAlpha,
                ConvertNumbers = ConvertNumbers,
                ConvertKatakana = ConvertKatakana,
                ConvertPunctuation = ConvertPunctuation,
                ConvertSpace = ConvertSpace
            };

            SampleOutput = JapaneseTextConverterService.ConvertText(SampleInput, options);
        }

        public async Task ExecuteConversionAsync()
        {
            IsProcessing = true;
            ProgressStatus = "Đang bắt đầu chuyển đổi...";
            ProgressPercent = 0;
            ResultSummary = string.Empty;

            var options = new JapaneseConversionOptions
            {
                ToHankaku = ToHankaku,
                ConvertAlpha = ConvertAlpha,
                ConvertNumbers = ConvertNumbers,
                ConvertKatakana = ConvertKatakana,
                ConvertPunctuation = ConvertPunctuation,
                ConvertSpace = ConvertSpace,
                Scope = Scope
            };

            try
            {
                var result = await Task.Run(() =>
                {
                    return JapaneseTextConverterService.ExecuteConversion(
                        _excelApp,
                        options,
                        (msg, p) =>
                        {
                            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                            {
                                ProgressStatus = msg;
                                ProgressPercent = p;
                            });
                        });
                });

                string direction = ToHankaku ? "Toàn giác ➔ Bán giác" : "Bán giác ➔ Toàn giác";
                ResultSummary = $"✅ Chuyển đổi thành công {direction}!\n- Tổng số ô đã duyệt: {result.TotalCellsProcessed:N0}\n- Số ô đã thay đổi: {result.TotalCellsChanged:N0}\n- Số ký tự đã chuyển đổi: {result.TotalCharactersChanged:N0}\n- Thời gian xử lý: {result.Duration.TotalSeconds:F2}s";

                WpfMessageBox.Show(ResultSummary, "Chuyển Đổi Hoàn Tất", WpfMessageBoxButton.OK, WpfMessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                WpfMessageBox.Show($"Lỗi khi chuyển đổi:\n{ex.Message}", "Lỗi", WpfMessageBoxButton.OK, WpfMessageBoxImage.Error);
            }
            finally
            {
                IsProcessing = false;
                ProgressStatus = string.Empty;
            }
        }
    }
}
