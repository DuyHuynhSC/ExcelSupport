using System;
using System.Windows.Data;
using System.Windows.Markup;
using ExcelSupport.Services;

namespace ExcelSupport.Helpers
{
    [MarkupExtensionReturnType(typeof(object))]
    public class LocExtension : MarkupExtension
    {
        public string Key { get; set; }

        public LocExtension(string key)
        {
            Key = key;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            if (string.IsNullOrEmpty(Key)) return string.Empty;

            var binding = new System.Windows.Data.Binding($"[{Key}]")
            {
                Source = LocalizationService.Instance,
                Mode = System.Windows.Data.BindingMode.OneWay
            };

            return binding.ProvideValue(serviceProvider);
        }
    }
}
