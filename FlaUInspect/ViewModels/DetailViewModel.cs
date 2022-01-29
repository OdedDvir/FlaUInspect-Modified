using FlaUI.Core;
using FlaUInspect.Core;

namespace FlaUInspect.ViewModels
{
    public interface IDetailViewModel
    {
        string Key { get; set; }
        string Value { get; set; }
        bool Important { get; set; }
    }

    public class DetailViewModel : ObservableObject, IDetailViewModel
    {
        public DetailViewModel(string key, string value)
        {
            Key = key;
            Value = value;
        }

        public string Key { get { return GetProperty<string>("Key"); } set { SetProperty(value,"Key"); } }
        public string Value { get { return GetProperty<string>("Value"); } set { SetProperty(value,"Value"); } }
        public bool Important { get { return GetProperty<bool>("Important"); } set { SetProperty(value,"Important"); } }

        public static DetailViewModel FromAutomationProperty<T>(string key, IAutomationProperty<T> value)
        {
            return new DetailViewModel(key, value.ToDisplayText());
        }
    }
}
