using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection;
using System.Windows.Input;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA2;
using FlaUI.UIA3;
using FlaUInspect.Core;
using Microsoft.Win32;
using System.IO;
using System.Runtime.Serialization;
using Newtonsoft.Json;
using System.Text;
using System.Windows.Forms;

namespace FlaUInspect.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        private HoverMode _hoverMode;
        private FocusTrackingMode _focusTrackingMode;
        private ITreeWalker _treeWalker;
        private AutomationBase _automation;
        private AutomationElement _rootElement;

        public MainViewModel()
        {
            Elements = new ObservableCollection<ElementViewModel>();
            StartNewInstanceCommand = new RelayCommand(o =>
            {
                var info = new ProcessStartInfo(Assembly.GetExecutingAssembly().Location);
                Process.Start(info);
            });
            CaptureSelectedItemCommand = new RelayCommand(o =>
            {
                if (SelectedItemInTree == null)
                {
                    return;
                }
                var capturedImage = SelectedItemInTree.AutomationElement.Capture();
                var saveDialog = new Microsoft.Win32.SaveFileDialog();
                saveDialog.Filter = "Png file (*.png)|*.png";
                if (saveDialog.ShowDialog() == true)
                {
                    capturedImage.Save(saveDialog.FileName, ImageFormat.Png);
                }
                capturedImage.Dispose();
            });
            RefreshCommand = new RelayCommand(o =>
            {
                RefreshTree();
            });

            DumpTree = new RelayCommand(o =>
            {
                //Get selected branch
                ElementViewModel obj = Elements[0].Children.Where(e => e.IsSelected == true).FirstOrDefault();
                if (obj == null)
                {
                    MessageBox.Show("Please first select a branch.");
                    return;
                }
                string branchName = obj.Name;
                DialogResult dialogResult = MessageBox.Show($"Would you like to dump this branch to a Json file??\n{branchName}", "Dump this Branch?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    UIElement root = ScanElements(obj);
                    string jsonUI = JsonConvert.SerializeObject(root, Formatting.Indented);
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string outputFileName = branchName.ReplaceInvalidChars();
                    File.WriteAllText($@"{desktopPath}\{outputFileName}.json", jsonUI, Encoding.UTF8);
                    MessageBox.Show($"File:\n\n{outputFileName}.json\n\nWas created on your desktop.");

                }
            });

        }

        public UIElement ScanElements(ElementViewModel e)
        {
            UIElement ui =new UIElement();
            ui.Name = e.Name;
            ui.ControlType = e.ControlType.ToString();
            // ui.Value = e.ItemDetails.ToJson();
            e.LoadChildren(true);
            //ui.ChildrenCount = e.Children.Count;
            if (e.Children.Count > 0)
            {
                ui.Children = new List<UIElement>();
                foreach (var item in e.Children)
                {
                    ui.Children.Add(ScanElements(item));
                }
            }
            return ui;
        }

        public bool IsInitialized
        {
            get { return GetProperty<bool>("IsInitialized"); }
            private set { SetProperty(value, "IsInitialized"); }
        }

        public bool EnableHoverMode
        {
            get { return GetProperty<bool>("EnableHoverMode"); }
            set
            {
                if (SetProperty(value, "EnableHoverMode"))
                {
                    if (value) { _hoverMode.Start(); }
                    else { _hoverMode.Stop(); }
                }
            }
        }

        public bool EnableFocusTrackingMode
        {
            get { return GetProperty<bool>("EnableFocusTrackingMode"); }
            set
            {
                if (SetProperty(value, "EnableFocusTrackingMode"))
                {
                    if (value) { _focusTrackingMode.Start(); }
                    else { _focusTrackingMode.Stop(); }
                }
            }
        }

        public bool EnableXPath
        {
            get { return GetProperty<bool>("EnableXPath"); }
            set { SetProperty(value, "EnableXPath"); }
        }

        public AutomationType SelectedAutomationType
        {
            get { return GetProperty<AutomationType>("SelectedAutomationType"); }
            private set { SetProperty(value, "SelectedAutomationType"); }
        }

        public ObservableCollection<ElementViewModel> Elements { get; private set; }

        public ICommand StartNewInstanceCommand { get; private set; }
        public ICommand CaptureSelectedItemCommand { get; private set; }
        public ICommand RefreshCommand { get; private set; }
        public ICommand DumpTree { get; private set; } //eden

        public ObservableCollection<DetailGroupViewModel> SelectedItemDetails => SelectedItemInTree?.ItemDetails;

        public ElementViewModel SelectedItemInTree
        {
            get { return GetProperty<ElementViewModel>("SelectedItemInTree"); }
            private set { SetProperty(value, "SelectedItemInTree"); }
        }

        public void Initialize(AutomationType selectedAutomationType)
        {
            SelectedAutomationType = selectedAutomationType;
            IsInitialized = true;

            _automation = selectedAutomationType == AutomationType.UIA2 ? (AutomationBase)new UIA2Automation() : new UIA3Automation();
            _rootElement = _automation.GetDesktop();
            var desktopViewModel = new ElementViewModel(_rootElement);
            desktopViewModel.SelectionChanged += DesktopViewModel_SelectionChanged;
            desktopViewModel.LoadChildren(false);
            Elements.Add(desktopViewModel);
            Elements[0].IsExpanded = true;

            // Initialize TreeWalker
            _treeWalker = _automation.TreeWalkerFactory.GetControlViewWalker();

            // Initialize hover
            _hoverMode = new HoverMode(_automation);
            _hoverMode.ElementHovered += ElementToSelectChanged;

            // Initialize focus tracking
            _focusTrackingMode = new FocusTrackingMode(_automation);
            _focusTrackingMode.ElementFocused += ElementToSelectChanged;
        }

        private void ElementToSelectChanged(AutomationElement obj)
        {
            // Build a stack from the root to the hovered item
            var pathToRoot = new Stack<AutomationElement>();
            while (obj != null)
            {
                // Break on circular relationship (should not happen?)
                if (pathToRoot.Contains(obj) || obj.Equals(_rootElement)) { break; }

                pathToRoot.Push(obj);
                try
                {
                    obj = _treeWalker.GetParent(obj);
                }
                catch (Exception ex)
                {
                    // TODO: Log
                    Console.WriteLine($"Exception: {ex.Message}");
                }
            }

            // Expand the root element if needed
            if (!Elements[0].IsExpanded)
            {
                Elements[0].IsExpanded = true;
                System.Threading.Thread.Sleep(1000);
            }

            var elementVm = Elements[0];
            while (pathToRoot.Count > 0)
            {
                var elementOnPath = pathToRoot.Pop();
                var nextElementVm = FindElement(elementVm, elementOnPath);
                if (nextElementVm == null)
                {
                    // Could not find next element, try reloading the parent
                    elementVm.LoadChildren(true);
                    // Now search again
                    nextElementVm = FindElement(elementVm, elementOnPath);
                    if (nextElementVm == null)
                    {
                        // The next element is still not found, exit the loop
                        Console.WriteLine("Could not find the next element!");
                        break;
                    }
                }
                elementVm = nextElementVm;
                if (!elementVm.IsExpanded)
                {
                    elementVm.IsExpanded = true;
                }
            }
            // Select the last element
            elementVm.IsSelected = true;
        }

        private ElementViewModel FindElement(ElementViewModel parent, AutomationElement element)
        {
            return parent.Children.FirstOrDefault(child => child.AutomationElement.Equals(element));
        }

        private void DesktopViewModel_SelectionChanged(ElementViewModel obj)
        {
            SelectedItemInTree = obj;
            OnPropertyChanged(() => SelectedItemDetails);
        }

        private void RefreshTree()
        {
            Elements.Clear();
            Initialize(SelectedAutomationType);
        }
    }
}
