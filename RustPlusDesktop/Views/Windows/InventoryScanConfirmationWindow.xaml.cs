using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using RustPlusDesk.Services.InventoryScan;

namespace RustPlusDesk.Views
{
    public partial class InventoryScanConfirmationWindow : Window
    {
        public ObservableCollection<RecognizedInventoryStackViewModel> Items { get; } = new();

        public InventoryScanConfirmationWindow(
            IEnumerable<RecognizedInventoryStack> results, 
            IEnumerable<RecyclerItemViewModel> recyclableItems)
        {
            InitializeComponent();
            ResultsListView.ItemsSource = Items;

            foreach (var res in results)
            {
                var refItem = recyclableItems.FirstOrDefault(x => 
                    string.Equals(x.ShortName, res.ShortName, StringComparison.OrdinalIgnoreCase));
                
                var vm = new RecognizedInventoryStackViewModel
                {
                    ShouldApply = res.IconConfidence >= 0.60, // Unchecked by default if low confidence
                    ShortName = res.ShortName,
                    DisplayName = res.DisplayName,
                    Icon = refItem?.Icon,
                    Quantity = res.Quantity,
                    IconConfidence = res.IconConfidence
                };
                Items.Add(vm);
            }
        }

        public IEnumerable<RecognizedInventoryStack> GetApprovedResults()
        {
            return Items
                .Where(x => x.ShouldApply)
                .Select(x => new RecognizedInventoryStack
                {
                    ShortName = x.ShortName,
                    DisplayName = x.DisplayName,
                    Quantity = x.Quantity,
                    IconConfidence = x.IconConfidence
                });
        }

        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        protected override void OnMouseLeftButtonDown(System.Windows.Input.MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
    }

    public sealed class RecognizedInventoryStackViewModel : INotifyPropertyChanged
    {
        private bool _shouldApply = true;
        public bool ShouldApply
        {
            get => _shouldApply;
            set { _shouldApply = value; OnPropertyChanged(nameof(ShouldApply)); }
        }

        public string ShortName { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public ImageSource? Icon { get; set; }

        private string _quantityText = "1";
        public string QuantityText
        {
            get => _quantityText;
            set { _quantityText = value; OnPropertyChanged(nameof(QuantityText)); }
        }

        public int Quantity
        {
            get
            {
                if (int.TryParse(_quantityText, out int val))
                    return val;
                return 1;
            }
            set
            {
                _quantityText = value.ToString();
                OnPropertyChanged(nameof(QuantityText));
            }
        }

        public double IconConfidence { get; set; }
        
        public string ConfidenceText => $"{IconConfidence * 100:0}%";
        
        public Brush ConfidenceBrush => IconConfidence >= 0.60 
            ? new SolidColorBrush(Color.FromRgb(0x13, 0x7E, 0x3E)) // rich green (excellent contrast with white text)
            : new SolidColorBrush(Color.FromRgb(0xAD, 0x36, 0x25)); // rich dark red (excellent contrast with white text)

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
