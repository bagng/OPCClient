using System.Windows;

namespace OpcClient
{
    /// <summary>
    /// Interaction logic for EditDialog.xaml
    /// </summary>
    public partial class EditDialog : Window
    {
        public delegate void UpdateDelegate(SensorData tItem);
        public UpdateDelegate UpdateCallback = null;

        private SensorData cNewItem, cOldItem;
        public EditDialog(SensorData tItem, UpdateDelegate NewCallback)
        {
            InitializeComponent();

            cNewItem = new SensorData();
            cOldItem = tItem;
            cNewItem.Name = tItem.Name;
            cNewItem.DataType = tItem.DataType;
            cNewItem.Value = tItem.Value;

            NameText.DataContext = cNewItem;
            UnitText.DataContext = cNewItem;
            ValueText.DataContext = cNewItem;

            UpdateCallback += NewCallback;
        }

        private void OK_Button_Click(object sender, RoutedEventArgs e)
        {
//            cOldItem.Value = cNewItem.Value;
            UpdateCallback(cNewItem);
            this.Hide();
        }

        private void Cancel_Button_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }
    }
}
