using CsvHelper;
using Opc;
using Opc.Da;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace OpcClient
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {

        private Opc.Da.Server m_server = null;
        // Define the enumeration based on the COM server interface, used to search all such servers.
        private IDiscovery m_discovery = new OpcCom.ServerEnumerator();

        // Define group objects (subscribers)
        private Subscription mMonitoringSubscription = null;
        // Define the group (subscriber) status, which is equivalent to the group parameter in the OPC specification
        private SubscriptionState mReadGroup = null;

        SensorViewModel cSensorViewModel;
        private string[] MonitoringItemNames;
        private DateTimeViewModel cDateTime;

        SensorData cCurrentCounter;
        int iCurrentCount;

        private DispatcherTimer dataRateTimer;

        public MainWindow()
        {
            cSensorViewModel = new SensorViewModel();
            cDateTime = new DateTimeViewModel();
            cCurrentCounter = new SensorData();
            InitializeComponent();
            SonsorList.DataContext = cSensorViewModel.ListSensorItem;
            DateTimeText.DataContext = cDateTime;
            CountText.DataContext = cCurrentCounter;
            //LoadOPCNames();

            cDateTime.DateTimeCurrent = DateTime.Now;
            dataRateTimer = new DispatcherTimer(DispatcherPriority.Render);
            //dataRateTimer = new DispatcherTimer();
            //dataRateTimer.IsEnabled = true;
            dataRateTimer.Interval = new TimeSpan(0, 0, 0, 0, 100);          // 100ms timer...
            dataRateTimer.Tick += new EventHandler(DataGatheringCallback);
        }

        private void Browse_Button_Click(object sender, RoutedEventArgs e)
        {
            // Query the server
            //Remote, TX1 is the computer name
            // Opc.Server[] servers = m_discovery.GetAvailableServers(Specification.COM_DA_20, "TX1", null);
            Opc.Server[] servers = m_discovery.GetAvailableServers(Specification.COM_DA_20, ServerText.Text.ToString(), null);
//            Opc.Server[] servers = m_discovery.GetAvailableServers(Specification.COM_DA_10, "DESKTOP-MONSTER", null);//Local
            // Indicates the data access specification version, Specification.COMDA_20 is equal to version 2.0.
            // host is the computer name, null means no network security certification is required.

            if (servers != null)
            {
                foreach (Opc.Da.Server server in servers)
                {
                    ServerList.Items.Add(server.Name);
                }
            }
        }

        List<string> mItempNameList;

        private void GetItemsInChildren(BrowseElement[] tParent)
        {
            for (int i = 0; i < tParent.Length; i++)
            {
                if (tParent[i].IsItem == true)
                {
                    mItempNameList.Add(tParent[i].ItemName);
                }
                else {
                    BrowsePosition position;
                    ItemIdentifier tItemChild = new ItemIdentifier();
                    BrowseFilters tFilterChild = new BrowseFilters();
                    tFilterChild.BrowseFilter = Opc.Da.browseFilter.all;
                    tItemChild.ItemName = tParent[i].ItemName;
                    //tItemChild.ItemPath = tChildren[i].ItemPath;
                    BrowseElement[] tChildren = m_server.Browse(tItemChild, tFilterChild, out position);

                    GetItemsInChildren(tChildren);
                }
            }
        }

        private void GetAllItemInServer()
        {
            mItempNameList = new List<string>();

            BrowsePosition position;
            ItemIdentifier tItem = new ItemIdentifier();
            BrowseFilters tFilter = new BrowseFilters();
            tFilter.BrowseFilter = Opc.Da.browseFilter.all;
            BrowseElement[] children = m_server.Browse(tItem, tFilter, out position);
            GetItemsInChildren(children);
            MonitoringItemNames = new string[mItempNameList.Count];
            for (int i = 0; i < mItempNameList.Count; i++)
            {
                MonitoringItemNames[i] = mItempNameList[i];
            }
        }

        private void Connect_Button_Click(object sender, RoutedEventArgs e)
        {
            Opc.Server[] servers = m_discovery.GetAvailableServers(Specification.COM_DA_20, ServerText.Text.ToString(), null);
            // Indicates the data access specification version, Specification.COMDA_20 is equal to version 2.0.
            // host is the computer name, null means no network security certification is required.

            if (servers != null)
            {
                foreach (Opc.Da.Server server in servers)
                {
                    if (String.Compare(server.Name, ServerNameText.Text, true) == 0)    // For "true" to ignore case
                    {
                        //establish connection.
                        m_server = server;
                        break;
                    }
                }
            }

            //connect to the server
            if (m_server != null)
                m_server.Connect();
            else
                return;

            var serverStatus = m_server.GetStatus();

            string strTemp = serverStatus.VendorInfo;
            strTemp = serverStatus.StatusInfo;
            strTemp = serverStatus.ProductVersion;
            serverState tState = serverStatus.ServerState;
            DateTime tTime = serverStatus.CurrentTime;
            tTime = serverStatus.StartTime;
            tTime = serverStatus.LastUpdateTime;

            GetAllItemInServer();

            // Set group status
            // Group (subscriber) status, equivalent to the parameters of the group in the OPC specification
            mReadGroup = new SubscriptionState();
            mReadGroup.Name = "Monitoring";                       // Group Name
            mReadGroup.ServerHandle = null;                          // The handle assigned by the server to the group.
            mReadGroup.ClientHandle = Guid.NewGuid().ToString();     // The handle assigned by the client to the group.
            mReadGroup.Active = true;                                // Activate the group.
            mReadGroup.UpdateRate = 10;                             // The refresh rate is 1 second. -> 1000
            mReadGroup.Deadband = 0;                                 // When the dead zone value is set to 0, the server will notify the group of any data changes in the group.
            mReadGroup.Locale = null;                                //No regional values are set.

            // Add Group
            mMonitoringSubscription = (Subscription)m_server.CreateSubscription(mReadGroup); // Create Group

            // Define Item List
            Item[] items = new Item[MonitoringItemNames.Length];                             // Define the data item, ie item
            for (int i = 0; i < items.Length; i++)                      // Item initial assignment
            {
                items[i] = new Item();                              // Create an Item object.
                items[i].ClientHandle = Guid.NewGuid().ToString();  // The handle assigned by the client to the data item.
                items[i].ItemPath = null;                           // The path of the data item in the server.
                items[i].ItemName = MonitoringItemNames[i];                    // The name of the data item in the server.
                SensorData tSensorItem;
                tSensorItem = new SensorData();                              // Create an Item object.
                tSensorItem.Name = MonitoringItemNames[i];
                cSensorViewModel.ListSensorItem.Add(tSensorItem);
            }

            // Add Item
            ItemResult[] tItemResult = mMonitoringSubscription.AddItems(items);
            ItemValueResult[] itemValues = mMonitoringSubscription.Read(mMonitoringSubscription.Items);
            foreach (ItemValueResult titem in itemValues)
            {
                cSensorViewModel.AddType(titem.ItemName, titem.Value.GetType().ToString());
            }
            //Thread.Sleep(500); // sleep 500ms.
            OnDataChange(mMonitoringSubscription, itemValues);
            // Register callback event
            mMonitoringSubscription.DataChanged += new DataChangedEventHandler(OnDataChange);

            //SubscriptionState tSubState = mMonitoringSubscription.GetState();
        }

        // DataChange callback
        public void OnDataChange(object subscriptionHandle, ItemValueResult[] values)
        {
            foreach (ItemValueResult item in values)
            {
                cSensorViewModel.AddValue(item.ItemName, item.Value);
            }
        }

        private void Close_Button_Click(object sender, RoutedEventArgs e)
        {
            cSensorViewModel.ListSensorItem.Clear();
            if (mMonitoringSubscription != null)
            {
                mMonitoringSubscription.RemoveItems(mMonitoringSubscription.Items);
                m_server.CancelSubscription(mMonitoringSubscription);
                mMonitoringSubscription.Dispose();
                mMonitoringSubscription = null;
            }
            if (m_server != null)
            {
                m_server.Disconnect();
                m_server = null;
            }
            this.Close();
        }

        private void Start_Button_Click(object sender, RoutedEventArgs e)
        {
            if (m_server == null)
                return;

            DataGatheringCallback(sender, e);
            dataRateTimer.Start();
        }

        public void DataGatheringCallback(object sender, EventArgs e)
        {
            string strFilename = "data.csv";
            bool bCreated = false;

            cDateTime.DateTimeCurrent = DateTime.Now;
            TextWriter fileWriter;
            if (!File.Exists(strFilename))
            {
                fileWriter = File.CreateText(strFilename);
                bCreated = true;
            }
            else
            {
                fileWriter = File.AppendText(strFilename);
            }

            var csv = new CsvWriter(fileWriter);
            if (bCreated)
            {
                //csv.WriteField("1");
                //csv.NextRecord();
                for (int i = 0; i < cSensorViewModel.ListSensorItem.Count; i++)
                    csv.WriteField(cSensorViewModel.ListSensorItem[i].Name);
                csv.NextRecord();
            }

            csv.WriteField(cDateTime.DateTimeCurrent.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            for (int i = 0; i < cSensorViewModel.ListSensorItem.Count; i++)
            {
                csv.WriteField(cSensorViewModel.ListSensorItem[i].Value);
            }
            csv.NextRecord();
            fileWriter.Close();
            iCurrentCount++;
            cCurrentCounter.Value = iCurrentCount.ToString();
        }

        private void Stop_Button_Click(object sender, RoutedEventArgs e)
        {
            dataRateTimer.Stop();
        }

        private void TagListGridViewColumnHeaderClickedHandler(object sender, RoutedEventArgs e)
        {

        }

        private void ServerList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var itemString = ServerList.SelectedItem as String;
            ServerNameText.Text = itemString;
        }

        private void Export_Button_Click(object sender, RoutedEventArgs e)
        {
            if (m_server == null)
                return;
            string strFilename = "items.csv";

            cDateTime.DateTimeCurrent = DateTime.Now;
            TextWriter fileWriter;
            fileWriter = File.CreateText(strFilename);

            var csv = new CsvWriter(fileWriter);
            //csv.WriteField("1");
            //csv.NextRecord();
            for (int i = 0; i < cSensorViewModel.ListSensorItem.Count; i++)
            {
                csv.WriteField(cSensorViewModel.ListSensorItem[i].Name);
                csv.WriteField(cSensorViewModel.ListSensorItem[i].DataType);
                csv.NextRecord();
            }

            csv.NextRecord();
            fileWriter.Close();
        }
    }

    public class DateTimeViewModel : INotifyPropertyChanged
    {
        private int _iYear;
        public int iYear
        {
            get { return _iYear; }
            set
            {
                _iYear = value;
                OnPropertyChanged();
            }
        }

        private int _iMonth;
        public int iMonth
        {
            get { return _iMonth; }
            set
            {
                _iMonth = value;
                OnPropertyChanged();
            }
        }

        private int _iDay;
        public int iDay
        {
            get { return _iDay; }
            set
            {
                _iDay = value;
                OnPropertyChanged();
            }
        }

        private int _iHour;
        public int iHour
        {
            get { return _iHour; }
            set
            {
                _iHour = value;
                OnPropertyChanged();
            }
        }

        private int _iMinute;
        public int iMinute
        {
            get { return _iMinute; }
            set
            {
                _iMinute = value;
                OnPropertyChanged();
            }
        }

        private int _iSecond;
        public int iSecond
        {
            get { return _iSecond; }
            set
            {
                _iSecond = value;
                OnPropertyChanged();
            }
        }

        private int _iMilisecond;
        public int iMilisecond
        {
            get { return _iMilisecond; }
            set
            {
                _iMilisecond = value;
                OnPropertyChanged();
            }
        }

        private DateTime _DateTimeCurrent;
        public DateTime DateTimeCurrent
        {
            get { return _DateTimeCurrent; }
            set
            {
                _DateTimeCurrent = value;
                iYear = _DateTimeCurrent.Year;
                iMonth = _DateTimeCurrent.Month;
                iDay = _DateTimeCurrent.Day;
                iHour = _DateTimeCurrent.Hour;
                iMinute = _DateTimeCurrent.Minute;
                iSecond = _DateTimeCurrent.Second;
                iMilisecond = _DateTimeCurrent.Millisecond;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SensorData : INotifyPropertyChanged
    {
        private string _name;
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        private string _dataType;
        public string DataType
        {
            get { return _dataType; }
            set { _dataType = value; }
        }

        private string _value;
        public string Value
        {
            get { return _value; }
            set
            {
                _value = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SensorViewModel
    {
        public ObservableCollection<SensorData> ListSensorItem;

        public SensorViewModel()
        {
            ListSensorItem = new ObservableCollection<SensorData>();
        }

        public void AddValue(string tName, object tValue)
        {
            for (int i = 0; i < ListSensorItem.Count; i++)
            {
                if (ListSensorItem[i].Name == tName)
                {
                    //ListSensorItem[i].Value = System.Convert.ToDouble(tValue);
                    ListSensorItem[i].Value = tValue.ToString();
                }
            }
        }
        public void AddType(string tName, string strType)
        {
            for (int i = 0; i < ListSensorItem.Count; i++)
            {
                if (ListSensorItem[i].Name == tName)
                {
                    ListSensorItem[i].DataType = strType;
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
