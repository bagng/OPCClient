using CsvHelper;
using Opc;
using Opc.Da;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace OpcClient
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        private Opc.Da.Server m_server = null;
        // Define the enumeration based on the COM server interface, used to search all such servers.
        private IDiscovery m_discovery;

        // Define group objects (subscribers)
        private Subscription mMonitoringSubscription = null;
        // Define the group (subscriber) status, which is equivalent to the group parameter in the OPC specification
        private SubscriptionState mMonitoringGroup = null;

        private SensorViewModel cSensorViewModel;
        private string[] MonitoringItemNames;
        private DateTimeViewModel cDateTime;

        private SensorData cCurrentCounter;
        private int iCurrentCount;

        private DispatcherTimer dataRateTimer;
        private SensorData cUpdateRate;

        public MainWindow()
        {
            m_discovery = new OpcCom.ServerEnumerator();
            cSensorViewModel = new SensorViewModel();
            cDateTime = new DateTimeViewModel();
            cCurrentCounter = new SensorData();
            cUpdateRate = new SensorData();
            InitializeComponent();
            SonsorList.DataContext = cSensorViewModel.ListSensorItem;
            DateTimeText.DataContext = cDateTime;
            CountText.DataContext = cCurrentCounter;
            UpdateRateText.DataContext = cUpdateRate;
            cUpdateRate.Value = "10";
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
            ServerList.Items.Clear();

            // Query the server
            //Remote, TX1 is the computer name
            // Opc.Server[] servers = m_discovery.GetAvailableServers(Specification.COM_DA_20, "TX1", null);
            Opc.Server[] servers = m_discovery.GetAvailableServers(Specification.COM_DA_30, ServerText.Text.ToString(), null);
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
            mMonitoringGroup = new SubscriptionState();
            mMonitoringGroup.Name = "Monitoring";                       // Group Name
            mMonitoringGroup.ServerHandle = null;                          // The handle assigned by the server to the group.
            mMonitoringGroup.ClientHandle = Guid.NewGuid().ToString();     // The handle assigned by the client to the group.
            mMonitoringGroup.Active = true;                                // Activate the group.
            mMonitoringGroup.UpdateRate = int.Parse(cUpdateRate.Value);    // The refresh rate is 1 second. -> 1000
            mMonitoringGroup.Deadband = 0;                                 // When the dead zone value is set to 0, the server will notify the group of any data changes in the group.
            mMonitoringGroup.Locale = null;                                //No regional values are set.

            // Add Group
            mMonitoringSubscription = (Subscription)m_server.CreateSubscription(mMonitoringGroup); // Create Group

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
            OnDataChange(mMonitoringSubscription, null, itemValues);
            // Register callback event
            mMonitoringSubscription.DataChanged += new DataChangedEventHandler(OnDataChange);

            //SubscriptionState tSubState = mMonitoringSubscription.GetState();
        }

        // DataChange callback
        public void OnDataChange(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            foreach (ItemValueResult item in values)
            {
                cSensorViewModel.AddValue(item.ItemName, item.Value);
            }
        }

        private void Close_Button_Click(object sender, RoutedEventArgs e)
        {
            dataRateTimer.Stop();

            if (mMonitoringSubscription != null)
            {
                mMonitoringSubscription.DataChanged -= OnDataChange;
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
            cSensorViewModel.ListSensorItem.Clear();

            cSensorViewModel = null;
            cDateTime = null;
            cCurrentCounter = null;
            cUpdateRate = null;

            this.Close();
        }

        private bool disposed;

        public void Dispose()
        {
            this.Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed) return;
            if (disposing)
            {
                // IDisposable 인터페이스를 구현하는 멤버들을 여기서 정리합니다.
            }
            // .NET Framework에 의하여 관리되지 않는 외부 리소스들을 여기서 정리합니다.
            m_discovery.Dispose();
            this.disposed = true;
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

        GridViewColumnHeader _lastHeaderClicked = null;
        ListSortDirection _lastDirection = ListSortDirection.Ascending;
        private SortAdorner listViewSortAdorner = null;

        private void TagListGridViewColumnHeaderClickedHandler(object sender, RoutedEventArgs e)
        {
            GridViewColumnHeader headerClicked =
                  e.OriginalSource as GridViewColumnHeader;
            ListSortDirection direction;

            if (headerClicked != null)
            {
                if (headerClicked.Role != GridViewColumnHeaderRole.Padding)
                {
                    if (listViewSortAdorner != null)
                        AdornerLayer.GetAdornerLayer(headerClicked).Remove(listViewSortAdorner);
                    if (headerClicked != _lastHeaderClicked)
                    {
                        direction = ListSortDirection.Ascending;
                    }
                    else
                    {
                        if (_lastDirection == ListSortDirection.Ascending)
                        {
                            direction = ListSortDirection.Descending;
                        }
                        else
                        {
                            direction = ListSortDirection.Ascending;
                        }
                    }
                    listViewSortAdorner = new SortAdorner(headerClicked, direction);
                    AdornerLayer.GetAdornerLayer(headerClicked).Add(listViewSortAdorner);

                    string header = headerClicked.Column.Header as string;
                    Sort(header, direction);
                    /*
                    if (direction == ListSortDirection.Ascending)
                    {
                        if(headerClicked.Column.Header.ToString().LastIndexOf(" ") == -1)
                        headerClicked.Column.Header = header + " △";
                    }
                    else
                    {
                        if (headerClicked.Column.Header.ToString().LastIndexOf(" ") == -1)
                            headerClicked.Column.Header = header + " ▽";

                    }*/
                    // Remove arrow from previously sorted header
                    if (_lastHeaderClicked != null && _lastHeaderClicked != headerClicked)
                    {
                        _lastHeaderClicked.Column.HeaderTemplate = null;
                    }
                    _lastHeaderClicked = headerClicked;
                    _lastDirection = direction;
                }
            }
        }

        /// <summary>
        /// 정렬하기
        /// </summary>
        /// <param name="header">헤더</param>
        /// <param name="listSortDirection">리스트 정렬 방향</param>
        private void Sort(string sortBy, ListSortDirection direction)
        {
            ICollectionView dataView =
              CollectionViewSource.GetDefaultView(SonsorList.ItemsSource);

            dataView.SortDescriptions.Clear();
            SortDescription sd = new SortDescription("Name", direction);
            dataView.SortDescriptions.Add(sd);
            dataView.Refresh();
        }

        public class SortAdorner : Adorner
        {
            private static Geometry ascGeometry =
                Geometry.Parse("M 0 4 L 3.5 0 L 7 4 Z");

            private static Geometry descGeometry =
                Geometry.Parse("M 0 0 L 3.5 4 L 7 0 Z");

            public ListSortDirection Direction { get; private set; }

            public SortAdorner(UIElement element, ListSortDirection dir)
                : base(element)
            {
                this.Direction = dir;
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                base.OnRender(drawingContext);

                if (AdornedElement.RenderSize.Width < 20)
                    return;

                TranslateTransform transform = new TranslateTransform
                    (
                        AdornedElement.RenderSize.Width - 15,
                        (AdornedElement.RenderSize.Height - 5) / 2
                    );
                drawingContext.PushTransform(transform);

                Geometry geometry = ascGeometry;
                if (this.Direction == ListSortDirection.Descending)
                    geometry = descGeometry;
                drawingContext.DrawGeometry(Brushes.Black, null, geometry);

                drawingContext.Pop();
            }
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

        private void SonsorList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListView tList = sender as ListView;
            SensorData tItem = tList.SelectedItem as SensorData;
            EditDialog cDlg = new EditDialog(tItem, UpdateOPCData);
            cDlg.ShowDialog();
            cDlg = null;
        }


        public void UpdateOPCData(SensorData tItem)
        {
            Item OPC_WriteItem = Array.Find(mMonitoringSubscription.Items, x => x.ItemName.Equals(tItem.Name));
            ItemValue[] writeValues = new ItemValue[1];
            writeValues[0] = new ItemValue(OPC_WriteItem);
            writeValues[0].Value = tItem.Value;
            IdentifiedResult[] retValues = mMonitoringSubscription.Write(writeValues);
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
                _DateTimeCurrent = new DateTime(_iYear, _DateTimeCurrent.Month, _DateTimeCurrent.Day);
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
                _DateTimeCurrent = new DateTime(_DateTimeCurrent.Year, _iMonth, _DateTimeCurrent.Day);
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
                _DateTimeCurrent = new DateTime(_DateTimeCurrent.Year, _DateTimeCurrent.Month, _iDay);
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
                _DateTimeCurrent = new DateTime(_DateTimeCurrent.Year, _DateTimeCurrent.Month, _DateTimeCurrent.Day, _iHour, _DateTimeCurrent.Minute, _DateTimeCurrent.Second, _DateTimeCurrent.Millisecond);
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
                _DateTimeCurrent = new DateTime(_DateTimeCurrent.Year, _DateTimeCurrent.Month, _DateTimeCurrent.Day, _DateTimeCurrent.Hour, _iMinute, _DateTimeCurrent.Second, _DateTimeCurrent.Millisecond);
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
                _DateTimeCurrent = new DateTime(_DateTimeCurrent.Year, _DateTimeCurrent.Month, _DateTimeCurrent.Day, _DateTimeCurrent.Hour, _DateTimeCurrent.Minute, _iSecond, _DateTimeCurrent.Millisecond);
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
                _DateTimeCurrent = new DateTime(_DateTimeCurrent.Year, _DateTimeCurrent.Month, _DateTimeCurrent.Day, _DateTimeCurrent.Hour, _DateTimeCurrent.Minute, _DateTimeCurrent.Second, _iMilisecond);
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
                _iYear = _DateTimeCurrent.Year;
                _iMonth = _DateTimeCurrent.Month;
                _iDay = _DateTimeCurrent.Day;
                _iHour = _DateTimeCurrent.Hour;
                _iMinute = _DateTimeCurrent.Minute;
                _iSecond = _DateTimeCurrent.Second;
                _iMilisecond = _DateTimeCurrent.Millisecond;
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
