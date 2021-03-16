using System;
using System.Collections.Generic;
using System.Windows.Input;
using System.Threading.Tasks;
using LiveCharts;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using LiveCharts.Wpf;
using System.Windows;
using System.Linq;

namespace SpasticityClient
{
    public class MeasureModel : INotifyPropertyChanged
    {
        public DateTime DateTime { get; set; }
        private double _value { get; set; }
        public double Value
        {
            get { return _value; }
            set
            {
                _value = value;
                OnPropertyChanged("Value");
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            //Raise PropertyChanged event
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ChartModel : INotifyPropertyChanged, IDisposable
    {
        #region variables
        private double _min;
        private double _max;
        private double _value;
        private double _count;
        private XBeeData _xbeeData;
        private string _buttonLabel;
        private string _channelName;
        private float _batteryLevel;
        #endregion

        #region properties
        public Queue<string> csvData { get; set; }
        public RelayCommand ReadCommand { get; set; }
        public bool IsRunning { get; set; }

        public ChartValues<MeasureModel> EMGValues { get; set; }
        public Func<double, string> DateTimeFormatter { get; set; }
        public double AxisStep { get; set; }
        public double AxisUnit { get; set; }

        public double Min
        {
            get { return _min; }
            set
            {
                _min = value;
                OnPropertyChanged("Min");
            }
        }
        public double Max
        {
            get { return _max; }
            set
            {
                _max = value;
                OnPropertyChanged("Max");
            }
        }
        public double Value
        {
            get { return _value; }
            set
            {
                _value = value;
                OnPropertyChanged("Value");
            }
        }
        public double Count
        {
            get { return _count; }
            set
            {
                _count = value;
                OnPropertyChanged("Count");
            }
        }
        public string ChannelName
        {
            get { return _channelName; }
            set
            {
                _channelName = value;
                NotifyPropertyChanged("ChannelName");
            }
        }
        public string ButtonLabel
        {
            get { return _buttonLabel; }
            set
            {
                _buttonLabel = value;
                NotifyPropertyChanged("ButtonLabel");
            }
        }
        public string BatteryColor
        {
            get
            {
                if (_batteryLevel > 3.5)
                    return "Green";
                if (_batteryLevel > 3.3 && _batteryLevel <= 3.5)
                    return "Yellow";
                else
                    return "Red";
            }
        }
        public float BatteryLevel
        {
            get { return _batteryLevel; }
            set
            {
                _batteryLevel = value;
                NotifyPropertyChanged("BatteryLevel");
                NotifyPropertyChanged("BatteryColor");
            }
        }

        public string PortName { get; set; }
        #endregion

        #region event
        public event PropertyChangedEventHandler PropertyChanged;
        #endregion

        #region constructor
        public ChartModel(string portname)
        {
            PortName = portname;
            _xbeeData = new XBeeData(portname);
            ReadCommand = new RelayCommand(Read);
            csvData = new Queue<string>();
            ButtonLabel = "Start Reading";
            IsRunning = false;
            BatteryLevel = 0;

            //EMGValues = new ChartValues<double>();

            var mapper = LiveCharts.Configurations.Mappers.Xy<MeasureModel>()
                .X(model => model.DateTime.Ticks)   //use DateTime.Ticks as X
                .Y(model => model.Value);           //use the value property as Y

            //lets save the mapper globally.
            Charting.For<MeasureModel>(mapper);

            //the values property will store our values array
            EMGValues = new ChartValues<MeasureModel>();

            //lets set how to display the X Labels
            DateTimeFormatter = value => new DateTime((long)value).ToString("mm:ss");

            //AxisStep forces the distance between each separator in the X axis
            AxisStep = TimeSpan.FromSeconds(1).Ticks;
            //AxisUnit forces lets the axis know that we are plotting seconds
            //this is not always necessary, but it can prevent wrong labeling
            AxisUnit = TimeSpan.TicksPerSecond;

            SetAxisLimits(DateTime.Now);
        }
        #endregion

        #region method

        public void Read()
        {
            try
            {
                //lets keep in memory only the last 200 records,
                //to keep everything running faster
                if (!IsRunning)
                {
                    _xbeeData.IsCancelled = false;
                    const int keepRecords = 80;
                    var mainApp = (MainWindow)App.Current.MainWindow;

                    Action readFromXBee = () =>
                    {
                        try
                        {
                            _xbeeData.Read(keepRecords, this);
                        }
                        catch (Exception ex)
                        {
                            //there is one exception, the problem is stopped, but the serial still tries to read data. It doesn't happen much, but the condition below ignores when the problem happen.
                            if (!(ex.Message == "The port is closed." && _xbeeData.IsCancelled == true))
                            {
                                //MessageBox.Show("Error reading from serial port: " + ex.Message);
                                Dispose();
                            }
                        }
                    };

                    Task.Factory.StartNew(readFromXBee);
                    ButtonLabel = "Stop Reading";
                    _xbeeData.IsCancelled = false;
                    IsRunning = true;
                }
                else
                {
                    Dispose();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        public void SetAxisLimits(DateTime now)
        {
            Max = now.Ticks + TimeSpan.FromSeconds(0.3).Ticks; // lets force the axis to be 1 second ahead
            Min = now.Ticks - TimeSpan.FromSeconds(1.5).Ticks; // and 8 seconds behind
        }
        #endregion

        #region interface implementation

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected virtual void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }

        public void Dispose()
        {
            _xbeeData.IsCancelled = true;
            EMGValues.Clear();
            _xbeeData.Stop();
            ButtonLabel = "Start Reading";
            BatteryLevel = 0;
            IsRunning = false;
        }
        #endregion
    }
}
