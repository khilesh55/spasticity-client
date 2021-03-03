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
    public class ChartModel : INotifyPropertyChanged, IDisposable
    {
        #region variables
        private double _min;
        private double _count;
        private XBeeData _xbeeData;
        private string _buttonLabel;
        private string _channelName;
        private float _batteryLevel;
        #endregion

        #region properties
        public Queue<string> csvData { get; set; }
        public ChartValues<double> EMGValues { get; set; }
        public RelayCommand ReadCommand { get; set; }
        public bool IsRunning { get; set; }
        public double Min
        {
            get { return _min; }
            set
            {
                _min = value;
                OnPropertyChanged("Min");
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

            EMGValues = new ChartValues<double>();
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
                    const int keepRecords = 20;
                    var mainApp = (MainWindow)App.Current.MainWindow;

                    Action readFromXBee = () =>
                    {
                        try
                        {
#if DEBUG
                            //while (!_xbeeData.IsCancelled)
                            //{
                            //    Random random = new Random();
                            //    double force = random.NextDouble() * 6 - 3;
                            //    this.Values.Add(force);
                            //    this.Min = this.Values.Count - keepRecords;
                            //    if (this.Values.Count > keepRecords) this.Values.RemoveAt(0);
                            //}
#endif
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
