using System;
using System.Threading.Tasks;
using LiveCharts;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Prism.Commands;
using System.Collections.Generic;
using Syncfusion.XlsIO;
using System.Windows;
using System.IO;
using System.Diagnostics;
using System.Threading;

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
        private XBeeData _xbeeData;
        #endregion

        #region properties
        public DelegateCommand ReadCommand { get; private set; }
        public DelegateCommand SaveCommand { get; private set; }
        public DelegateCommand StopCommand { get; set; }

        public bool IsRunning { get; set; }

        public ChartValues<MeasureModel> EMGValues { get; set; }
        public ChartValues<MeasureModel> ForceValues { get; set; }
        public ChartValues<MeasureModel> AngleValues { get; set; }
        public ChartValues<MeasureModel> AngularVelocityValues { get; set; }

        public List<SessionData> SessionDatas { get; set; }

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

            ReadCommand = new DelegateCommand(
                executeMethod: () =>
                {
                    Read();
                    QueryCanExecute();
                },
                canExecuteMethod: () =>
                {
                    return IsRunning == false;
                });

            SaveCommand = new DelegateCommand(
                executeMethod: () =>
                {
                    SaveData();
                    QueryCanExecute();
                },
                canExecuteMethod: () =>
                {
                    return IsRunning == false &&
                           _xbeeData.IsCancelled == true;
                });

            StopCommand = new DelegateCommand(
                executeMethod: () =>
                {
                    Dispose();
                    QueryCanExecute();
                },
                canExecuteMethod: () =>
                {
                    return IsRunning == true &&
                        _xbeeData.IsCancelled == false;
                });

            ApplicationCommands.ReadCommand.RegisterCommand(ReadCommand);
            ApplicationCommands.SaveCommand.RegisterCommand(SaveCommand);
            ApplicationCommands.StopCommand.RegisterCommand(StopCommand);

            IsRunning = false;

            //For configuring LiveCharts to use MeasureModel for X and Y
            var mapper = LiveCharts.Configurations.Mappers.Xy<MeasureModel>()
                .X(model => model.DateTime.Ticks)   //use DateTime.Ticks as X
                .Y(model => model.Value);           //use the value property as Y

            //lets save the mapper globally.
            Charting.For<MeasureModel>(mapper);

            //the values property will store our values array
            EMGValues = new ChartValues<MeasureModel>();
            ForceValues = new ChartValues<MeasureModel>();
            AngleValues = new ChartValues<MeasureModel>();
            AngularVelocityValues = new ChartValues<MeasureModel>();

            SessionDatas = new List<SessionData>();

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

        #region methods

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

        public void SaveData()
        {
            if (!IsRunning)
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //Create a workbook and enable calculations
                    IWorkbook workbook = application.Workbooks.Create(1);
                    IWorksheet worksheet = workbook.Worksheets[0];
                    worksheet.EnableSheetCalculations();

                    //Import data from SessionDatas
                    ExcelImportDataOptions importDataOptions = new ExcelImportDataOptions();
                    importDataOptions.FirstRow = 1;
                    importDataOptions.FirstColumn = 1;
                    importDataOptions.IncludeHeader = true;
                    importDataOptions.PreserveTypes = false;
                    worksheet.ImportData(SessionDatas, importDataOptions);

                    #region Calculate summary statistics
                    //Set Labels
                    worksheet.Range["G2"].Text = "Min";
                    worksheet.Range["G3"].Text = "Max";
                    worksheet.Range["G4"].Text = "Range";
                    worksheet.Range["G5"].Text = "Mean";
                    worksheet.Range["G6"].Text = "StDev";
                    worksheet.Range["H1"].Text = "Angle";
                    worksheet.Range["I1"].Text = "AngVel";
                    worksheet.Range["J1"].Text = "EMG";
                    worksheet.Range["K1"].Text = "Force";

                    //Angle
                    worksheet.Range["H2"].Formula = "=MIN(B:B)";
                    worksheet.Range["H3"].Formula = "=MAX(B:B)";
                    worksheet.Range["H4"].Formula = "=ABS(H3-H2)";
                    worksheet.Range["H5"].Formula = "=AVERAGE(B:B)";
                    worksheet.Range["H6"].Formula = "=STDEV.P(B:B)";

                    //Angular Velocity
                    worksheet.Range["I2"].Formula = "=MIN(C:C)";
                    worksheet.Range["I3"].Formula = "=MAX(C:C)";
                    worksheet.Range["I4"].Formula = "=ABS(I3-I2)";
                    worksheet.Range["I5"].Formula = "=AVERAGE(C:C)";
                    worksheet.Range["I6"].Formula = "=STDEV.P(C:C)";

                    //EMG
                    worksheet.Range["J2"].Formula = "=MIN(D:D)";
                    worksheet.Range["J3"].Formula = "=MAX(D:D)";
                    worksheet.Range["J4"].Formula = "=ABS(J3-J2)";
                    worksheet.Range["J5"].Formula = "=AVERAGE(D:D)";
                    worksheet.Range["J6"].Formula = "=STDEV.P(D:D)";

                    //Force
                    worksheet.Range["K2"].Formula = "=MIN(E:E)";
                    worksheet.Range["K3"].Formula = "=MAX(E:E)";
                    worksheet.Range["K4"].Formula = "=ABS(K3-K2)";
                    worksheet.Range["K5"].Formula = "=AVERAGE(E:E)";
                    worksheet.Range["K6"].Formula = "=STDEV.P(E:E)";
                    #endregion

                    //Set path and save
                    string spreadsheetNamePath = "acquiredData/";
                    string spreadsheetNameDate = DateTime.Now.ToString("dddd dd MMM y HHmmss");
                    string spreadsheetName = spreadsheetNamePath + spreadsheetNameDate;
                    string path = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName)
                    + "\\acquiredData\\";

                    worksheet.DisableSheetCalculations();
                    workbook.SaveAs(spreadsheetName + ".xlsx");

                    #region View the Workbook
                    //Message box confirmation to view the created document.
                    if (MessageBox.Show("Do you want to view the Excel file?", "Excel file has been created",
                        MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {
                        try
                        {
                            //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                            Process.Start(path+spreadsheetNameDate + ".xlsx");

                            //Exit
                        }
                        catch (Win32Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                    }
                    // Now that the file has been created, delete contents of SessionDatas
                    Thread.Sleep(500);
                    SessionDatas.Clear();
                    #endregion
                }
            }
        }

        public void SetAxisLimits(DateTime now)
        {
            Max = now.Ticks + TimeSpan.FromSeconds(0.3).Ticks; // lets force the axis to be 1 second ahead
            Min = now.Ticks - TimeSpan.FromSeconds(1.5).Ticks; // and 8 seconds behind
        }
        #endregion

        #region interface implementation

        void QueryCanExecute()
        {
            ReadCommand.RaiseCanExecuteChanged();
            StopCommand.RaiseCanExecuteChanged();
            SaveCommand.RaiseCanExecuteChanged();
        }

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
            _xbeeData.Stop();
            EMGValues.Clear();
            ForceValues.Clear();
            AngleValues.Clear();
            AngularVelocityValues.Clear();
            IsRunning = false;
        }
        #endregion
    }
}
