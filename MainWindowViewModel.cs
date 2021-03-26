using Prism.Commands;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace SpasticityClient
{
    class MainWindowViewModel: INotifyPropertyChanged
    {
        private ChartModel _chartModel;
        private string _portName;

        public List<string> PortNames { get; internal set; }
        public DelegateCommand SaveCommand { get; private set; }
        public DelegateCommand StopCommand { get; private set; }
        //public DelegateCommand UpdateCommand { get; private set; }

        public ChartModel ChartModel
        {
            get { return _chartModel; }
            set
            {
                _chartModel = value;
                NotifyPropertyChanged("ChartModel");
            }
        }

        public string PortName
        {
            get { return _portName; }
            set
            {
                _portName = value;
                ChartModel = new ChartModel(_portName);
                NotifyPropertyChanged("PortName");
            }
        }

        public MainWindowViewModel()
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MjQyNDEwQDMxMzgyZTMxMmUzMGNKSzduS2FuZktxRmJERmkySTJsZjVCWlFPeVRGM3pMa1NPYWlVeUttSzA9");

            SaveCommand = new DelegateCommand(SaveData);
            ApplicationCommands.SaveCommand.RegisterCommand(SaveCommand);
            StopCommand = new DelegateCommand(Stop);
            ApplicationCommands.StopCommand.RegisterCommand(StopCommand);

            PortNames = XBeeFunctions.GetPortNamesByBaudrate(57600);

            if (PortNames.Count >= 1)
            {
                //PortName = "";
                if (PortNames.Count == 1)
                    _portName = PortNames[0];
                    //_portName = "COM 7";
            }
            else
            {
                MessageBox.Show("At least one port needed for chart application to work");
                //If no COM port added, then close window.
                MainWindow mainWindow = Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                mainWindow.Close();
            }
        }

        private void Stop()
        {
            ChartModel.Dispose();
        }

        private void SaveData()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //#region Add headers to spreadsheet
                //worksheet.Range["A1"].Text = "TimeStamp";
                //worksheet.Range["B1"].Text = "AngVelX_A";
                //worksheet.Range["C1"].Text = "AngVelY_A";
                //worksheet.Range["D1"].Text = "AngVelZ_A";
                //worksheet.Range["E1"].Text = "OrientX_A";
                //worksheet.Range["F1"].Text = "OrientY_A";
                //worksheet.Range["G1"].Text = "OrientZ_A";
                //worksheet.Range["H1"].Text = "AngVelX_B";
                //worksheet.Range["I1"].Text = "AngVelY_B";
                //worksheet.Range["J1"].Text = "AngVelZ_B";
                //worksheet.Range["K1"].Text = "OrientX_B";
                //worksheet.Range["L1"].Text = "OrientY_B";
                //worksheet.Range["M1"].Text = "OrientZ_B";
                //worksheet.Range["N1"].Text = "EMG";
                //worksheet.Range["O1"].Text = "Force";
                //#endregion

                ExcelImportDataOptions importDataOptions = new ExcelImportDataOptions();
                importDataOptions.FirstRow = 2;
                importDataOptions.FirstColumn = 1;
                importDataOptions.IncludeHeader = true;
                importDataOptions.PreserveTypes = true;

                worksheet.ImportData(ChartModel.SessionDatas, importDataOptions);
                workbook.SaveAs("ImportData.xlsx");

                #region View the Workbook
                //Message box confirmation to view the created document.
                if (MessageBox.Show("Do you want to view the Excel file?", "Excel file has been created",
                    MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    try
                    {
                        //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                        System.Diagnostics.Process.Start("ImportData.xlsx");

                        //Exit
                    }
                    catch (Win32Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }
                else;
                #endregion
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

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
    }
}

