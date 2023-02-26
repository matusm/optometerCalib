using At.Matus.StatisticPod;
using Bev.Instruments.P9710;
using System;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace optometerCalib
{
    class Program
    {
        static int Main(string[] args)
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

            string appName = Assembly.GetExecutingAssembly().GetName().Name;
            var appVersion = Assembly.GetExecutingAssembly().GetName().Version;
            string appVersionString = $"{appVersion.Major}.{appVersion.Minor}";
            Options options = new Options();
            if (!CommandLine.Parser.Default.ParseArgumentsStrict(args, options))
                Console.WriteLine("*** ParseArgumentsStrict returned false");


            DateTime timeStamp = DateTime.UtcNow;
            P9710 device = new P9710(options.Port);
            StreamWriter logWriter = new StreamWriter(options.LogFileName + ".log", true);
            StreamWriter csvWriter = new StreamWriter(options.LogFileName + ".csv", false);
            StatisticPod stpCurrent = new StatisticPod("Current in A");

            if (options.MaximumSamples < 2) options.MaximumSamples = 2;
            if (string.IsNullOrWhiteSpace(options.UserComment))
            {
                options.UserComment = "---";
            }

            device.DeselectAutoRange();
            device.SetMeasurementRange(MeasurementRange.Range03);

            DisplayOnly("");
            LogOnly(fatSeparator);
            DisplayOnly($"Application:     {appName} {appVersionString}");
            LogOnly($"Application:     {appName} {appVersion}");
            LogAndDisplay($"StartTimeUTC:    {timeStamp:dd-MM-yyyy HH:mm}");
            LogAndDisplay($"InstrumentManu:  {device.InstrumentManufacturer}");
            LogAndDisplay($"InstrumentType:  {device.InstrumentType}");
            LogAndDisplay($"InstrumentSN:    {device.InstrumentSerialNumber}");
            LogAndDisplay($"RS232 port:      {device.DevicePort}");
            LogAndDisplay($"Samples (n):     {options.MaximumSamples}");
            LogAndDisplay($"Comment:         {options.UserComment}");
            LogOnly(fatSeparator);
            DisplayOnly("");
            CsvLog(CsvHeader());

            int measurementIndex = 0;
            bool shallLoop = true;
            while (shallLoop)
            {
                DisplayOnly("press any key to start a measurement - 'q' to quit, arrow keys to change range");
                ConsoleKeyInfo cki = Console.ReadKey(true);
                switch (cki.Key)
                {
                    case ConsoleKey.Q:
                        shallLoop = false;
                        DisplayOnly("bye.");
                        break;
                    case ConsoleKey.DownArrow:
                    case ConsoleKey.PageDown:
                        var r = device.GetMeasurementRange();
                        device.SetMeasurementRange(r.Decrement());
                        DisplayCurrentRange();
                        break;
                    case ConsoleKey.UpArrow:
                    case ConsoleKey.PageUp:
                        var s = device.GetMeasurementRange();
                        device.SetMeasurementRange(s.Increment());
                        DisplayCurrentRange();
                        break;
                    default:
                        int iterationIndex = 0;
                        measurementIndex++;
                        DisplayOnly("");
                        DisplayOnly($"Measurement #{measurementIndex} at {device.GetMeasurementRange()}");
                        RestartValues();
                        timeStamp = DateTime.UtcNow;

                        while (iterationIndex < options.MaximumSamples)
                        {
                            iterationIndex++;
                            double current = device.GetCurrent();
                            UpdateValues(current);
                            DisplayOnly($"{iterationIndex,4}:  {current * 1e9:F3} nA");
                        }

                        DisplayOnly("");
                        LogOnly($"Measurement number:   {measurementIndex} ({device.GetMeasurementRange()})");
                        LogOnly($"Triggered at:         {timeStamp:dd-MM-yyyy HH:mm:ss}");
                        //LogAndDisplay($"current:              {stpCurrent.AverageValue.ToString("0.000E0")} ± {stpCurrent.StandardDeviation.ToString("0.000E0")} A");
                        LogAndDisplay($"Actual sample size:   {stpCurrent.SampleSize}");
                        LogAndDisplay($"Current:              {stpCurrent.AverageValue * 1e9:F3} ± {stpCurrent.StandardDeviation * 1e9:F3} nA");
                        LogOnly(thinSeparator);
                        DisplayOnly("");
                        CsvLog(CsvLine(measurementIndex));
                        break;
                }
            }

            logWriter.Close();
            csvWriter.Close();
            device.SelectAutoRange();
            return 0;

            /***************************************************/
            void LogAndDisplay(string line)
            {
                DisplayOnly(line);
                LogOnly(line);
            }
            /***************************************************/
            void LogOnly(string line)
            {
                logWriter.WriteLine(line);
                logWriter.Flush();
            }
            /***************************************************/
            void DisplayOnly(string line)
            {
                Console.WriteLine(line);
            }
            /***************************************************/
            void RestartValues()
            {
                stpCurrent.Restart();
            }
            /***************************************************/
            void UpdateValues(double x)
            {
                if (double.IsInfinity(x))
                    x = double.NaN;
                stpCurrent.Update(x);
            }
            /***************************************************/
            void DisplayCurrentRange()
            {
                DisplayOnly("");
                DisplayOnly($"Current measurement range: {device.GetMeasurementRange()}");
                DisplayOnly("");
            }
            /***************************************************/
            string CsvHeader() => $"measurement number, range, specification (A), measured current (A), standard deviation (A), test current (A), standard uncertainty (A)";
            /***************************************************/
            string CsvLine(int index) => $"{index}, {device.GetMeasurementRange()}, {device.GetSpecification(stpCurrent.AverageValue, device.GetMeasurementRange())}, {stpCurrent.AverageValue}, {stpCurrent.StandardDeviation}, {"[TestCurrent]"}, {"[u(TestCurrent)]"}";
            /***************************************************/
            void CsvLog(string line)
            {
                csvWriter.WriteLine(line);
                csvWriter.Flush();
            }
            /***************************************************/
        }

        private static readonly string fatSeparator = new string('=', 80);
        private static readonly string thinSeparator = new string('-', 80);
    }
}
