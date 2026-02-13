/*namespace MSTeamsIndicator
{
    public class Worker(ILogger<Worker> logger) : BackgroundService
    {
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                if (logger.IsEnabled(LogLevel.Information))
                {
                    logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                }
                await Task.Delay(1000, stoppingToken);
            }
        }
    }
}*/
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System;
using System.Data;
using System.Diagnostics;
using System.IO.Ports;

namespace MSTeamsIndicator
{
   public class Worker : BackgroundService
   {
      private readonly IConfiguration _config;
      private readonly ILogger<Worker> _logger;
      private SerialPort _serialPort;
      private enum CallState { Active, Unknown, Inactive};

      public Worker(IConfiguration config, ILogger<Worker> logger)
      {
         _config = config;
         _logger = logger;
      }

      protected override async Task ExecuteAsync(CancellationToken stoppingToken)
      {
         _serialPort = null;
         _logger.LogInformation("Worker started");
         try
         {
            var portName = _config["SerialPort:PortName"];
            var baudRate = int.Parse(_config["SerialPort:BaudRate"]);
            _serialPort = new SerialPort(portName, baudRate);
            _serialPort.Open();
            _logger.LogInformation("Opened serial port {PortName} at {BaudRate} baud", portName, baudRate);
               _serialPort.Write("Start: set color settings: " + _config["Indicator:ColorOff"] + "\n");
               await Task.Delay(100, stoppingToken);
               _serialPort.Write("Start: teams call started\n");
         }
         catch (Exception exception)
         {
            _logger.LogError(exception, "Unhandled error in worker loop");
         }

         while (!stoppingToken.IsCancellationRequested)
         {
            try
            {
               if (_serialPort != null && _serialPort.IsOpen)
               {
                  CallState callState = IsTeamsCallActive();
                  if (callState != CallState.Unknown)
                  {
                     var colorString = callState == CallState.Active ? _config["Indicator:ColorOn"] : _config["Indicator:ColorOff"];
                     try
                     {
                        _serialPort.Write("Start: set color settings: " + colorString + "\n");
                        await Task.Delay(100, stoppingToken);
                        _serialPort.Write("Start: teams call started\n");
                     }
                     catch (Exception exception)
                     {
                        _logger.LogError(exception, "Failed to write serial port");
                     }
                  }
                  await Task.Delay(int.Parse(_config["PollingIntervalMs"]), stoppingToken);
               }
               else
               {
                  await Task.Delay(5000, stoppingToken);
                  try
                  {
                     var portName = _config["SerialPort:PortName"];
                     var baudRate = int.Parse(_config["SerialPort:BaudRate"]);
                     _serialPort = new SerialPort(portName, baudRate);
                     _serialPort.Open();
                     _logger.LogInformation("Opened serial port {PortName} at {BaudRate} baud", portName, baudRate);
                  }
                  catch (FileNotFoundException exception)
                  {
                     _logger.LogError(exception, "Failed to open serial port");
                  }
               }
            }
            catch (Exception exception)
            {
               _logger.LogError(exception, "Unhandled error in worker loop");
            }
         }
         _serialPort?.Close();
      }/*
      private CallState IsTeamsCallActive()
      {
         var deviceEnumerator = new MMDeviceEnumerator();
         var device = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);

         var sessions = device.AudioSessionManager.Sessions;
         _logger.LogInformation("Sessions: " + sessions.Count.ToString());

         bool validCheck = false;

         for (int i = 0; i < sessions.Count; i++)
         {
            var session = sessions[i];
            var processId = session.GetProcessID;

            try
            {
               var process = Process.GetProcessById((int)processId);
               _logger.LogInformation("Session " + i + ": " + process.ProcessName);

               if (process.ProcessName.Contains("ms-teams", StringComparison.OrdinalIgnoreCase) ||
                   process.ProcessName.Contains("teams", StringComparison.OrdinalIgnoreCase))
               {
                  validCheck = true;
                  if (session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive)
                     return CallState.Active;
               }
            }
            catch (Exception exception)
            {
               _logger.LogError(exception, "Unhandled error in IsTeamsCallActive()");
            }
         }

         if (validCheck)
         {
            return CallState.Inactive;
         }
         else
         {
            return CallState.Unknown;
         }

      }*/
      
      private CallState IsTeamsCallActive()
      {
         bool userNameConfigured = false;
         try
         {
            foreach (var userDir in Directory.GetDirectories(@"C:\Users"))
            {
               var userName = _config["userFolder"];
               string dir = userDir;
               if ("Invalid".CompareTo(userName) == 0)
               {
                  userName = Path.GetFileName(userDir);
                  if (userName is "Public" or "Default" or "Default User" or "All Users")
                     continue;
               }
               else
               {
                  userNameConfigured = true;
                  dir = @"C:\Users" + "\\" + userName;
               }

               // New Teams (ms-teams) log location
               var logDir = Path.Combine(dir,
                  @"AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs");

               _logger.LogTrace("Checking: " + logDir);
               if (!Directory.Exists(logDir))
                  continue;

               var latestLog = Directory.GetFiles(logDir, "MSTeams_*.log")
                  .OrderByDescending(File.GetLastWriteTimeUtc)
                  .FirstOrDefault();

               if (latestLog == null)
               {
                  continue;
               }
               else
               {
                  _logger.LogDebug("Latest log: " + latestLog);
                  CallState callState = IsCallActiveInLogFile(latestLog);
                  if (callState != CallState.Unknown)
                  {
                     return callState;
                  }
               }
               if (userNameConfigured)
               {
                  break;
               }
            }
         }
         catch (Exception ex)
         {
            _logger.LogError(ex, "Error scanning for Teams call activity");
         }

         return CallState.Unknown;
      }

      private CallState IsCallActiveInLogFile(string logFilePath)
      {
         try
         {
            using var fs = new FileStream(logFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // Read the last 64KB to find the most recent call state transition
            var tailSize = (int)Math.Min(fs.Length, 100000);
            fs.Seek(-tailSize, SeekOrigin.End);

            using var reader = new StreamReader(fs);
            var content = reader.ReadToEnd();

            // Teams logs record call state transitions with these markers
            var lastStarted = content.LastIndexOf("CallConnected", StringComparison.OrdinalIgnoreCase);

            var lastEnded = content.LastIndexOf("CallEnded", StringComparison.OrdinalIgnoreCase);

            // Also check presence-based indicators
            var lastInCall = Math.Max(
               content.LastIndexOf("InACall", StringComparison.OrdinalIgnoreCase),
               content.LastIndexOf("InAMeeting", StringComparison.OrdinalIgnoreCase));

            var lastAvailable = Math.Max(
               content.LastIndexOf("\"Available\"", StringComparison.OrdinalIgnoreCase),
               content.LastIndexOf("\"Away\"", StringComparison.OrdinalIgnoreCase));
            if (lastStarted == -1 && lastEnded == -1 && lastInCall == -1 && lastAvailable == -1)
            {
               return CallState.Unknown;
            }
            else if (Math.Max(lastStarted, lastInCall) > Math.Max(lastEnded, lastAvailable))
            {
               return CallState.Active;
            }
            else
            {
               return CallState.Inactive;
            }
         }
         catch (Exception ex)
         {
            _logger.LogError(ex, "Error reading Teams log {LogFile}", logFilePath);
            return CallState.Unknown;
         }
      }
      
      public override void Dispose()
      {
         _serialPort?.Close();
         base.Dispose();
         _logger.LogInformation("Worker stopped");
      }
   }
}