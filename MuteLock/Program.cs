using Microsoft.Win32;
using NAudio.CoreAudioApi;
using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ComponentModel;
using System.ServiceProcess;
using System.Diagnostics;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;

namespace MuteLock
{
    class Program
    {
        private static readonly string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs.txt");
        private static readonly string taskXmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MuteLock.xml");

        static void Main(string[] args)
        {
            if (Environment.UserInteractive)
            {
                if (args.Length > 0)
                {
                    if (string.Equals(args[0], "/installservice", StringComparison.OrdinalIgnoreCase))
                    {
                        InstallService();
                        return;
                    }
                    else if (string.Equals(args[0], "/uninstallservice", StringComparison.OrdinalIgnoreCase))
                    {
                        UninstallService();
                        return;
                    }
                }

                // Original console mode logic
                // Log program start
                Log($"Program started at {DateTime.Now}");

                // Generate Scheduled Task XML
                GenerateTaskXml();

                // Check if another instance is already running
                bool createdNew;
                using (Mutex mutex = new Mutex(true, "MuteLock_SingleInstance", out createdNew))
                {
                    if (!createdNew)
                    {
                        Log("Another instance of MuteLock is already running. Exiting.");
                        return;
                    }

                    // Ensure log file exists and apply NTFS compression if needed
                    try
                    {
                        if (!File.Exists(logFilePath))
                        {
                            Log("Creating log file");
                            File.Create(logFilePath).Dispose();
                            Log("Applying NTFS compression to log file");
                            SetNTFSCompression(logFilePath);
                        }
                        else if (!IsFileCompressed(logFilePath))
                        {
                            Log("Checking log file compression");
                            Log("Applying NTFS compression to log file");
                            SetNTFSCompression(logFilePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Failed to create or compress log file: {ex.Message}");
                    }

                    // Extract embedded NAudio.Wasapi.dll if it doesn't exist
                    string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NAudio.Wasapi.dll");
                    if (!File.Exists(dllPath))
                    {
                        Log($"Attempting to extract NAudio.Wasapi.dll to {dllPath}");
                        try
                        {
                            using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("MuteLock.NAudio.Wasapi.dll"))
                            {
                                if (resourceStream == null)
                                {
                                    Log("Error: Embedded NAudio.Wasapi.dll resource not found.");
                                    return;
                                }
                                using (FileStream fileStream = new FileStream(dllPath, FileMode.Create, FileAccess.Write))
                                {
                                    resourceStream.CopyTo(fileStream);
                                }
                            }
                        }
                        catch (UnauthorizedAccessException)
                        {
                            Log("Error: Permission denied when extracting NAudio.Wasapi.dll. Please run the program as administrator to extract the DLL to Program Files.");
                            return;
                        }
                        catch (Exception ex)
                        {
                            Log($"Error extracting NAudio.Wasapi.dll: {ex.Message}");
                            return;
                        }
                    }

                    // Mute audio after DLL extraction to ensure NAudio.Wasapi.dll is available
                    try
                    {
                        Log("Attempting initial mute");
                        MuteVolume();
                        Log($"Set volume to 0 and muted at program start at {DateTime.Now}");
                    }
                    catch (Exception ex)
                    {
                        Log($"Error during initial mute: {ex.Message}");
                    }

                    // Subscribe to session switch events (lock/unlock)
                    SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

                    // Subscribe to session end events (log out, shutdown)
                    SystemEvents.SessionEnded += SystemEvents_SessionEnded;

                    // Handle application exit
                    AppDomain.CurrentDomain.ProcessExit += (s, e) =>
                    {
                        SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
                        SystemEvents.SessionEnded -= SystemEvents_SessionEnded;
                    };

                    // Keep the application running
                    while (true)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                }
            }
            else
            {
                ServiceBase.Run(new MuteLockService());
            }
        }

        private static void InstallService()
        {
            if (!IsAdministrator())
            {
                Log("Please run the program as administrator to install the service.");
                return;
            }

            string exePath = $"\"{Assembly.GetExecutingAssembly().Location}\"";
            string serviceName = "MuteLock";

            try
            {
                RunCommand("sc", $"create {serviceName} binpath= {exePath} start= auto");
                RunCommand("sc", $"description {serviceName} \"Mutes audio on lock, unlock, and session end events.\"");
                RunCommand("sc", $"start {serviceName}");
                Log("Service installed and started successfully.");
            }
            catch (Exception ex)
            {
                Log($"Failed to install service: {ex.Message}");
            }
        }

        private static void UninstallService()
        {
            if (!IsAdministrator())
            {
                Log("Please run the program as administrator to uninstall the service.");
                return;
            }

            string serviceName = "MuteLock";

            try
            {
                RunCommand("sc", $"stop {serviceName}");
                RunCommand("sc", $"delete {serviceName}");
                Log("Service uninstalled successfully.");
            }
            catch (Exception ex)
            {
                Log($"Failed to uninstall service: {ex.Message}");
            }
        }

        private static void RunCommand(string fileName, string arguments)
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                throw new Exception($"Command '{fileName} {arguments}' failed with error: {error}");
            }

            Log($"Command output: {output}");
        }

        private static bool IsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private static void GenerateTaskXml()
        {
            try
            {
                // Check if XML already exists
                if (File.Exists(taskXmlPath))
                {
                    Log("MuteLock.xml already exists. Delete the existing file to create a new one.");
                    return;
                }

                string exePath = Assembly.GetExecutingAssembly().Location;
                string currentUserSid = WindowsIdentity.GetCurrent().User.Value;
                string currentDateTime = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
                string xmlContent = $@"<?xml version=""1.0"" encoding=""UTF-16""?>
<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">
  <RegistrationInfo>
    <Date>{currentDateTime}</Date>
    <Author>ChronoZaga</Author>
    <URI>\MuteLock</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{currentUserSid}</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id=""Author"">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context=""Author"">
    <Exec>
      <Command>""{exePath}""</Command>
    </Exec>
  </Actions>
</Task>";

                // Write the XML file
                File.WriteAllText(taskXmlPath, xmlContent);
                Log($"Successfully created Scheduled Task XML at: {taskXmlPath}");
            }
            catch (Exception ex)
            {
                Log($"Failed to create Scheduled Task XML: {ex.Message}");
            }
        }

        private static void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            // Handle lock and unlock events
            if (e.Reason == SessionSwitchReason.SessionLock || e.Reason == SessionSwitchReason.SessionUnlock)
            {
                MuteVolume();
                Log($"Set volume to 0 and muted due to {e.Reason} at {DateTime.Now}");
            }
        }

        private static void SystemEvents_SessionEnded(object sender, SessionEndedEventArgs e)
        {
            // Handle log out, shutdown, or restart (session end)
            MuteVolume();
            Log($"Set volume to 0 and muted due to session end ({e.Reason}) at {DateTime.Now}");
        }

        public static void MuteVolume()
        {
            try
            {
                // Use NAudio to set volume to 0 and mute
                using (var deviceEnumerator = new MMDeviceEnumerator())
                {
                    var device = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    device.AudioEndpointVolume.MasterVolumeLevelScalar = 0f; // Set volume to 0%
                    device.AudioEndpointVolume.Mute = true; // Mute the volume
                }
            }
            catch (Exception ex)
            {
                Log($"Error setting volume to 0 or muting: {ex.Message}");
            }
        }

        public static void Log(string message)
        {
            try
            {
                File.AppendAllText(logFilePath, $"[{DateTime.Now}] {message}{Environment.NewLine}");
            }
            catch
            {
                // Silently ignore logging errors to avoid creating any additional files
            }
        }

        // NTFS Compression methods
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool DeviceIoControl(
            IntPtr hDevice,
            uint dwIoControlCode,
            IntPtr lpInBuffer,
            uint nInBufferSize,
            IntPtr lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileInformationByHandle(
            IntPtr hFile,
            out BY_HANDLE_FILE_INFORMATION lpFileInformation);

        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;
        private const uint FSCTL_SET_COMPRESSION = 0x9C040;
        private const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;
        private const short COMPRESSION_FORMAT_DEFAULT = 1;
        private const uint FILE_ATTRIBUTE_COMPRESSED = 0x800;

        [StructLayout(LayoutKind.Sequential)]
        private struct BY_HANDLE_FILE_INFORMATION
        {
            public uint FileAttributes;
            public FILETIME CreationTime;
            public FILETIME LastAccessTime;
            public FILETIME LastWriteTime;
            public uint VolumeSerialNumber;
            public uint FileSizeHigh;
            public uint FileSizeLow;
            public uint NumberOfLinks;
            public uint FileIndexHigh;
            public uint FileIndexLow;
        }

        public static bool IsFileCompressed(string filePath)
        {
            try
            {
                IntPtr handle = CreateFile(
                    filePath,
                    GENERIC_READ,
                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    FILE_FLAG_BACKUP_SEMANTICS,
                    IntPtr.Zero);

                if (handle == IntPtr.Zero || handle.ToInt64() == -1)
                {
                    Log($"Failed to open file to check compression: {Marshal.GetLastWin32Error()}");
                    return false;
                }

                try
                {
                    if (!GetFileInformationByHandle(handle, out BY_HANDLE_FILE_INFORMATION fileInfo))
                    {
                        Log($"Failed to get file information for compression check: {Marshal.GetLastWin32Error()}");
                        return false;
                    }

                    return (fileInfo.FileAttributes & FILE_ATTRIBUTE_COMPRESSED) != 0;
                }
                finally
                {
                    CloseHandle(handle);
                }
            }
            catch (Exception ex)
            {
                Log($"Error checking compression status: {ex.Message}");
                return false;
            }
        }

        public static void SetNTFSCompression(string filePath)
        {
            IntPtr handle = IntPtr.Zero;
            GCHandle pinnedBuffer = default;
            try
            {
                handle = CreateFile(
                    filePath,
                    GENERIC_READ | GENERIC_WRITE,
                    FILE_SHARE_READ | FILE_SHARE_WRITE,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    0,
                    IntPtr.Zero);

                if (handle == IntPtr.Zero || handle.ToInt64() == -1)
                {
                    Log($"Failed to open file for compression: {Marshal.GetLastWin32Error()}");
                    return;
                }

                uint bytesReturned;
                short compressionState = COMPRESSION_FORMAT_DEFAULT;
                pinnedBuffer = GCHandle.Alloc(compressionState, GCHandleType.Pinned);
                bool result = DeviceIoControl(
                    handle,
                    FSCTL_SET_COMPRESSION,
                    pinnedBuffer.AddrOfPinnedObject(),
                    sizeof(short),
                    IntPtr.Zero,
                    0,
                    out bytesReturned,
                    IntPtr.Zero);

                if (!result)
                {
                    Log($"Failed to set compression: {Marshal.GetLastWin32Error()}");
                    return;
                }
            }
            catch (Exception ex)
            {
                Log($"Error applying NTFS compression: {ex.Message}");
            }
            finally
            {
                if (pinnedBuffer.IsAllocated)
                {
                    pinnedBuffer.Free();
                }
                if (handle != IntPtr.Zero && handle.ToInt64() != -1)
                {
                    CloseHandle(handle);
                }
            }
        }
    }

    public class MuteLockService : ServiceBase
    {
        private static readonly string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs.txt");

        public MuteLockService()
        {
            ServiceName = "MuteLock";
            CanHandleSessionChangeEvent = true;
        }

        protected override void OnStart(string[] args)
        {
            Program.Log($"Service started at {DateTime.Now}");

            // Ensure log file exists and apply NTFS compression if needed
            try
            {
                if (!File.Exists(logFilePath))
                {
                    Program.Log("Creating log file");
                    File.Create(logFilePath).Dispose();
                    Program.Log("Applying NTFS compression to log file");
                    Program.SetNTFSCompression(logFilePath);
                }
                else if (!Program.IsFileCompressed(logFilePath))
                {
                    Program.Log("Checking log file compression");
                    Program.Log("Applying NTFS compression to log file");
                    Program.SetNTFSCompression(logFilePath);
                }
            }
            catch (Exception ex)
            {
                Program.Log($"Failed to create or compress log file: {ex.Message}");
            }

            // Extract embedded NAudio.Wasapi.dll if it doesn't exist
            string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NAudio.Wasapi.dll");
            if (!File.Exists(dllPath))
            {
                Program.Log($"Attempting to extract NAudio.Wasapi.dll to {dllPath}");
                try
                {
                    using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("MuteLock.NAudio.Wasapi.dll"))
                    {
                        if (resourceStream == null)
                        {
                            Program.Log("Error: Embedded NAudio.Wasapi.dll resource not found.");
                            return;
                        }
                        using (FileStream fileStream = new FileStream(dllPath, FileMode.Create, FileAccess.Write))
                        {
                            resourceStream.CopyTo(fileStream);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Program.Log($"Error extracting NAudio.Wasapi.dll: {ex.Message}");
                    return;
                }
            }

            // Mute audio after DLL extraction to ensure NAudio.Wasapi.dll is available
            try
            {
                Program.Log("Attempting initial mute");
                Program.MuteVolume();
                Program.Log($"Set volume to 0 and muted at service start at {DateTime.Now}");
            }
            catch (Exception ex)
            {
                Program.Log($"Error during initial mute: {ex.Message}");
            }
        }

        protected override void OnStop()
        {
            Program.Log($"Service stopped at {DateTime.Now}");
        }

        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            if (changeDescription.Reason == SessionChangeReason.SessionLock ||
                changeDescription.Reason == SessionChangeReason.SessionUnlock ||
                changeDescription.Reason == SessionChangeReason.SessionLogoff)
            {
                Program.MuteVolume();
                Program.Log($"Set volume to 0 and muted due to {changeDescription.Reason} (Session {changeDescription.SessionId}) at {DateTime.Now}");
            }
        }

        protected override void OnShutdown()
        {
            Program.MuteVolume();
            Program.Log($"Set volume to 0 and muted due to system shutdown at {DateTime.Now}");
            base.OnShutdown();
        }
    }
}