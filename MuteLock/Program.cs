using Microsoft.Win32;
using NAudio.CoreAudioApi;
using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Runtime.InteropServices;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;

namespace MuteLock
{
    class Program
    {
        private static readonly string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs.txt");

        static void Main(string[] args)
        {
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
                        File.Create(logFilePath).Dispose();
                        SetNTFSCompression(logFilePath);
                    }
                    else if (!IsFileCompressed(logFilePath))
                    {
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
                    }
                    catch (Exception ex)
                    {
                        Log($"Error extracting NAudio.Wasapi.dll: {ex.Message}");
                    }
                }

                // Mute audio after DLL extraction to ensure NAudio.Wasapi.dll is available
                try
                {
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

        private static void MuteVolume()
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

        private static void Log(string message)
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

        private static bool IsFileCompressed(string filePath)
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

        private static void SetNTFSCompression(string filePath)
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
}