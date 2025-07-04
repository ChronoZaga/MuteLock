using Microsoft.Win32;
using NAudio.CoreAudioApi;
using System;
using System.IO;
using System.Reflection;
using System.Threading;

namespace MuteLock
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check if another instance is already running
            bool createdNew;
            using (Mutex mutex = new Mutex(true, "MuteLock_SingleInstance", out createdNew))
            {
                if (!createdNew)
                {
                    Console.WriteLine("Another instance of MuteLock is already running. Exiting.");
                    return;
                }

                // Debug output to confirm program starts
                Console.WriteLine("Program starting...");

                // List all embedded resource names for debugging
                Console.WriteLine("Embedded resources:");
                var resourceNames = Assembly.GetExecutingAssembly().GetManifestResourceNames();
                foreach (var name in resourceNames)
                {
                    Console.WriteLine($" - {name}");
                }

                // Extract embedded NAudio.Wasapi.dll if it doesn't exist
                string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NAudio.Wasapi.dll");
                if (!File.Exists(dllPath))
                {
                    try
                    {
                        Console.WriteLine("Attempting to extract NAudio.Wasapi.dll...");
                        using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("MuteLock.NAudio.Wasapi.dll"))
                        {
                            if (resourceStream == null)
                            {
                                Console.WriteLine("Error: Embedded NAudio.Wasapi.dll resource not found.");
                                return;
                            }
                            using (FileStream fileStream = new FileStream(dllPath, FileMode.Create, FileAccess.Write))
                            {
                                resourceStream.CopyTo(fileStream);
                            }
                        }
                        Console.WriteLine("Extracted NAudio.Wasapi.dll to application directory.");
                    }
                    catch (UnauthorizedAccessException)
                    {
                        Console.WriteLine("Error: Permission denied when extracting NAudio.Wasapi.dll. Please run the program as administrator to extract the DLL to Program Files.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error extracting NAudio.Wasapi.dll: {ex.Message}");
                    }
                }

                // Mute audio after DLL extraction to ensure NAudio.Wasapi.dll is available
                try
                {
                    Console.WriteLine("Attempting initial mute...");
                    MuteVolume();
                    Console.WriteLine($"Set volume to 0 and muted at program start at {DateTime.Now}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during initial mute: {ex.Message}");
                }

                Console.WriteLine("Volume Muter started. Press Ctrl+C to exit.");

                // Subscribe to session switch events (lock/unlock)
                SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

                // Subscribe to session end events (log out, shutdown)
                SystemEvents.SessionEnded += SystemEvents_SessionEnded;

                // Keep the console application running
                Console.CancelKeyPress += (s, e) =>
                {
                    SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
                    SystemEvents.SessionEnded -= SystemEvents_SessionEnded;
                    Console.WriteLine("Exiting Volume Muter.");
                };

                // Run until manually stopped
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
                Console.WriteLine($"Set volume to 0 and muted due to {e.Reason} at {DateTime.Now}");
            }
        }

        private static void SystemEvents_SessionEnded(object sender, SessionEndedEventArgs e)
        {
            // Handle log out, shutdown, or restart (session end)
            MuteVolume();
            Console.WriteLine($"Set volume to 0 and muted due to session end ({e.Reason}) at {DateTime.Now}");
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
                Console.WriteLine($"Error setting volume to 0 or muting: {ex.Message}");
            }
        }
    }
}