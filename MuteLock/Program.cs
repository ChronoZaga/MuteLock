using Microsoft.Win32;
using NAudio.CoreAudioApi;
using System;

namespace MuteLock
{
    class Program
    {
        static void Main(string[] args)
        {
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