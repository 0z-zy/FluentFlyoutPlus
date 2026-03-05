using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using Windows.Media.Control;

namespace FluentFlyoutWPF.Classes
{
    public class ExtensionMediaData
    {
        public string title { get; set; } = string.Empty;
        public string artist { get; set; } = string.Empty;
        public double duration { get; set; } // in seconds
        public double progress { get; set; } // in seconds
        public bool playing { get; set; }
        public DateTime LastUpdated { get; set; } = DateTime.MinValue;
    }

    public class MergedMediaInfo
    {
        public string Title { get; set; } = string.Empty;
        public string Artist { get; set; } = string.Empty;
        public TimeSpan Duration { get; set; }
        public TimeSpan Position { get; set; }
        public bool IsPlaying { get; set; }
        public bool IsExtensionActive { get; set; }
    }

    public class MediaSyncManager
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private HttpListener _listener;
        private ExtensionMediaData _latestExtensionData = new ExtensionMediaData();
        private DateTime _lastExtensionActiveTime = DateTime.MinValue;
        private DateTime _stickyLockUntil = DateTime.MinValue;
        private string _stickyTitle = string.Empty;
        private string _stickyArtist = string.Empty;

        private const int Port = 8888;
        private const int InactiveTimeoutSeconds = 5;
        private const int StickyLockSeconds = 5;
        private const int JitterThresholdMs = 2000;

        public event EventHandler<MergedMediaInfo>? OnMediaDataUpdated;

        public MediaSyncManager()
        {
            _listener = new HttpListener();
            _listener.Prefixes.Add($"http://127.0.0.1:{Port}/extension_data/");
        }

        public void Start()
        {
            try
            {
                _listener.Start();
                Task.Run(() => ListenLoop());
                Logger.Info($"MediaSyncManager: Listening on port {Port}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "MediaSyncManager: Failed to start listener");
            }
        }

        private async Task ListenLoop()
        {
            while (_listener.IsListening)
            {
                try
                {
                    var context = await _listener.GetContextAsync();
                    if (context.Request.HttpMethod == "POST" && context.Request.Url?.AbsolutePath == "/extension_data")
                    {
                        ProcessRequest(context);
                    }
                    else if (context.Request.HttpMethod == "OPTIONS")
                    {
                        HandleCors(context);
                    }
                    else
                    {
                        context.Response.StatusCode = (int)HttpStatusCode.NotFound;
                        context.Response.Close();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, "MediaSyncManager: Error in ListenLoop");
                }
            }
        }

        private void HandleCors(HttpListenerContext context)
        {
            context.Response.Headers.Add("Access-Control-Allow-Origin", "*");
            context.Response.Headers.Add("Access-Control-Allow-Methods", "POST, OPTIONS");
            context.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
            context.Response.StatusCode = (int)HttpStatusCode.NoContent;
            context.Response.Close();
        }

        private void ProcessRequest(HttpListenerContext context)
        {
            try
            {
                using (var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
                {
                    string json = reader.ReadToEnd();
                    var data = JsonSerializer.Deserialize<ExtensionMediaData>(json);
                    if (data != null)
                    {
                        data.LastUpdated = DateTime.Now;
                        _latestExtensionData = data;
                        _lastExtensionActiveTime = DateTime.Now;
                        
                        // Update sticky names whenever we get fresh data from extension
                        if (data.playing)
                        {
                            _stickyTitle = data.title;
                            _stickyArtist = data.artist;
                            _stickyLockUntil = DateTime.MinValue; // Clear lock if playing
                        }
                        else
                        {
                            // If stopped, start the sticky lock
                            if (_stickyLockUntil == DateTime.MinValue)
                            {
                                _stickyLockUntil = DateTime.Now.AddSeconds(StickyLockSeconds);
                            }
                        }

                        OnMediaDataUpdated?.Invoke(this, GetMergedData("-", "-", TimeSpan.Zero, TimeSpan.Zero, false, TimeSpan.Zero));
                    }
                }

                context.Response.Headers.Add("Access-Control-Allow-Origin", "*");
                context.Response.StatusCode = (int)HttpStatusCode.OK;
                context.Response.Close();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "MediaSyncManager: Error processing request");
                context.Response.StatusCode = (int)HttpStatusCode.BadRequest;
                context.Response.Close();
            }
        }

        public MergedMediaInfo GetMergedData(
            string smtcTitle, 
            string smtcArtist, 
            TimeSpan smtcDuration, 
            TimeSpan smtcPosition, 
            bool smtcPlaying,
            TimeSpan localElapsedSinceSMTCUpdate)
        {
            bool extensionActive = (DateTime.Now - _lastExtensionActiveTime).TotalSeconds < InactiveTimeoutSeconds;
            
            var merged = new MergedMediaInfo();
            merged.IsExtensionActive = extensionActive;

            // 1. Names Priority: Extension wins if active OR if in sticky lock
            bool useStickyNames = DateTime.Now < _stickyLockUntil;
            
            if (extensionActive && !string.IsNullOrEmpty(_latestExtensionData.title))
            {
                merged.Title = _latestExtensionData.title;
                merged.Artist = _latestExtensionData.artist;
            }
            else if (useStickyNames && !string.IsNullOrEmpty(_stickyTitle))
            {
                merged.Title = _stickyTitle;
                merged.Artist = _stickyArtist;
            }
            else
            {
                merged.Title = smtcTitle;
                merged.Artist = smtcArtist;
            }

            // 2. Timeline Priority: SMTC wins if valid (duration > 0)
            // Spotify fix: If SMTC is playing but reports 0 duration, it might be in transition.
            // If we have extension data, fallback to it. 
            // If we don't, we have to stick with SMTC but it might look like 0:00.
            
            if (smtcDuration.TotalSeconds > 1) // Using > 1s because some apps report 0.001 or such for "infinite" or "unknown"
            {
                merged.Duration = smtcDuration;
                
                // Jitter-free sync for position
                TimeSpan currentLocalPosition = smtcPosition + localElapsedSinceSMTCUpdate;
                
                // Safety bound for position
                if (currentLocalPosition > smtcDuration) currentLocalPosition = smtcDuration;
                if (currentLocalPosition < TimeSpan.Zero) currentLocalPosition = TimeSpan.Zero;

                merged.Position = currentLocalPosition;

                if (extensionActive)
                {
                    TimeSpan extensionPosition = TimeSpan.FromSeconds(_latestExtensionData.progress);
                    // If drift is too large, it might be because Extension is more accurate or SMTC is stuck
                    if (Math.Abs((extensionPosition - currentLocalPosition).TotalMilliseconds) > JitterThresholdMs)
                    {
                        // Note: We prioritize SMTC timeline if valid, so we only snap if drift is EXTREME (e.g. > 5s)
                        // Or if extension says it's playing but SMTC is stuck.
                        if (Math.Abs((extensionPosition - currentLocalPosition).TotalSeconds) > 5)
                        {
                            merged.Position = extensionPosition;
                        }
                    }
                }
            }
            else if (extensionActive && _latestExtensionData.duration > 0)
            {
                merged.Duration = TimeSpan.FromSeconds(_latestExtensionData.duration);
                merged.Position = TimeSpan.FromSeconds(_latestExtensionData.progress) + (DateTime.Now - _latestExtensionData.LastUpdated);
            }
            else
            {
                // Last resort: use whatever SMTC gave us
                merged.Duration = smtcDuration;
                merged.Position = smtcPosition + localElapsedSinceSMTCUpdate;
            }

            merged.IsPlaying = extensionActive ? _latestExtensionData.playing : smtcPlaying;

            return merged;
        }
    }
}
