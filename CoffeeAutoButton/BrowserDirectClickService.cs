using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using static CoffeeAutoButton.NativeMethods;

namespace CoffeeAutoButton
{
    internal sealed class BrowserClickTarget
    {
        internal int Port { get; set; }

        internal string Id { get; set; } = string.Empty;

        internal string Title { get; set; } = string.Empty;

        internal string Url { get; set; } = string.Empty;

        internal string WebSocketDebuggerUrl { get; set; } = string.Empty;

        internal bool IsRecognized => Port > 0 && !string.IsNullOrWhiteSpace(WebSocketDebuggerUrl);

        internal string DisplayName
        {
            get
            {
                var title = string.IsNullOrWhiteSpace(Title) ? "タイトルなし" : Title;
                return $"CDP:{Port} / {title}";
            }
        }

        internal void CopyFrom(BrowserClickTarget other)
        {
            Port = other.Port;
            Id = other.Id;
            Title = other.Title;
            Url = other.Url;
            WebSocketDebuggerUrl = other.WebSocketDebuggerUrl;
        }
    }

    internal sealed class BrowserDirectClickService
    {
        private static readonly int[] DebugPorts = { 9223, 9222, 9224, 9225 };
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };

        private const int ProbeTimeoutMs = 180;
        private const int ProbeBackoffMs = 5000;

        private DateTime _lastProbeFailedAtUtc = DateTime.MinValue;
        private int? _lastWorkingPort;

        internal async Task<BrowserClickTarget> TryResolveTargetAsync(
            string targetTitle,
            string processName,
            int preferredPort,
            CancellationToken cancellationToken)
        {
            if (!IsChromiumProcess(processName) && string.IsNullOrWhiteSpace(targetTitle))
            {
                return null;
            }

            if (DateTime.UtcNow - _lastProbeFailedAtUtc < TimeSpan.FromMilliseconds(ProbeBackoffMs))
            {
                return null;
            }

            foreach (var port in GetProbePorts(preferredPort))
            {
                var target = await FindTargetAsync(port, targetTitle, processName, cancellationToken);
                if (target is null)
                {
                    continue;
                }

                _lastWorkingPort = port;
                return target;
            }

            _lastProbeFailedAtUtc = DateTime.UtcNow;
            return null;
        }

        internal async Task<bool> TryClickAsync(
            BrowserClickTarget recognizedTarget,
            POINT clientPoint,
            MouseClickAction action,
            int holdDurationMs,
            CancellationToken cancellationToken)
        {
            if (recognizedTarget is null || !recognizedTarget.IsRecognized)
            {
                return false;
            }

            var target = await RefreshTargetAsync(recognizedTarget, cancellationToken) ?? recognizedTarget;

            try
            {
                await using var session = await CdpPageSession.ConnectAsync(target.WebSocketDebuggerUrl, cancellationToken);
                var point = await NormalizeClientPointAsync(session, clientPoint, cancellationToken);
                await DispatchClickAsync(session, point, action, holdDurationMs, cancellationToken);
                recognizedTarget.CopyFrom(target);
                _lastWorkingPort = target.Port;
                return true;
            }
            catch
            {
                // CDP direct click is strict: if the recognized target cannot be reached, do not click elsewhere.
            }

            return false;
        }

        private IEnumerable<int> GetProbePorts(int preferredPort)
        {
            if (preferredPort > 0)
            {
                yield return preferredPort;
            }

            if (_lastWorkingPort is int rememberedPort)
            {
                if (rememberedPort != preferredPort)
                {
                    yield return rememberedPort;
                }
            }

            foreach (var port in DebugPorts)
            {
                if (port != preferredPort && port != _lastWorkingPort)
                {
                    yield return port;
                }
            }
        }

        private async Task<BrowserClickTarget> FindTargetAsync(
            int port,
            string targetTitle,
            string processName,
            CancellationToken cancellationToken)
        {
            var targets = await GetTargetsAsync(port, cancellationToken);
            if (targets.Count == 0)
            {
                return null;
            }

            var pageTargets = targets
                .Where(target =>
                    string.Equals(target.Type, "page", StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(target.WebSocketDebuggerUrl)
                    && !target.Url.StartsWith("devtools://", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (pageTargets.Count == 0)
            {
                return null;
            }

            var normalizedTitle = NormalizeTitle(targetTitle);
            if (!string.IsNullOrWhiteSpace(normalizedTitle))
            {
                var matched = pageTargets.FirstOrDefault(target =>
                    IsTitleMatch(NormalizeTitle(target.Title), normalizedTitle));
                if (matched is not null)
                {
                    return CreateClickTarget(port, matched);
                }
            }

            return IsChromiumProcess(processName) && pageTargets.Count == 1
                ? CreateClickTarget(port, pageTargets[0])
                : null;
        }

        private static async Task<BrowserClickTarget> RefreshTargetAsync(
            BrowserClickTarget recognizedTarget,
            CancellationToken cancellationToken)
        {
            var targets = await GetTargetsAsync(recognizedTarget.Port, cancellationToken);
            var pageTargets = targets
                .Where(target =>
                    string.Equals(target.Type, "page", StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(target.WebSocketDebuggerUrl)
                    && !target.Url.StartsWith("devtools://", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var matched = pageTargets.FirstOrDefault(target =>
                !string.IsNullOrWhiteSpace(recognizedTarget.Id)
                && string.Equals(target.Id, recognizedTarget.Id, StringComparison.OrdinalIgnoreCase));
            matched ??= pageTargets.FirstOrDefault(target =>
                !string.IsNullOrWhiteSpace(recognizedTarget.Url)
                && string.Equals(target.Url, recognizedTarget.Url, StringComparison.OrdinalIgnoreCase));
            matched ??= pageTargets.FirstOrDefault(target =>
                !string.IsNullOrWhiteSpace(recognizedTarget.Title)
                && IsTitleMatch(NormalizeTitle(target.Title), NormalizeTitle(recognizedTarget.Title)));

            return matched is null
                ? null
                : CreateClickTarget(recognizedTarget.Port, matched);
        }

        private static async Task<List<BrowserDebugTarget>> GetTargetsAsync(int port, CancellationToken cancellationToken)
        {
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(ProbeTimeoutMs);

            try
            {
                using var client = new HttpClient
                {
                    Timeout = TimeSpan.FromMilliseconds(ProbeTimeoutMs)
                };
                var json = await client.GetStringAsync($"http://127.0.0.1:{port}/json/list", timeoutCts.Token);
                return JsonSerializer.Deserialize<List<BrowserDebugTarget>>(json, JsonOptions) ?? new List<BrowserDebugTarget>();
            }
            catch
            {
                return new List<BrowserDebugTarget>();
            }
        }

        private static async Task<BrowserPoint> NormalizeClientPointAsync(
            CdpPageSession session,
            POINT clientPoint,
            CancellationToken cancellationToken)
        {
            var point = new BrowserPoint(clientPoint.X, clientPoint.Y);

            try
            {
                var metrics = await session.SendCommandAsync(
                    "Runtime.evaluate",
                    new
                    {
                        expression = "({ dpr: window.devicePixelRatio || 1, width: window.innerWidth || 0, height: window.innerHeight || 0 })",
                        returnByValue = true
                    },
                    cancellationToken);

                if (!TryGetRuntimeValue(metrics, out var value))
                {
                    return point;
                }

                var dpr = TryGetDouble(value, "dpr", 1);
                var width = TryGetDouble(value, "width", 0);
                var height = TryGetDouble(value, "height", 0);
                if (dpr > 1
                    && width > 0
                    && height > 0
                    && (point.X > width || point.Y > height)
                    && point.X / dpr <= width + 2
                    && point.Y / dpr <= height + 2)
                {
                    return new BrowserPoint(point.X / dpr, point.Y / dpr);
                }
            }
            catch
            {
                return point;
            }

            return point;
        }

        private static bool TryGetRuntimeValue(JsonElement element, out JsonElement value)
        {
            value = default;
            return element.TryGetProperty("result", out var result)
                && result.TryGetProperty("value", out value);
        }

        private static double TryGetDouble(JsonElement element, string name, double fallback)
        {
            if (!element.TryGetProperty(name, out var property))
            {
                return fallback;
            }

            return property.ValueKind == JsonValueKind.Number && property.TryGetDouble(out var value)
                ? value
                : fallback;
        }

        private static async Task DispatchClickAsync(
            CdpPageSession session,
            BrowserPoint point,
            MouseClickAction action,
            int holdDurationMs,
            CancellationToken cancellationToken)
        {
            switch (action)
            {
                case MouseClickAction.Right:
                    await DispatchButtonAsync(session, point, "right", 2, 1, 40, cancellationToken);
                    break;
                case MouseClickAction.DoubleLeft:
                    await DispatchButtonAsync(session, point, "left", 1, 1, 40, cancellationToken);
                    await Task.Delay(80, cancellationToken);
                    await DispatchButtonAsync(session, point, "left", 1, 2, 40, cancellationToken);
                    break;
                case MouseClickAction.HoldLeft:
                    await DispatchMouseEventAsync(session, "mouseMoved", "none", 0, point, 0, cancellationToken);
                    await DispatchMouseEventAsync(session, "mousePressed", "left", 1, point, 1, cancellationToken);
                    await Task.Delay(holdDurationMs, cancellationToken);
                    await DispatchMouseEventAsync(session, "mouseReleased", "left", 0, point, 1, cancellationToken);
                    break;
                default:
                    await DispatchButtonAsync(session, point, "left", 1, 1, 40, cancellationToken);
                    break;
            }
        }

        private static async Task DispatchButtonAsync(
            CdpPageSession session,
            BrowserPoint point,
            string button,
            int buttonMask,
            int clickCount,
            int delayMs,
            CancellationToken cancellationToken)
        {
            await DispatchMouseEventAsync(session, "mouseMoved", "none", 0, point, 0, cancellationToken);
            await Task.Delay(delayMs, cancellationToken);
            await DispatchMouseEventAsync(session, "mousePressed", button, buttonMask, point, clickCount, cancellationToken);
            await Task.Delay(delayMs, cancellationToken);
            await DispatchMouseEventAsync(session, "mouseReleased", button, 0, point, clickCount, cancellationToken);
        }

        private static Task DispatchMouseEventAsync(
            CdpPageSession session,
            string type,
            string button,
            int buttons,
            BrowserPoint point,
            int clickCount,
            CancellationToken cancellationToken)
        {
            return session.SendCommandAsync(
                "Input.dispatchMouseEvent",
                new
                {
                    type,
                    x = point.X,
                    y = point.Y,
                    button,
                    buttons,
                    clickCount
                },
                cancellationToken);
        }

        internal static bool IsChromiumProcess(string processName)
        {
            if (string.IsNullOrWhiteSpace(processName))
            {
                return false;
            }

            var normalized = processName.Trim().ToLowerInvariant();
            return normalized is "chrome"
                or "msedge"
                or "chromium"
                or "brave"
                or "brave-browser"
                or "vivaldi"
                or "opera"
                or "electron";
        }

        private static bool IsTitleMatch(string pageTitle, string targetTitle)
        {
            if (string.IsNullOrWhiteSpace(pageTitle) || string.IsNullOrWhiteSpace(targetTitle))
            {
                return false;
            }

            return string.Equals(pageTitle, targetTitle, StringComparison.OrdinalIgnoreCase)
                || pageTitle.Contains(targetTitle, StringComparison.OrdinalIgnoreCase)
                || targetTitle.Contains(pageTitle, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeTitle(string title)
        {
            var normalized = (title ?? string.Empty).Trim();
            var suffixes = new[]
            {
                " - Google Chrome",
                " - Microsoft Edge",
                " - Chromium",
                " - Brave",
                " - Vivaldi",
                " - Opera"
            };

            foreach (var suffix in suffixes)
            {
                if (normalized.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    normalized = normalized[..^suffix.Length].Trim();
                    break;
                }
            }

            return normalized;
        }

        private sealed class BrowserDebugTarget
        {
            public string Id { get; set; } = string.Empty;

            public string Type { get; set; } = string.Empty;

            public string Title { get; set; } = string.Empty;

            public string Url { get; set; } = string.Empty;

            public string WebSocketDebuggerUrl { get; set; } = string.Empty;
        }

        private static BrowserClickTarget CreateClickTarget(int port, BrowserDebugTarget target)
        {
            return new BrowserClickTarget
            {
                Port = port,
                Id = target.Id ?? string.Empty,
                Title = target.Title ?? string.Empty,
                Url = target.Url ?? string.Empty,
                WebSocketDebuggerUrl = target.WebSocketDebuggerUrl ?? string.Empty
            };
        }

        private readonly record struct BrowserPoint(double X, double Y);

        private sealed class CdpPageSession : IAsyncDisposable
        {
            private readonly ClientWebSocket _socket = new ClientWebSocket();
            private int _nextCommandId;

            public static async Task<CdpPageSession> ConnectAsync(
                string webSocketDebuggerUrl,
                CancellationToken cancellationToken)
            {
                var session = new CdpPageSession();
                await session._socket.ConnectAsync(new Uri(webSocketDebuggerUrl), cancellationToken);
                return session;
            }

            public async Task<JsonElement> SendCommandAsync(
                string method,
                object parameters,
                CancellationToken cancellationToken)
            {
                var commandId = Interlocked.Increment(ref _nextCommandId);
                var payload = new Dictionary<string, object>
                {
                    ["id"] = commandId,
                    ["method"] = method
                };
                if (parameters is not null)
                {
                    payload["params"] = parameters;
                }

                var bytes = JsonSerializer.SerializeToUtf8Bytes(payload);
                await _socket.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, cancellationToken);

                while (true)
                {
                    var message = await ReceiveMessageAsync(cancellationToken);
                    using var document = JsonDocument.Parse(message);
                    var root = document.RootElement;
                    if (!root.TryGetProperty("id", out var idElement) || idElement.GetInt32() != commandId)
                    {
                        continue;
                    }

                    if (root.TryGetProperty("error", out var errorElement))
                    {
                        throw new InvalidOperationException(errorElement.ToString());
                    }

                    return root.TryGetProperty("result", out var resultElement)
                        ? resultElement.Clone()
                        : default;
                }
            }

            private async Task<string> ReceiveMessageAsync(CancellationToken cancellationToken)
            {
                var buffer = new byte[64 * 1024];
                using var stream = new System.IO.MemoryStream();

                while (true)
                {
                    var result = await _socket.ReceiveAsync(new ArraySegment<byte>(buffer), cancellationToken);
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        throw new InvalidOperationException("Chrome DevTools connection was closed.");
                    }

                    stream.Write(buffer, 0, result.Count);
                    if (result.EndOfMessage)
                    {
                        return Encoding.UTF8.GetString(stream.ToArray());
                    }
                }
            }

            public async ValueTask DisposeAsync()
            {
                if (_socket.State == WebSocketState.Open)
                {
                    await _socket.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", CancellationToken.None);
                }

                _socket.Dispose();
            }
        }
    }
}
