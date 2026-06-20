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
        private static readonly string[] LoopbackHosts = { "127.0.0.1", "[::1]" };
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };

        private const int ProbeTimeoutMs = 1000;
        private const int WindowBoundsTolerance = 96;

        private int? _lastWorkingPort;

        internal async Task<BrowserClickTarget> TryResolveTargetAsync(
            string targetTitle,
            string processName,
            int preferredPort,
            CancellationToken cancellationToken)
        {
            return await TryResolveTargetAsync(
                targetTitle,
                processName,
                preferredPort,
                null,
                null,
                cancellationToken);
        }

        internal async Task<BrowserClickTarget> TryResolveTargetAsync(
            string targetTitle,
            string processName,
            int preferredPort,
            POINT? screenPoint,
            RECT? targetWindowRect,
            CancellationToken cancellationToken)
        {
            if (!IsChromiumProcess(processName) && string.IsNullOrWhiteSpace(targetTitle))
            {
                return null;
            }

            foreach (var port in GetProbePorts(preferredPort))
            {
                var target = await FindTargetAsync(
                    port,
                    targetTitle,
                    processName,
                    screenPoint,
                    targetWindowRect,
                    cancellationToken);
                if (target is null)
                {
                    continue;
                }

                _lastWorkingPort = port;
                return target;
            }

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
                yield break;
            }

            if (_lastWorkingPort is int rememberedPort)
            {
                yield return rememberedPort;
            }

            foreach (var port in DebugPorts)
            {
                if (port != _lastWorkingPort)
                {
                    yield return port;
                }
            }
        }

        private async Task<BrowserClickTarget> FindTargetAsync(
            int port,
            string targetTitle,
            string processName,
            POINT? screenPoint,
            RECT? targetWindowRect,
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
                var titleMatches = pageTargets
                    .Where(target => IsTitleMatch(NormalizeTitle(target.Title), normalizedTitle))
                    .ToList();
                if (titleMatches.Count == 1)
                {
                    return CreateClickTarget(port, titleMatches[0]);
                }

                if (titleMatches.Count > 1)
                {
                    var boundedTitleMatch = await TryFindTargetByWindowBoundsAsync(
                        titleMatches,
                        normalizedTitle,
                        screenPoint,
                        targetWindowRect,
                        cancellationToken);
                    return CreateClickTarget(port, boundedTitleMatch ?? titleMatches[0]);
                }
            }

            var boundedMatch = await TryFindTargetByWindowBoundsAsync(
                pageTargets,
                normalizedTitle,
                screenPoint,
                targetWindowRect,
                cancellationToken);
            if (boundedMatch is not null)
            {
                return CreateClickTarget(port, boundedMatch);
            }

            return IsChromiumProcess(processName) && pageTargets.Count == 1
                ? CreateClickTarget(port, pageTargets[0])
                : null;
        }

        private static async Task<BrowserDebugTarget> TryFindTargetByWindowBoundsAsync(
            List<BrowserDebugTarget> pageTargets,
            string normalizedTitle,
            POINT? screenPoint,
            RECT? targetWindowRect,
            CancellationToken cancellationToken)
        {
            if ((screenPoint is null && targetWindowRect is null) || pageTargets.Count == 0)
            {
                return null;
            }

            var candidates = new List<BrowserWindowTargetCandidate>();
            foreach (var pageTarget in pageTargets)
            {
                var bounds = await TryGetWindowBoundsAsync(pageTarget, cancellationToken);
                if (bounds is null)
                {
                    continue;
                }

                var score = ScoreWindowBounds(bounds.Value, screenPoint, targetWindowRect);
                if (score <= 0)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(normalizedTitle)
                    && IsTitleMatch(NormalizeTitle(pageTarget.Title), normalizedTitle))
                {
                    score += 100;
                }

                candidates.Add(new BrowserWindowTargetCandidate(pageTarget, score));
            }

            if (candidates.Count == 0)
            {
                return null;
            }

            var ordered = candidates
                .OrderByDescending(candidate => candidate.Score)
                .ToList();
            if (ordered.Count > 1 && ordered[0].Score == ordered[1].Score)
            {
                return null;
            }

            return ordered[0].Target;
        }

        private static async Task<BrowserWindowBounds?> TryGetWindowBoundsAsync(
            BrowserDebugTarget target,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(target.WebSocketDebuggerUrl))
            {
                return null;
            }

            try
            {
                using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                timeoutCts.CancelAfter(ProbeTimeoutMs);
                await using var session = await CdpPageSession.ConnectAsync(target.WebSocketDebuggerUrl, timeoutCts.Token);
                var result = await session.SendCommandAsync(
                    "Browser.getWindowForTarget",
                    null,
                    timeoutCts.Token);

                if (!result.TryGetProperty("bounds", out var bounds))
                {
                    return null;
                }

                var left = TryGetInt(bounds, "left", int.MinValue);
                var top = TryGetInt(bounds, "top", int.MinValue);
                var width = TryGetInt(bounds, "width", 0);
                var height = TryGetInt(bounds, "height", 0);
                if (left == int.MinValue || top == int.MinValue || width <= 0 || height <= 0)
                {
                    return null;
                }

                return new BrowserWindowBounds(left, top, left + width, top + height);
            }
            catch
            {
                return null;
            }
        }

        private static int ScoreWindowBounds(
            BrowserWindowBounds bounds,
            POINT? screenPoint,
            RECT? targetWindowRect)
        {
            var score = 0;
            if (screenPoint.HasValue && IsPointInside(bounds, screenPoint.Value, WindowBoundsTolerance))
            {
                score += 70;
            }

            if (targetWindowRect.HasValue)
            {
                var rect = targetWindowRect.Value;
                if (IsSimilarBounds(bounds, rect, WindowBoundsTolerance))
                {
                    score += 90;
                }
                else if (HasMeaningfulOverlap(bounds, rect))
                {
                    score += 40;
                }
            }

            return score;
        }

        private static bool IsPointInside(BrowserWindowBounds bounds, POINT point, int tolerance)
        {
            return point.X >= bounds.Left - tolerance
                && point.X <= bounds.Right + tolerance
                && point.Y >= bounds.Top - tolerance
                && point.Y <= bounds.Bottom + tolerance;
        }

        private static bool IsSimilarBounds(BrowserWindowBounds bounds, RECT rect, int tolerance)
        {
            return Math.Abs(bounds.Left - rect.Left) <= tolerance
                && Math.Abs(bounds.Top - rect.Top) <= tolerance
                && Math.Abs(bounds.Right - rect.Right) <= tolerance
                && Math.Abs(bounds.Bottom - rect.Bottom) <= tolerance;
        }

        private static bool HasMeaningfulOverlap(BrowserWindowBounds bounds, RECT rect)
        {
            var overlapLeft = Math.Max(bounds.Left, rect.Left);
            var overlapTop = Math.Max(bounds.Top, rect.Top);
            var overlapRight = Math.Min(bounds.Right, rect.Right);
            var overlapBottom = Math.Min(bounds.Bottom, rect.Bottom);
            var overlapWidth = Math.Max(0, overlapRight - overlapLeft);
            var overlapHeight = Math.Max(0, overlapBottom - overlapTop);
            if (overlapWidth == 0 || overlapHeight == 0)
            {
                return false;
            }

            var overlapArea = (long)overlapWidth * overlapHeight;
            var boundsArea = (long)Math.Max(0, bounds.Right - bounds.Left) * Math.Max(0, bounds.Bottom - bounds.Top);
            var rectArea = (long)Math.Max(0, rect.Right - rect.Left) * Math.Max(0, rect.Bottom - rect.Top);
            var smallerArea = Math.Min(boundsArea, rectArea);
            return smallerArea > 0 && overlapArea * 2 >= smallerArea;
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
            var allTargets = new List<BrowserDebugTarget>();
            foreach (var host in LoopbackHosts)
            {
                using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                timeoutCts.CancelAfter(ProbeTimeoutMs);

                try
                {
                    using var client = new HttpClient
                    {
                        Timeout = TimeSpan.FromMilliseconds(ProbeTimeoutMs)
                    };
                    var json = await client.GetStringAsync($"http://{host}:{port}/json/list", timeoutCts.Token);
                    var targets = JsonSerializer.Deserialize<List<BrowserDebugTarget>>(json, JsonOptions);
                    if (targets is null)
                    {
                        continue;
                    }

                    foreach (var target in targets)
                    {
                        target.WebSocketDebuggerUrl = NormalizeLoopbackWebSocketUrl(target.WebSocketDebuggerUrl, host, port);
                        allTargets.Add(target);
                    }
                }
                catch
                {
                    // Either IPv4 or IPv6 loopback may be unavailable; keep probing the other endpoint.
                }
            }

            return allTargets;
        }

        private static string NormalizeLoopbackWebSocketUrl(string webSocketDebuggerUrl, string host, int port)
        {
            if (string.IsNullOrWhiteSpace(webSocketDebuggerUrl))
            {
                return string.Empty;
            }

            var normalizedHost = string.Equals(host, "[::1]", StringComparison.Ordinal)
                ? "[::1]"
                : "127.0.0.1";
            var pathStart = webSocketDebuggerUrl.IndexOf("/devtools/", StringComparison.OrdinalIgnoreCase);
            return pathStart >= 0
                ? $"ws://{normalizedHost}:{port}{webSocketDebuggerUrl[pathStart..]}"
                : webSocketDebuggerUrl;
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

        private static int TryGetInt(JsonElement element, string name, int fallback)
        {
            if (!element.TryGetProperty(name, out var property))
            {
                return fallback;
            }

            return property.ValueKind == JsonValueKind.Number && property.TryGetInt32(out var value)
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

        private sealed record BrowserWindowTargetCandidate(BrowserDebugTarget Target, int Score);

        private readonly record struct BrowserWindowBounds(int Left, int Top, int Right, int Bottom);

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
