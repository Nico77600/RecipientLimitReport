// =============================================================================
//  RecipientLimitReport.Engine.cs
// -----------------------------------------------------------------------------
//  Author  : Nicolas Fabert
//  Version : 1.0.1
//  Purpose : Performance-critical part of the Recipient Limit Report tool.
//            This file is compiled automatically by RecipientLimitReport.psm1 the
//            first time it is used (and again whenever it changes). The compiled
//            DLL is cached in the .\bin folder.
//
//  Contents
//    0. ConsoleFont    : font of the classic console (frame characters)
//    1. Common         : time ranges, coverage, time and text helpers
//    2. Graph          : access token, rolling quota, HTTP client, message trace
//                        collector, recipient count collector (getDetailsByRecipient)
//    3. RlrStore       : SQLite database (schema, ingestion, counts, compaction, queries)
//    4. ReportWriter   : CSV and HTML files, split into parts
//
//  Conventions
//    - Every time stored in the database is a Unix epoch in MILLISECONDS (UTC).
//    - Time ranges are always [start, end) : start included, end excluded.
//    - The message trace returns ONE row per message and recipient. The database
//      keeps one row per message trace ID and one row per recipient of it; reports
//      show one row per Message ID.
//    - Recipient count = number of recipients BEFORE the expansion of distribution
//      lists, the value checked by Exchange Online (RecipientLimits) and by the DLP
//      condition RecipientCountOver (a distribution list counts as one recipient):
//        no distribution list : the recipients of the trace (each one was addressed)
//        distribution list(s) : RcptCount of the Submit event of the message
//                               (getDetailsByRecipient), read once per message
//    - Recipient list = the recipients as the sender addressed them. For a message
//      over the limit sent to distribution lists, the route of each recipient tells
//      whether the sender addressed it (the route starts with the Submit event) or the
//      expansion of a list added it (the route starts after the expansion).
// =============================================================================
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;

namespace RecipientLimitReport
{
    // -------------------------------------------------------------------------
    // 0. Console
    // -------------------------------------------------------------------------

    /// <summary>
    /// Font of the classic Windows console (conhost). The console has no font fallback, so the
    /// module chooses the frame characters from it. Returns null when the output is not a classic
    /// console window (Windows Terminal, VS Code, redirected output) or on any error.
    /// </summary>
    public static class ConsoleFont
    {
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        struct ConsoleFontInfoEx
        {
            public uint Size; public uint Font; public short Width; public short Height; public int Family; public int Weight;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 32)] public string FaceName;
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GetStdHandle(int handle);

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        static extern bool GetCurrentConsoleFontEx(IntPtr output, bool maximumWindow, ref ConsoleFontInfoEx info);

        public static string FaceName()
        {
            try
            {
                if (Console.IsOutputRedirected || !OperatingSystem.IsWindows()) return null;
                var info = new ConsoleFontInfoEx { Size = (uint)System.Runtime.InteropServices.Marshal.SizeOf<ConsoleFontInfoEx>() };
                return GetCurrentConsoleFontEx(GetStdHandle(-11), false, ref info) ? info.FaceName : null;
            }
            catch { return null; }
        }
    }

    // -------------------------------------------------------------------------
    // 1. Common: time ranges, time, text
    // -------------------------------------------------------------------------

    public struct TimeRange
    {
        public long Start;   // Unix ms, included
        public long End;     // Unix ms, excluded
        public TimeRange(long start, long end) { Start = start; End = end; }
        public long Length { get { return End - Start; } }
        public override string ToString() { return Time.Iso(Start) + " -> " + Time.Iso(End); }
    }

    public static class Coverage
    {
        /// <summary>Merges overlapping or adjacent ranges.</summary>
        public static List<TimeRange> Merge(IEnumerable<TimeRange> ranges)
        {
            var result = new List<TimeRange>();
            foreach (TimeRange r in ranges.Where(x => x.End > x.Start).OrderBy(x => x.Start))
            {
                if (result.Count > 0 && r.Start <= result[result.Count - 1].End)
                {
                    TimeRange last = result[result.Count - 1];
                    if (r.End > last.End) result[result.Count - 1] = new TimeRange(last.Start, r.End);
                }
                else result.Add(r);
            }
            return result;
        }

        /// <summary>Returns the parts of [start, end) that are NOT covered by the given ranges.</summary>
        public static List<TimeRange> Gaps(IEnumerable<TimeRange> covered, long start, long end)
        {
            var gaps = new List<TimeRange>();
            if (end <= start) return gaps;
            long cursor = start;
            foreach (TimeRange r in Merge(covered))
            {
                if (r.End <= cursor) continue;
                if (r.Start >= end) break;
                if (r.Start > cursor) gaps.Add(new TimeRange(cursor, Math.Min(r.Start, end)));
                cursor = Math.Max(cursor, r.End);
                if (cursor >= end) break;
            }
            if (cursor < end) gaps.Add(new TimeRange(cursor, end));
            return gaps;
        }

        /// <summary>Parts covered by BOTH lists of ranges.</summary>
        public static List<TimeRange> Intersect(IEnumerable<TimeRange> a, IEnumerable<TimeRange> b)
        {
            List<TimeRange> x = Merge(a), y = Merge(b);
            var result = new List<TimeRange>();
            int i = 0, j = 0;
            while (i < x.Count && j < y.Count)
            {
                long start = Math.Max(x[i].Start, y[j].Start), end = Math.Min(x[i].End, y[j].End);
                if (end > start) result.Add(new TimeRange(start, end));
                if (x[i].End < y[j].End) i++; else j++;
            }
            return result;
        }

        /// <summary>Total length of the ranges, in milliseconds (overlaps counted once).</summary>
        public static long Length(IEnumerable<TimeRange> ranges)
        {
            long sum = 0;
            foreach (TimeRange r in Merge(ranges)) sum += r.Length;
            return sum;
        }

        /// <summary>
        /// Splits ranges at local midnights, then into pieces of at most maxHours, the newest first.
        /// The message trace API returns the newest messages first: collecting the newest slices
        /// first gives the most useful data early.
        /// </summary>
        public static List<TimeRange> SplitIntoSlices(IEnumerable<TimeRange> ranges, TimeZoneInfo zone, int maxHours)
        {
            var slices = new List<TimeRange>();
            long maxMs = Math.Max(1, maxHours) * 3600000L;
            foreach (TimeRange range in Merge(ranges))
            {
                long cursor = range.Start;
                while (cursor < range.End)
                {
                    long nextMidnight = NextLocalMidnight(cursor, zone);
                    long stop = Math.Min(range.End, Math.Min(nextMidnight, cursor + maxMs));
                    slices.Add(new TimeRange(cursor, stop));
                    cursor = stop;
                }
            }
            return slices.OrderByDescending(s => s.End).ToList();
        }

        public static long NextLocalMidnight(long unixMs, TimeZoneInfo zone)
        {
            DateTime local = TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).DateTime;
            return LocalToUnixMs(local.Date.AddDays(1), zone);
        }

        public static long LocalToUnixMs(DateTime local, TimeZoneInfo zone)
        {
            DateTime unspecified = DateTime.SpecifyKind(local, DateTimeKind.Unspecified);
            // Local midnight can be invalid (spring DST gap) in a few zones: move forward until valid.
            while (zone.IsInvalidTime(unspecified)) unspecified = unspecified.AddMinutes(30);
            TimeSpan offset = zone.IsAmbiguousTime(unspecified)
                ? zone.GetAmbiguousTimeOffsets(unspecified).Max()
                : zone.GetUtcOffset(unspecified);
            return new DateTimeOffset(unspecified, offset).ToUnixTimeMilliseconds();
        }
    }

    public static class Time
    {
        public static long NowMs() { return DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(); }

        public static string Iso(long unixMs)
        {
            return DateTimeOffset.FromUnixTimeMilliseconds(unixMs).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
        }

        /// <summary>Graph filter value in whole seconds: floor for a start, ceiling for an end, so a query never misses a row.</summary>
        public static string GraphTime(long unixMs, bool ceiling)
        {
            long s = unixMs / 1000;
            if (ceiling && unixMs % 1000 != 0) s++;
            return DateTimeOffset.FromUnixTimeSeconds(s).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        }

        public static bool TryParse(string text, out long unixMs)
        {
            unixMs = 0;
            if (string.IsNullOrEmpty(text)) return false;
            DateTimeOffset value;
            if (!DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out value)) return false;
            unixMs = value.ToUnixTimeMilliseconds();
            return true;
        }

        public static string FormatLocal(long unixMs, TimeZoneInfo zone)
        {
            return TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }
    }

    public static class Text
    {
        /// <summary>
        /// CSV cell for Excel: quoted when needed. Text starting with = + - @ (or a control
        /// character) is prefixed with an apostrophe so Excel never evaluates it as a formula.
        /// </summary>
        public static string SafeCsv(string value, string delimiter)
        {
            if (string.IsNullOrEmpty(value)) return "";
            string v = value;
            char first = v[0];
            if (first == '=' || first == '+' || first == '-' || first == '@' || first == '\t' || first == '\r') v = "'" + v;
            bool quote = v.Contains(delimiter) || v.IndexOf('"') >= 0 || v.IndexOf('\n') >= 0 || v.IndexOf('\r') >= 0;
            return quote ? "\"" + v.Replace("\"", "\"\"") + "\"" : v;
        }

        /// <summary>OData string literal: single quotes doubled.</summary>
        public static string ODataString(string value) { return "'" + (value ?? "").Replace("'", "''") + "'"; }

        public static string Shorten(string value, int max)
        {
            if (value == null) return "";
            return value.Length <= max ? value : value.Substring(0, Math.Max(0, max - 3)) + "...";
        }
    }

    // -------------------------------------------------------------------------
    // 2. Microsoft Graph: message trace and recipient count
    // -------------------------------------------------------------------------
    //  Throughput is bounded by Microsoft, not by the tool: 100 requests per 5 minutes
    //  and per tenant (rolling window), 5,000 rows per request. getDetailsByRecipient
    //  has its own quota of the same size. What matters is to send as few requests as
    //  possible and never to waste one on a 429.

    /// <summary>One row of the message trace (exchangeMessageTrace): one message for one recipient.</summary>
    public sealed class TraceRow
    {
        public string Id, MessageId, Sender, Recipient, Subject, Status;
        public long ReceivedMs;
    }

    public sealed class TokenLease { public string Token; public int Version; }

    /// <summary>
    /// Access token shared by the workers. The module (PowerShell) owns the authentication: it reads
    /// <see cref="NeedsRefresh"/> while the collection runs and calls <see cref="Set"/>.
    /// </summary>
    public sealed class TokenSlot
    {
        readonly object _lock = new object();
        string _token;
        long _expiresMs;
        int _version;
        bool _refreshRequested;

        public long ExpiresMs { get { lock (_lock) return _expiresMs; } }

        public void Set(string token, long expiresMs)
        {
            lock (_lock) { _token = token; _expiresMs = expiresMs; _version++; _refreshRequested = false; }
        }

        /// <summary>True when the token is missing, expires within marginMs, or a worker got a 401.</summary>
        public bool NeedsRefresh(long marginMs)
        {
            lock (_lock) return _token == null || _refreshRequested || Time.NowMs() > _expiresMs - marginMs;
        }

        /// <summary>A worker got a 401 with this token version: ask the module for a new one (once per version).</summary>
        public void RequestRefresh(int version)
        {
            lock (_lock) { if (version == _version) _refreshRequested = true; }
        }

        public async Task<TokenLease> GetAsync(CancellationToken ct, int timeoutSeconds = 180)
        {
            long deadline = Time.NowMs() + timeoutSeconds * 1000L;
            while (true)
            {
                lock (_lock)
                {
                    if (_token != null && !_refreshRequested && Time.NowMs() < _expiresMs - 30000) return new TokenLease { Token = _token, Version = _version };
                }
                if (Time.NowMs() > deadline) throw new FatalCollectionException("No valid access token: the token could not be renewed in time.");
                await Task.Delay(200, ct).ConfigureAwait(false);
            }
        }
    }

    /// <summary>
    /// Rolling-window budget: at most MaxRequests requests in any PeriodMs, shared by every worker.
    /// The quota is per tenant: a 429 pauses every worker (Retry-After), and the request times of the
    /// last window are saved in the database so that the next run starts with what is already used.
    /// </summary>
    public sealed class RateLimiter
    {
        readonly object _lock = new object();
        readonly Queue<long> _stamps = new Queue<long>();
        long _pauseUntil;
        long _waitedMs;
        public readonly int MaxRequests;
        public readonly long PeriodMs;

        public RateLimiter(int maxRequests, long periodMs, IEnumerable<long> previousStamps = null)
        {
            MaxRequests = Math.Max(1, maxRequests);
            PeriodMs = Math.Max(1000, periodMs);
            if (previousStamps != null)
            {
                long now = Time.NowMs();
                foreach (long s in previousStamps.Where(x => x > now - PeriodMs && x <= now).OrderBy(x => x)) _stamps.Enqueue(s);
            }
        }

        public long WaitedMs { get { return Interlocked.Read(ref _waitedMs); } }
        public int InWindow { get { lock (_lock) { Purge(Time.NowMs()); return _stamps.Count; } } }
        public long PausedUntil { get { lock (_lock) return _pauseUntil; } }

        void Purge(long now) { while (_stamps.Count > 0 && _stamps.Peek() <= now - PeriodMs) _stamps.Dequeue(); }

        /// <summary>Time to wait before the next request could start (0 = now).</summary>
        public long DelayMs()
        {
            lock (_lock)
            {
                long now = Time.NowMs();
                Purge(now);
                if (now < _pauseUntil) return _pauseUntil - now;
                if (_stamps.Count < MaxRequests) return 0;
                return Math.Max(1, _stamps.Peek() + PeriodMs - now);
            }
        }

        public async Task WaitAsync(CancellationToken ct)
        {
            while (true)
            {
                long delay;
                lock (_lock)
                {
                    long now = Time.NowMs();
                    Purge(now);
                    if (now >= _pauseUntil && _stamps.Count < MaxRequests) { _stamps.Enqueue(now); return; }
                    delay = now < _pauseUntil ? _pauseUntil - now : _stamps.Peek() + PeriodMs - now + 20;
                }
                int step = (int)Math.Max(10, Math.Min(delay, 1000));
                Interlocked.Add(ref _waitedMs, step);
                await Task.Delay(step, ct).ConfigureAwait(false);
            }
        }

        /// <summary>Every worker waits until now + ms (429 from Graph).</summary>
        public void Pause(long ms)
        {
            lock (_lock) { _pauseUntil = Math.Max(_pauseUntil, Time.NowMs() + ms); }
        }

        public long[] Stamps()
        {
            lock (_lock) { Purge(Time.NowMs()); return _stamps.ToArray(); }
        }
    }

    public sealed class CollectorOptions
    {
        public string GraphRoot = "https://graph.microsoft.com/v1.0";
        public int PageSize = 5000;
        public int MaxConcurrency = 3;
        public int MaxRetries = 5;              // 5xx, timeouts, network errors
        public int MaxThrottleRetries = 20;     // 429
        public int TimeoutSeconds = 180;
        public int DefaultRetryAfterSeconds = 30;
        public string UserAgent = "RecipientLimitReport/1.0.1";
        public HttpMessageHandler Handler;      // tests: a fake Graph
    }

    public sealed class CollectorEvent
    {
        public long Ms = Time.NowMs();
        public string Level;   // INFO | WARN | ERROR
        public string Text;
    }

    /// <summary>An error that stops the whole collection (permissions, service principal, token).</summary>
    public sealed class FatalCollectionException : Exception
    {
        public int Status;
        public FatalCollectionException(string message, int status = 0) : base(message) { Status = status; }
    }

    /// <summary>An error that fails one work item only (refused request, too old, retries exhausted).</summary>
    public sealed class ItemFailureException : Exception
    {
        public int Status;
        public bool Unrecoverable;
        public ItemFailureException(string message, int status, bool unrecoverable = false) : base(message) { Status = status; Unrecoverable = unrecoverable; }
    }

    /// <summary>One GET to Graph with the quota, the token and the retries.</summary>
    public sealed class GraphClient : IDisposable
    {
        readonly HttpClient _http;
        readonly CollectorOptions _o;
        readonly TokenSlot _token;
        readonly RateLimiter _limiter;
        readonly ConcurrentQueue<CollectorEvent> _events;
        public long Requests, Throttled, Retries, Bytes;

        public GraphClient(CollectorOptions o, TokenSlot token, RateLimiter limiter, ConcurrentQueue<CollectorEvent> events)
        {
            _o = o; _token = token; _limiter = limiter; _events = events;
            HttpMessageHandler handler = o.Handler ?? new SocketsHttpHandler
            {
                AutomaticDecompression = DecompressionMethods.All,
                PooledConnectionLifetime = TimeSpan.FromMinutes(5),
                MaxConnectionsPerServer = Math.Max(2, o.MaxConcurrency + 1)
            };
            _http = new HttpClient(handler, o.Handler == null) { Timeout = Timeout.InfiniteTimeSpan };
            _http.DefaultRequestHeaders.UserAgent.ParseAdd(o.UserAgent);
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        void Event(string level, string text) { _events.Enqueue(new CollectorEvent { Level = level, Text = text }); }

        public static string GraphError(string body)
        {
            if (string.IsNullOrEmpty(body)) return "";
            try
            {
                using (JsonDocument doc = JsonDocument.Parse(body))
                {
                    JsonElement e;
                    if (doc.RootElement.TryGetProperty("error", out e))
                    {
                        string code = e.TryGetProperty("code", out JsonElement c) ? c.GetString() : "";
                        string message = e.TryGetProperty("message", out JsonElement m) ? m.GetString() : "";
                        return (string.IsNullOrEmpty(code) ? "" : code + ": ") + message;
                    }
                }
            }
            catch { }
            return Text.Shorten(body.Replace("\r", " ").Replace("\n", " "), 300);
        }

        public async Task<string> GetAsync(string url, string what, CancellationToken ct)
        {
            int attempt = 0, throttles = 0, authRetries = 0;
            while (true)
            {
                ct.ThrowIfCancellationRequested();
                await _limiter.WaitAsync(ct).ConfigureAwait(false);
                TokenLease lease = await _token.GetAsync(ct).ConfigureAwait(false);
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", lease.Token);
                request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
                HttpResponseMessage response = null;
                string body = null;
                string failure = null;
                using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
                {
                    timeout.CancelAfter(TimeSpan.FromSeconds(_o.TimeoutSeconds));
                    try
                    {
                        response = await _http.SendAsync(request, HttpCompletionOption.ResponseContentRead, timeout.Token).ConfigureAwait(false);
                        body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    }
                    catch (OperationCanceledException) when (!ct.IsCancellationRequested) { failure = "no answer after " + _o.TimeoutSeconds + " s"; }
                    catch (HttpRequestException ex) { failure = ex.Message; }
                    finally { request.Dispose(); }
                }
                Interlocked.Increment(ref Requests);
                if (failure != null)
                {
                    if (++attempt > _o.MaxRetries) throw new ItemFailureException(what + ": " + failure + " (after " + _o.MaxRetries + " retries)", 0);
                    Interlocked.Increment(ref Retries);
                    int wait = Math.Min(60, 1 << attempt);
                    Event("WARN", what + ": " + failure + " - retry " + attempt + "/" + _o.MaxRetries + " in " + wait + " s");
                    await Task.Delay(wait * 1000, ct).ConfigureAwait(false);
                    continue;
                }
                int status = (int)response.StatusCode;
                Interlocked.Add(ref Bytes, body == null ? 0 : body.Length);
                if (status == 200) { response.Dispose(); return body; }

                string error = GraphError(body);
                if (status == 429)
                {
                    Interlocked.Increment(ref Throttled);
                    long waitMs = _o.DefaultRetryAfterSeconds * 1000L * Math.Min(4, throttles + 1);
                    if (response.Headers.RetryAfter != null)
                    {
                        if (response.Headers.RetryAfter.Delta.HasValue) waitMs = (long)response.Headers.RetryAfter.Delta.Value.TotalMilliseconds + 1000;
                        else if (response.Headers.RetryAfter.Date.HasValue) waitMs = Math.Max(1000, (long)(response.Headers.RetryAfter.Date.Value - DateTimeOffset.UtcNow).TotalMilliseconds + 1000);
                    }
                    response.Dispose();
                    if (++throttles > _o.MaxThrottleRetries) throw new ItemFailureException(what + ": still throttled after " + _o.MaxThrottleRetries + " waits (429). Another tool may be using the message trace quota of the tenant.", 429);
                    _limiter.Pause(waitMs);
                    Event("WARN", "Throttled by Microsoft Graph (429): every request waits " + Math.Round(waitMs / 1000.0) + " s. " + error);
                    continue;
                }
                response.Dispose();
                if (status == 401)
                {
                    if (error.IndexOf("8bd644d1-64a1-4d4b-ae52-2e0cbf64e373", StringComparison.OrdinalIgnoreCase) >= 0 || error.IndexOf("service principal", StringComparison.OrdinalIgnoreCase) >= 0)
                        throw new FatalCollectionException("The service principal of the Microsoft message trace application (8bd644d1-64a1-4d4b-ae52-2e0cbf64e373) is missing or not provisioned yet in the tenant. Create it once (guide, chapter 4); provisioning can take several hours. Graph: " + error, 401);
                    if (++authRetries > 2) throw new FatalCollectionException("Access denied (401) with a renewed token: " + error, 401);
                    _token.RequestRefresh(lease.Version);
                    Event("WARN", "401 from Microsoft Graph: renewing the access token. " + error);
                    continue;
                }
                if (status == 403)
                    throw new FatalCollectionException("Permission denied (403): the application or the account needs ExchangeMessageTrace.Read.All (with admin consent). Graph: " + error, 403);
                if (status == 400)
                {
                    bool tooOld = error.IndexOf("90 days", StringComparison.OrdinalIgnoreCase) >= 0;
                    throw new ItemFailureException(what + ": request refused (400) - " + error, 400, tooOld);
                }
                if (status == 408 || status >= 500)
                {
                    if (++attempt > _o.MaxRetries) throw new ItemFailureException(what + ": HTTP " + status + " after " + _o.MaxRetries + " retries - " + error, status);
                    Interlocked.Increment(ref Retries);
                    int wait = Math.Min(60, 1 << attempt);
                    Event("WARN", what + ": HTTP " + status + " - retry " + attempt + "/" + _o.MaxRetries + " in " + wait + " s. " + error);
                    await Task.Delay(wait * 1000, ct).ConfigureAwait(false);
                    continue;
                }
                throw new ItemFailureException(what + ": HTTP " + status + " - " + error, status);
            }
        }

        public void Dispose() { _http.Dispose(); }
    }

    /// <summary>What the route of one recipient tells (getDetailsByRecipient).</summary>
    public sealed class CountEvidence
    {
        public long? Count;       // recipients before expansion
        public string Source;     // Submit | Receive
        public int Events;
        public bool HasSubmit, HasReceive;   // the recipient was in the message when it was submitted / received
    }

    public static class TraceParser
    {
        static readonly Regex RcptCount = new Regex(@"Name=""RcptCount""\s+Integer=""(\d{1,9})""", RegexOptions.CultureInvariant | RegexOptions.Compiled);

        static string Str(JsonElement e, string name)
        {
            JsonElement v;
            if (!e.TryGetProperty(name, out v) || v.ValueKind == JsonValueKind.Null) return "";
            return v.ValueKind == JsonValueKind.String ? v.GetString() : v.ToString();
        }

        /// <summary>Rows and @odata.nextLink of one message trace page.</summary>
        public static List<TraceRow> Parse(string body, out string nextLink)
        {
            var rows = new List<TraceRow>();
            nextLink = null;
            using (JsonDocument doc = JsonDocument.Parse(body))
            {
                JsonElement root = doc.RootElement, value, next;
                if (root.TryGetProperty("@odata.nextLink", out next) && next.ValueKind == JsonValueKind.String) nextLink = next.GetString();
                if (!root.TryGetProperty("value", out value) || value.ValueKind != JsonValueKind.Array) throw new FormatException("The Graph answer has no 'value' array.");
                foreach (JsonElement e in value.EnumerateArray())
                {
                    var r = new TraceRow
                    {
                        Id = Str(e, "id"), MessageId = Str(e, "messageId"), Sender = Str(e, "senderAddress"), Recipient = Str(e, "recipientAddress"),
                        Subject = Str(e, "subject"), Status = Str(e, "status")
                    };
                    long ms;
                    if (!Time.TryParse(Str(e, "receivedDateTime"), out ms)) throw new FormatException("receivedDateTime not readable for trace " + r.Id + ".");
                    r.ReceivedMs = ms;
                    rows.Add(r);
                }
            }
            return rows;
        }

        /// <summary>
        /// Recipient count before expansion, from the events of getDetailsByRecipient.
        /// The Submit event (message submitted by a mailbox) carries RcptCount = every recipient
        /// of the message as the sender addressed it, a distribution list counting as one. A
        /// message received by SMTP has no Submit event: its Receive event carries the count of the
        /// SMTP envelope. The other events (Expand DL, DLP rule, Deliver ...) carry other counts
        /// (members of one list, recipients after expansion): they are never used.
        /// The route of a recipient addressed by the sender starts with these events; the route of a
        /// member added by the expansion of a list starts after them (HasSubmit / HasReceive false).
        /// </summary>
        public static CountEvidence ParseCount(string body, out string nextLink)
        {
            var result = new CountEvidence();
            nextLink = null;
            long? receive = null;
            using (JsonDocument doc = JsonDocument.Parse(body))
            {
                JsonElement root = doc.RootElement, value, next;
                if (root.TryGetProperty("@odata.nextLink", out next) && next.ValueKind == JsonValueKind.String) nextLink = next.GetString();
                if (!root.TryGetProperty("value", out value) || value.ValueKind != JsonValueKind.Array) throw new FormatException("The Graph answer has no 'value' array.");
                foreach (JsonElement e in value.EnumerateArray())
                {
                    result.Events++;
                    string name = Str(e, "event");
                    bool submit = string.Equals(name, "Submit", StringComparison.OrdinalIgnoreCase), received = string.Equals(name, "Receive", StringComparison.OrdinalIgnoreCase);
                    if (submit) result.HasSubmit = true;
                    if (received) result.HasReceive = true;
                    Match m = RcptCount.Match(Str(e, "data"));
                    if (!m.Success) continue;
                    long n = long.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                    if (submit && !result.Count.HasValue) { result.Count = n; result.Source = "Submit"; }
                    else if (received && !receive.HasValue) receive = n;
                }
            }
            if (!result.Count.HasValue && receive.HasValue) { result.Count = receive; result.Source = "Receive"; }
            return result;
        }
    }

    /// <summary>One request series: one query (all senders or one sender domain) on one time slice.</summary>
    public sealed class WorkItem
    {
        public int Index;
        public string Signature = "";      // '' = every sender, or 'sender=*@contoso.com'
        public string Condition = "";      // OData condition added to the time filter ('' = none)
        public string Label = "";          // short text for the console
        public long SignatureId;
        public long StartMs, EndMs;
        public string State = "Pending";   // Pending | Running | Done | Failed | Cancelled
        public long Pages, Rows, NewMessages;
        public long OldestMs;              // oldest row received so far (the API returns the newest first)
        public long IntervalId;            // collection_interval row (created with the first page)
        public long CoveredFromMs;         // [CoveredFromMs, EndMs) is stored
        public long StartedMs, EndedMs;
        public string Error;
        public bool Unrecoverable;
        public bool Reported;              // already shown in the console table

        /// <summary>Share of the slice already collected.</summary>
        public double Fraction
        {
            get
            {
                if (State == "Done") return 1.0;
                if (OldestMs <= 0 || EndMs <= StartMs) return 0.0;
                return Math.Max(0.0, Math.Min(1.0, (double)(EndMs - OldestMs) / (EndMs - StartMs)));
            }
        }

        public string ODataFilter()
        {
            string f = "receivedDateTime ge " + Time.GraphTime(StartMs, false) + " and receivedDateTime le " + Time.GraphTime(EndMs, true);
            return string.IsNullOrEmpty(Condition) ? f : f + " and " + Condition;
        }
    }

    /// <summary>One page handed by a worker to the writer.</summary>
    sealed class PageResult
    {
        public WorkItem Item;
        public List<TraceRow> Rows;
        public bool Last;
        public bool Ordered;
        public long OldestMs;
        public long FetchedMs;
        public string Error;
        public bool Unrecoverable;
    }

    /// <summary>
    /// Runs the work items: MaxConcurrency workers send the requests, one writer thread stores the pages
    /// (the SQLite connection is never shared between threads). Cancelling keeps what was stored.
    /// </summary>
    public sealed class TraceCollector
    {
        readonly RlrStore _store;
        readonly CollectorOptions _o;
        readonly TokenSlot _token;
        readonly RateLimiter _limiter;
        readonly long _runId;
        List<WorkItem> _items = new List<WorkItem>();
        GraphClient _graph;
        CancellationTokenSource _stop;

        public readonly ConcurrentQueue<CollectorEvent> Events = new ConcurrentQueue<CollectorEvent>();
        public long Pages, Rows, NewMessages, NewRows, ItemsDone, ItemsFailed, ItemsRunning;
        public string Fatal;
        public int FatalStatus;
        public long StartedMs, EndedMs;

        public TraceCollector(RlrStore store, CollectorOptions options, TokenSlot token, RateLimiter limiter, long runId)
        {
            _store = store; _o = options; _token = token; _limiter = limiter; _runId = runId;
        }

        public long Requests { get { return _graph == null ? 0 : Interlocked.Read(ref _graph.Requests); } }
        public long Throttled { get { return _graph == null ? 0 : Interlocked.Read(ref _graph.Throttled); } }
        public long Retries { get { return _graph == null ? 0 : Interlocked.Read(ref _graph.Retries); } }

        /// <summary>Share of the planned time span already collected.</summary>
        public double Progress
        {
            get
            {
                double total = 0, done = 0;
                foreach (WorkItem i in _items)
                {
                    double span = Math.Max(1, i.EndMs - i.StartMs);
                    total += span;
                    done += (i.State == "Failed" ? 1.0 : i.Fraction) * span;
                }
                return total <= 0 ? 1.0 : done / total;
            }
        }

        void Event(string level, string text) { Events.Enqueue(new CollectorEvent { Level = level, Text = text }); }

        public Task RunAsync(List<WorkItem> items, CancellationToken ct)
        {
            _items = items;
            // Signatures are created on this (calling) thread, before the writer owns the connection.
            var ids = new Dictionary<string, long>(StringComparer.Ordinal);
            foreach (WorkItem i in items)
            {
                long id;
                if (!ids.TryGetValue(i.Signature, out id)) { id = _store.GetOrCreateSignature(i.Signature, i.Label); ids[i.Signature] = id; }
                i.SignatureId = id;
            }
            return Task.Run(() => Run(ct));
        }

        async Task Run(CancellationToken ct)
        {
            StartedMs = Time.NowMs();
            using (var stop = CancellationTokenSource.CreateLinkedTokenSource(ct))
            using (_graph = new GraphClient(_o, _token, _limiter, Events))
            using (var pages = new BlockingCollection<PageResult>(Math.Max(2, _o.MaxConcurrency * 2)))
            {
                _stop = stop;
                var queue = new ConcurrentQueue<WorkItem>(_items);
                Task writer = Task.Factory.StartNew(() => Write(pages), CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
                var workers = new List<Task>();
                for (int w = 0; w < Math.Max(1, _o.MaxConcurrency); w++) workers.Add(Task.Run(() => Work(queue, pages, stop)));
                try { await Task.WhenAll(workers).ConfigureAwait(false); }
                catch { }
                pages.CompleteAdding();
                await writer.ConfigureAwait(false);
                foreach (WorkItem i in _items) if (i.State == "Pending" || i.State == "Running") i.State = "Cancelled";
            }
            EndedMs = Time.NowMs();
        }

        async Task Work(ConcurrentQueue<WorkItem> queue, BlockingCollection<PageResult> pages, CancellationTokenSource stop)
        {
            WorkItem item;
            while (!stop.IsCancellationRequested && queue.TryDequeue(out item))
            {
                item.State = "Running";
                item.StartedMs = Time.NowMs();
                Interlocked.Increment(ref ItemsRunning);
                string what = item.Label + " [" + Time.GraphTime(item.StartMs, false) + " -> " + Time.GraphTime(item.EndMs, true) + "]";
                string url = _o.GraphRoot + "/admin/exchange/tracing/messageTraces?$filter=" + Uri.EscapeDataString(item.ODataFilter()) + "&$top=" + _o.PageSize.ToString(CultureInfo.InvariantCulture);
                long previousOldest = long.MaxValue;
                bool ordered = true;
                try
                {
                    while (url != null)
                    {
                        string body = await _graph.GetAsync(url, what, stop.Token).ConfigureAwait(false);
                        long fetched = Time.NowMs();
                        string next;
                        List<TraceRow> rows = TraceParser.Parse(body, out next);
                        long oldest = long.MaxValue;
                        foreach (TraceRow r in rows)
                        {
                            if (r.ReceivedMs > previousOldest) ordered = false;   // the API is expected to return the newest first
                            if (r.ReceivedMs < oldest) oldest = r.ReceivedMs;
                        }
                        if (rows.Count > 0) { previousOldest = Math.Min(previousOldest, oldest); item.OldestMs = Math.Max(item.StartMs, previousOldest); }
                        Interlocked.Increment(ref Pages);
                        Interlocked.Add(ref Rows, rows.Count);
                        pages.Add(new PageResult { Item = item, Rows = rows, Last = next == null, Ordered = ordered, OldestMs = rows.Count > 0 ? oldest : 0, FetchedMs = fetched });
                        url = next;
                    }
                }
                catch (FatalCollectionException ex)
                {
                    Fatal = ex.Message; FatalStatus = ex.Status;
                    Event("ERROR", ex.Message);
                    pages.Add(new PageResult { Item = item, Error = ex.Message, Unrecoverable = true, Rows = new List<TraceRow>() });
                    stop.Cancel();
                }
                catch (ItemFailureException ex)
                {
                    pages.Add(new PageResult { Item = item, Error = ex.Message, Unrecoverable = ex.Unrecoverable, Rows = new List<TraceRow>() });
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    pages.Add(new PageResult { Item = item, Error = what + ": " + ex.Message, Rows = new List<TraceRow>() });
                }
                finally { Interlocked.Decrement(ref ItemsRunning); }
            }
        }

        void Write(BlockingCollection<PageResult> pages)
        {
            foreach (PageResult p in pages.GetConsumingEnumerable())
            {
                WorkItem item = p.Item;
                try
                {
                    if (p.Error != null)
                    {
                        _store.FailInterval(item, p.Error);
                        item.State = "Failed"; item.Error = p.Error; item.Unrecoverable = p.Unrecoverable; item.EndedMs = Time.NowMs();
                        Interlocked.Increment(ref ItemsFailed);
                        if (Fatal == null) Event(p.Unrecoverable ? "WARN" : "ERROR", p.Error);
                        continue;
                    }
                    IngestResult r = _store.CommitPage(item, p.Rows, p.Last, p.Ordered, p.OldestMs, p.FetchedMs, _runId);
                    Interlocked.Add(ref NewMessages, r.NewMessages);
                    Interlocked.Add(ref NewRows, r.NewRows);
                    item.Pages++;
                    item.Rows += p.Rows.Count;
                    item.NewMessages += r.NewMessages;
                    if (p.Last) { item.State = "Done"; item.EndedMs = Time.NowMs(); Interlocked.Increment(ref ItemsDone); }
                }
                catch (Exception ex)
                {
                    // A database error is fatal: nothing more can be stored.
                    Fatal = "Database error: " + ex.Message;
                    item.State = "Failed"; item.Error = Fatal; item.EndedMs = Time.NowMs();
                    Event("ERROR", Fatal);
                    try { _stop.Cancel(); } catch (ObjectDisposedException) { }
                    foreach (PageResult rest in pages.GetConsumingEnumerable()) { }
                    return;
                }
            }
        }

        public long[] LimiterStamps() { return _limiter.Stamps(); }
    }

    /// <summary>
    /// One message whose routes are read with getDetailsByRecipient: its recipient count before
    /// expansion when it is not known yet and, for a message over the limit, which recipients the
    /// sender addressed (recipient list before expansion).
    /// </summary>
    public sealed class CountItem
    {
        public long MessageKey;
        public string TraceId;
        public int Limit;                                    // Target.RecipientLimit
        public bool RebuildList;                             // tell the recipients apart when the count is over the limit
        public int Budget = 2;                               // route reads allowed for this message in this run
        public int Unclassified;                             // recipients whose role is not known yet
        public int DirectKnown;                              // recipients already known as addressed by the sender
        public List<string> Targets = new List<string>();   // distribution lists first, then the other recipients, likely list members last
        public long? Count;                                  // known before the run, or read
        public string Source;                                // Submit | Receive
        public string State = "Pending";                     // Pending | Done | NoCount | Failed | Cancelled
        public string ListState;                             // null (not decided) | Done | Partial
        public string Error;
        public int Requests;
        public readonly Dictionary<string, CountEvidence> Routes = new Dictionary<string, CountEvidence>(StringComparer.OrdinalIgnoreCase);

        /// <summary>A recipient was addressed by the sender when its route holds the event that carries the count.</summary>
        public static bool IsDirect(CountEvidence route, string source) { return source == "Receive" ? route.HasReceive : route.HasSubmit; }

        public int DirectFound { get { return DirectKnown + Routes.Values.Count(r => IsDirect(r, Source)); } }

        /// <summary>
        /// Recipient list state once the routes are read. Done: every recipient the sender addressed is
        /// known (as many as the count, or every route read). Partial: the budget ran out first.
        /// </summary>
        public void Conclude()
        {
            if (!RebuildList || !Count.HasValue || Count.Value <= Limit) return;
            if (DirectFound >= Count.Value || Routes.Count >= Unclassified) ListState = "Done";
            else if (Requests >= Budget || Targets.Count < Unclassified) ListState = "Partial";
        }
    }

    /// <summary>
    /// Reads the routes of the messages sent to distribution lists (getDetailsByRecipient, separate quota
    /// of 100 requests / 5 min). The route of a distribution list always holds the Submit event: one
    /// request gives the count before expansion. For a message over the limit, the route of each other
    /// recipient is read until every recipient the sender addressed is found (as many as the count):
    /// the members added by the expansion of the lists are left out of the recipient list.
    /// </summary>
    public sealed class CountCollector
    {
        readonly RlrStore _store;
        readonly CollectorOptions _o;
        readonly TokenSlot _token;
        readonly RateLimiter _limiter;
        GraphClient _graph;
        List<CountItem> _items = new List<CountItem>();
        public readonly ConcurrentQueue<CollectorEvent> Events = new ConcurrentQueue<CollectorEvent>();
        public long Done, NoCount, Failed, ListsDone, ListsPartial;
        public string Fatal;

        public CountCollector(RlrStore store, CollectorOptions options, TokenSlot token, RateLimiter limiter)
        {
            _store = store; _o = options; _token = token; _limiter = limiter;
        }

        public long Requests { get { return _graph == null ? 0 : Interlocked.Read(ref _graph.Requests); } }
        public long Throttled { get { return _graph == null ? 0 : Interlocked.Read(ref _graph.Throttled); } }
        public long Finished { get { return Interlocked.Read(ref Done) + Interlocked.Read(ref NoCount) + Interlocked.Read(ref Failed); } }
        public double Progress { get { return _items.Count == 0 ? 1.0 : (double)Finished / _items.Count; } }

        public Task RunAsync(List<CountItem> items, CancellationToken ct)
        {
            _items = items;
            return Task.Run(() => Run(ct));
        }

        async Task ReadRoutes(CountItem item, CancellationToken ct)
        {
            foreach (string target in item.Targets)
            {
                if (item.Count.HasValue && (!item.RebuildList || item.Count.Value <= item.Limit)) return;
                if (item.Count.HasValue && item.DirectFound >= item.Count.Value) return;
                // The count is in the Submit event, the same for every recipient the sender addressed: when
                // neither a list nor another recipient has it, the other routes do not have it either.
                if (!item.Count.HasValue && item.Routes.Count >= 2) return;
                if (item.Requests >= item.Budget) return;
                string url = _o.GraphRoot + "/admin/exchange/tracing/messageTraces/" + Uri.EscapeDataString(item.TraceId)
                    + "/getDetailsByRecipient(recipientAddress=" + Uri.EscapeDataString(Text.ODataString(target)) + ")";
                var route = new CountEvidence();
                while (url != null)
                {
                    string body = await _graph.GetAsync(url, "Route of a message to " + target, ct).ConfigureAwait(false);
                    item.Requests++;
                    string next;
                    CountEvidence e = TraceParser.ParseCount(body, out next);
                    route.Events += e.Events;
                    route.HasSubmit |= e.HasSubmit;
                    route.HasReceive |= e.HasReceive;
                    if (!route.Count.HasValue && e.Count.HasValue) { route.Count = e.Count; route.Source = e.Source; }
                    url = route.HasSubmit ? null : next;
                }
                item.Routes[target] = route;
                if (!item.Count.HasValue && route.Count.HasValue) { item.Count = route.Count; item.Source = route.Source; }
            }
        }

        async Task Run(CancellationToken ct)
        {
            using (var stop = CancellationTokenSource.CreateLinkedTokenSource(ct))
            using (_graph = new GraphClient(_o, _token, _limiter, Events))
            using (var results = new BlockingCollection<CountItem>(Math.Max(2, _o.MaxConcurrency * 4)))
            {
                var queue = new ConcurrentQueue<CountItem>(_items);
                Task writer = Task.Factory.StartNew(() =>
                {
                    foreach (CountItem r in results.GetConsumingEnumerable())
                    {
                        try { _store.SaveCount(r); }
                        catch (Exception ex) { Fatal = "Database error: " + ex.Message; Events.Enqueue(new CollectorEvent { Level = "ERROR", Text = Fatal }); try { stop.Cancel(); } catch (ObjectDisposedException) { } }
                        if (r.State == "Cancelled") continue;
                        if (r.State == "Done") Interlocked.Increment(ref Done);
                        else if (r.State == "NoCount") Interlocked.Increment(ref NoCount);
                        else Interlocked.Increment(ref Failed);
                        if (r.ListState == "Done") Interlocked.Increment(ref ListsDone);
                        else if (r.ListState == "Partial") Interlocked.Increment(ref ListsPartial);
                    }
                }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
                // Routes already read are kept when the run stops in the middle of a message.
                Action<CountItem> keep = item =>
                {
                    if (item.Requests == 0) return;
                    item.State = "Cancelled";
                    try { results.Add(item); } catch (InvalidOperationException) { }
                };
                var workers = new List<Task>();
                for (int w = 0; w < Math.Max(1, _o.MaxConcurrency); w++)
                {
                    workers.Add(Task.Run(async () =>
                    {
                        CountItem item;
                        while (!stop.IsCancellationRequested && queue.TryDequeue(out item))
                        {
                            try
                            {
                                await ReadRoutes(item, stop.Token).ConfigureAwait(false);
                                item.State = item.Count.HasValue ? "Done" : "NoCount";
                                item.Conclude();
                                results.Add(item);
                            }
                            catch (FatalCollectionException ex) { Fatal = ex.Message; Events.Enqueue(new CollectorEvent { Level = "ERROR", Text = ex.Message }); stop.Cancel(); keep(item); }
                            catch (OperationCanceledException) { keep(item); }
                            catch (Exception ex)
                            {
                                item.State = "Failed"; item.Error = ex.Message;
                                Events.Enqueue(new CollectorEvent { Level = "WARN", Text = ex.Message });
                                try { results.Add(item); } catch (InvalidOperationException) { }
                            }
                        }
                    }));
                }
                try { await Task.WhenAll(workers).ConfigureAwait(false); } catch { }
                results.CompleteAdding();
                await writer.ConfigureAwait(false);
            }
        }
    }

    // -------------------------------------------------------------------------
    // 3. SQLite store
    // -------------------------------------------------------------------------

    public sealed class IngestResult { public long NewMessages, NewRows, UpdatedRows; }

    public sealed class StoreStatistics
    {
        public long FileBytes, Messages, Matches, Pending, ListsPending, Unmeasurable, Runs, Intervals;
        public long? FirstCoveredMs, LastCoveredMs, LastSuccessfulCollectionMs;
        public int SchemaVersion;
        public int? CompactionLimit;
        public string CreatedByVersion;
    }

    public sealed class DayStatistics
    {
        public DateTime LocalDate;
        public long StartMs, EndMs, Matches, Pending, CoveredMs, FinalCoveredMs;   // Pending: counts and recipient lists still to read
    }

    /// <summary>What the database holds for a period, before the report is written.</summary>
    public sealed class PeriodSummary
    {
        public long Messages;        // messages stored (all of them until the period is compacted)
        public long OverLimit;       // more than the limit in the trace (after expansion)
        public long WithLists;       // ... of which sent to at least one distribution list
        public long Measured;        // ... count before expansion read
        public long Pending;         // ... not read yet (will be read by the next run)
        public long Unmeasurable;    // ... no count found, or too many failed attempts
        public long Matches;         // more than the limit before expansion (trace IDs)
        public long ListsRebuilt;    // matches with distribution lists whose recipient list before expansion is known
        public long ListsPending;    // ... still to rebuild (next run)
        public long ListsNotRebuilt; // ... not rebuilt (MaxRoutesPerMessage reached, failures, or rebuild off): listed after expansion
    }

    public sealed class PurgeResult { public long Messages, Rows, Intervals; }

    public sealed class RlrStore : IDisposable
    {
        public const int SchemaVersion = 1;
        // Recipient count before expansion: the trace recipients when no distribution list was expanded,
        // otherwise the count read from the route of the message (null until it is read).
        public const string CountSql = "(CASE WHEN m.list_rows = 0 THEN m.recipient_rows ELSE m.count_value END)";
        // Messages whose count before expansion is still to read, and messages over the limit whose recipient
        // list before expansion is still to rebuild ($l limit, $a attempts, $r MaxRoutesPerMessage).
        const string CountPendingSql = "(m.recipient_rows > $l AND m.list_rows > 0 AND m.count_value IS NULL AND COALESCE(m.count_state, '') <> 'NoCount' AND m.count_attempts < $a)";
        const string ListPendingSql = "($r > 0 AND m.list_rows > 0 AND m.count_value > $l AND m.list_state IS NULL AND m.count_attempts < $a)";
        readonly SqliteConnection _db;
        readonly bool _readOnly;
        readonly Dictionary<string, long> _addresses = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, long> _statuses = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        public string DatabasePath { get; private set; }

        /// <summary>
        /// Route reads allowed per message to rebuild its recipient list before expansion
        /// (Counting.MaxRoutesPerMessage; 0 = the list is not rebuilt, only the count is read).
        /// </summary>
        public int MaxRoutesPerMessage = 250;

        public RlrStore(string path, bool readOnly, string toolVersion)
        {
            DatabasePath = Path.GetFullPath(path);
            _readOnly = readOnly;
            if (!readOnly) Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath));
            var builder = new SqliteConnectionStringBuilder
            {
                DataSource = DatabasePath,
                Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
                Cache = SqliteCacheMode.Private,
                Pooling = false          // release the file as soon as the store is disposed
            };
            _db = new SqliteConnection(builder.ToString());
            _db.Open();
            Exec("PRAGMA busy_timeout=60000;");
            Exec("PRAGMA cache_size=-65536;");   // 64 MB page cache
            Exec("PRAGMA temp_store=MEMORY;");
            if (readOnly)
            {
                long version = Scalar("PRAGMA user_version;");
                if (version != SchemaVersion)
                    throw new InvalidOperationException("The database schema version is " + version + "; this tool expects " + SchemaVersion + ".");
            }
            else Migrate(toolVersion);
        }

        public void Dispose()
        {
            if (!_readOnly)
            {
                try { Exec("PRAGMA optimize;"); Exec("PRAGMA wal_checkpoint(TRUNCATE);"); } catch (SqliteException) { }
            }
            _db.Dispose();
        }

        // ---- Schema ---------------------------------------------------------

        void Migrate(string toolVersion)
        {
            long version = Scalar("PRAGMA user_version;");
            if (version > SchemaVersion)
                throw new InvalidOperationException("The database was created by a newer version of the tool (schema " + version + ").");
            if (version == 0)
            {
                // auto_vacuum must be chosen before the first table is created.
                Exec("PRAGMA auto_vacuum=INCREMENTAL;");
                Exec("PRAGMA journal_mode=WAL;");
                using (SqliteTransaction tx = _db.BeginTransaction())
                {
                    using (SqliteCommand c = Command(@"
CREATE TABLE metadata (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
-- One row per execution of the tool.
CREATE TABLE run (
    run_id       INTEGER PRIMARY KEY,
    started_ms   INTEGER NOT NULL,
    finished_ms  INTEGER,
    mode         TEXT    NOT NULL,
    status       TEXT    NOT NULL,
    account      TEXT,
    host         TEXT,
    tool_version TEXT,
    details      TEXT
);
-- One row per kind of message trace query: '' = every sender, 'sender=*@contoso.com' = one domain.
CREATE TABLE signature (
    signature_id INTEGER PRIMARY KEY,
    text         TEXT NOT NULL UNIQUE,
    label        TEXT
);
-- Time slices requested from the message trace. The API returns the newest rows first:
-- [covered_from_ms, end_ms) is stored, page after page. status = Running | Completed | Failed.
CREATE TABLE collection_interval (
    interval_id     INTEGER PRIMARY KEY,
    run_id          INTEGER NOT NULL REFERENCES run(run_id),
    signature_id    INTEGER NOT NULL REFERENCES signature(signature_id),
    start_ms        INTEGER NOT NULL,
    end_ms          INTEGER NOT NULL,
    covered_from_ms INTEGER NOT NULL,
    status          TEXT    NOT NULL,
    started_ms      INTEGER NOT NULL,
    collected_ms    INTEGER NOT NULL,
    completed_ms    INTEGER,
    pages           INTEGER NOT NULL DEFAULT 0,
    rows            INTEGER NOT NULL DEFAULT 0,
    new_messages    INTEGER NOT NULL DEFAULT 0,
    error           TEXT
);
CREATE INDEX ix_interval_signature ON collection_interval (signature_id, end_ms);
-- Every sender and recipient address, once (case-insensitive).
CREATE TABLE address (
    address_id INTEGER PRIMARY KEY,
    address    TEXT NOT NULL UNIQUE COLLATE NOCASE
);
-- Delivery status names of the message trace (delivered, failed, expanded ...).
CREATE TABLE status (
    status_id INTEGER PRIMARY KEY,
    name      TEXT NOT NULL UNIQUE COLLATE NOCASE
);
-- One row per message trace ID. recipient_rows / list_rows are kept up to date at each page.
-- count_* : recipient count before expansion read from the route (messages with distribution lists).
-- list_state: recipient list as the sender addressed it (messages over the limit with distribution
-- lists): NULL = not rebuilt yet, Done = every recipient told apart, Partial = Counting.MaxRoutesPerMessage
-- reached. route_reads: getDetailsByRecipient requests spent on the message.
CREATE TABLE message (
    message_key    INTEGER PRIMARY KEY,
    trace_id       BLOB    NOT NULL UNIQUE,
    message_id     TEXT,
    sender_id      INTEGER,
    subject        TEXT,
    received_ms    INTEGER NOT NULL,
    recipient_rows INTEGER NOT NULL DEFAULT 0,
    list_rows      INTEGER NOT NULL DEFAULT 0,
    count_value    INTEGER,
    count_source   TEXT,
    count_state    TEXT,
    count_attempts INTEGER NOT NULL DEFAULT 0,
    count_ms       INTEGER,
    count_error    TEXT,
    list_state     TEXT,
    route_reads    INTEGER NOT NULL DEFAULT 0,
    first_seen_ms  INTEGER NOT NULL,
    updated_ms     INTEGER NOT NULL
);
CREATE INDEX ix_message_received ON message (received_ms);
CREATE INDEX ix_message_mid      ON message (message_id);
-- One row per message and recipient (what the message trace returns). direct: NULL = not known,
-- 1 = addressed by the sender (its route holds the Submit or Receive event), 0 = added by the
-- expansion of a distribution list.
CREATE TABLE recipient (
    message_key INTEGER NOT NULL,
    address_id  INTEGER NOT NULL,
    status_id   INTEGER NOT NULL,
    direct      INTEGER,
    PRIMARY KEY (message_key, address_id)
) WITHOUT ROWID;
CREATE INDEX ix_recipient_direct ON recipient (address_id, direct) WHERE direct IS NOT NULL;
", tx)) c.ExecuteNonQuery();
                    SetMetadata("created_utc", DateTimeOffset.UtcNow.ToString("o"), tx);
                    SetMetadata("created_by_version", toolVersion ?? "", tx);
                    using (SqliteCommand c = Command("PRAGMA user_version=" + SchemaVersion + ";", tx)) c.ExecuteNonQuery();
                    tx.Commit();
                }
            }
            Exec("PRAGMA journal_mode=WAL;");
            Exec("PRAGMA synchronous=NORMAL;");
            SetMetadata("last_opened_by_version", toolVersion ?? "", null);
        }

        // ---- Helpers --------------------------------------------------------

        SqliteCommand Command(string sql, SqliteTransaction tx = null)
        {
            SqliteCommand c = _db.CreateCommand();
            c.CommandText = sql;
            c.CommandTimeout = 0;
            c.Transaction = tx;
            return c;
        }

        void Exec(string sql)
        {
            using (SqliteCommand c = Command(sql)) c.ExecuteNonQuery();
        }

        long Scalar(string sql, params object[] args)
        {
            using (SqliteCommand c = Command(sql))
            {
                for (int i = 0; i < args.Length; i++) c.Parameters.AddWithValue("$p" + i, args[i] ?? DBNull.Value);
                object v = c.ExecuteScalar();
                return v == null || v is DBNull ? 0 : Convert.ToInt64(v, CultureInfo.InvariantCulture);
            }
        }

        static long? NullableLong(SqliteDataReader r, int i) { return r.IsDBNull(i) ? (long?)null : r.GetInt64(i); }
        static long FileSize(string path) { return File.Exists(path) ? new FileInfo(path).Length : 0; }

        public static byte[] TraceKey(string traceId)
        {
            Guid g;
            if (Guid.TryParse(traceId, out g)) return g.ToByteArray();
            return Encoding.UTF8.GetBytes(traceId ?? "");
        }

        public static string TraceText(byte[] key)
        {
            if (key == null) return "";
            return key.Length == 16 ? new Guid(key).ToString() : Encoding.UTF8.GetString(key);
        }

        /// <summary>SQL condition on the sender domain ('' when every sender is reported). Parameters $d0, $d1 ...</summary>
        static string DomainClause(string[] domains)
        {
            if (domains == null || domains.Length == 0) return "";
            return " AND (" + string.Join(" OR ", domains.Select((d, i) => "sa.address LIKE $d" + i)) + ")";
        }

        static void AddDomains(SqliteCommand c, string[] domains)
        {
            if (domains == null) return;
            for (int i = 0; i < domains.Length; i++) c.Parameters.AddWithValue("$d" + i, "%@" + domains[i]);
        }

        static string SenderJoin(string[] domains)
        {
            return domains == null || domains.Length == 0 ? "" : " JOIN address sa ON sa.address_id = m.sender_id";
        }

        // ---- Metadata and runs ------------------------------------------------

        void SetMetadata(string key, string value, SqliteTransaction tx)
        {
            using (SqliteCommand c = Command("INSERT INTO metadata(key, value) VALUES($k, $v) ON CONFLICT(key) DO UPDATE SET value = excluded.value;", tx))
            {
                c.Parameters.AddWithValue("$k", key);
                c.Parameters.AddWithValue("$v", value ?? "");
                c.ExecuteNonQuery();
            }
        }

        public void SetMetadata(string key, string value) { SetMetadata(key, value, null); }

        public string GetMetadata(string key)
        {
            using (SqliteCommand c = Command("SELECT value FROM metadata WHERE key = $k;"))
            {
                c.Parameters.AddWithValue("$k", key);
                object o = c.ExecuteScalar();
                return o == null || o is DBNull ? null : (string)o;
            }
        }

        /// <summary>Request times of the last quota window (saved by the previous run on this database).</summary>
        public long[] GetRequestStamps(string key)
        {
            string text = GetMetadata(key);
            if (string.IsNullOrEmpty(text)) return new long[0];
            var list = new List<long>();
            foreach (string part in text.Split(',')) { long v; if (long.TryParse(part, NumberStyles.Integer, CultureInfo.InvariantCulture, out v)) list.Add(v); }
            return list.ToArray();
        }

        public void SetRequestStamps(string key, long[] stamps)
        {
            SetMetadata(key, string.Join(",", stamps.Select(s => s.ToString(CultureInfo.InvariantCulture))));
        }

        public long StartRun(string mode, string account, string host, string toolVersion, string details)
        {
            using (SqliteCommand c = Command("INSERT INTO run(started_ms, mode, status, account, host, tool_version, details) VALUES($s, $m, 'Running', $a, $h, $v, $d) RETURNING run_id;"))
            {
                c.Parameters.AddWithValue("$s", Time.NowMs());
                c.Parameters.AddWithValue("$m", mode);
                c.Parameters.AddWithValue("$a", (object)account ?? DBNull.Value);
                c.Parameters.AddWithValue("$h", (object)host ?? DBNull.Value);
                c.Parameters.AddWithValue("$v", (object)toolVersion ?? DBNull.Value);
                c.Parameters.AddWithValue("$d", (object)details ?? DBNull.Value);
                return (long)c.ExecuteScalar();
            }
        }

        public void UpdateRunAccount(long runId, string account)
        {
            using (SqliteCommand c = Command("UPDATE run SET account = $a WHERE run_id = $r;"))
            {
                c.Parameters.AddWithValue("$a", (object)account ?? DBNull.Value);
                c.Parameters.AddWithValue("$r", runId);
                c.ExecuteNonQuery();
            }
        }

        public void FinishRun(long runId, string status, string details)
        {
            using (SqliteCommand c = Command("UPDATE run SET finished_ms = $f, status = $s, details = COALESCE($d, details) WHERE run_id = $r;"))
            {
                c.Parameters.AddWithValue("$f", Time.NowMs());
                c.Parameters.AddWithValue("$s", status);
                c.Parameters.AddWithValue("$d", (object)details ?? DBNull.Value);
                c.Parameters.AddWithValue("$r", runId);
                c.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Slices left 'Running' by an interrupted execution are marked as failed. What they stored is
        /// kept and counted as collected (the coverage advances page by page with the data).
        /// </summary>
        public int CloseAbandonedWork()
        {
            int n;
            using (SqliteCommand c = Command("UPDATE collection_interval SET status = 'Failed', error = 'Interrupted (the previous execution stopped before the end of this slice)' WHERE status = 'Running';"))
                n = c.ExecuteNonQuery();
            using (SqliteCommand c = Command("UPDATE run SET status = 'Interrupted' WHERE status = 'Running';"))
                c.ExecuteNonQuery();
            return n;
        }

        // ---- Signatures and coverage ---------------------------------------------

        public long GetOrCreateSignature(string text, string label)
        {
            using (SqliteCommand c = Command("INSERT INTO signature(text, label) VALUES($t, $l) ON CONFLICT(text) DO UPDATE SET label = COALESCE(excluded.label, signature.label) RETURNING signature_id;"))
            {
                c.Parameters.AddWithValue("$t", text ?? "");
                c.Parameters.AddWithValue("$l", (object)label ?? DBNull.Value);
                return (long)c.ExecuteScalar();
            }
        }

        /// <summary>
        /// Ranges collected for one query signature. A query on every sender ('') also covers each
        /// sender domain. When settlingMs is greater than zero, the part of each range that was less
        /// than settlingMs old when it was collected is left out ("settled" part only): the message
        /// trace still changes for a few hours (late rows, status changes), so those hours are
        /// collected again by the next run.
        /// </summary>
        public List<TimeRange> GetCoveredRanges(string signature, long settlingMs)
        {
            var list = new List<TimeRange>();
            using (SqliteCommand c = Command(@"SELECT i.covered_from_ms, i.end_ms, i.collected_ms FROM collection_interval i
JOIN signature s ON s.signature_id = i.signature_id WHERE s.text = $t OR s.text = '';"))
            {
                c.Parameters.AddWithValue("$t", signature ?? "");
                using (SqliteDataReader r = c.ExecuteReader())
                    while (r.Read())
                    {
                        long start = r.GetInt64(0), end = r.GetInt64(1), collected = r.GetInt64(2);
                        if (settlingMs > 0) end = Math.Min(end, collected - settlingMs);
                        if (end > start) list.Add(new TimeRange(start, end));
                    }
            }
            return Coverage.Merge(list);
        }

        /// <summary>Ranges collected for EVERY signature of the list (an empty list = every sender).</summary>
        public List<TimeRange> GetCoverage(string[] signatures, long settlingMs)
        {
            if (signatures == null || signatures.Length == 0) signatures = new[] { "" };
            List<TimeRange> result = null;
            foreach (string s in signatures)
            {
                List<TimeRange> ranges = GetCoveredRanges(s, settlingMs);
                result = result == null ? ranges : Coverage.Intersect(result, ranges);
            }
            return result ?? new List<TimeRange>();
        }

        // ---- Ingestion ------------------------------------------------------

        long AddressId(string address, SqliteTransaction tx)
        {
            address = (address ?? "").Trim();
            long id;
            if (_addresses.TryGetValue(address, out id)) return id;
            using (SqliteCommand c = Command("INSERT INTO address(address) VALUES($a) ON CONFLICT(address) DO UPDATE SET address = address RETURNING address_id;", tx))
            {
                c.Parameters.AddWithValue("$a", address);
                id = Convert.ToInt64(c.ExecuteScalar(), CultureInfo.InvariantCulture);
            }
            if (_addresses.Count > 500000) _addresses.Clear();
            _addresses[address] = id;
            return id;
        }

        long StatusId(string name, SqliteTransaction tx)
        {
            name = (name ?? "").Trim();
            long id;
            if (_statuses.TryGetValue(name, out id)) return id;
            using (SqliteCommand c = Command("INSERT INTO status(name) VALUES($n) ON CONFLICT(name) DO UPDATE SET name = name RETURNING status_id;", tx))
            {
                c.Parameters.AddWithValue("$n", name);
                id = Convert.ToInt64(c.ExecuteScalar(), CultureInfo.InvariantCulture);
            }
            _statuses[name] = id;
            return id;
        }

        /// <summary>
        /// Stores one page and extends the coverage of its slice, in one transaction. The API returns
        /// the newest rows first: after a page, the slice is complete from just after the oldest row of
        /// the page to its end. After the last page, the whole slice is complete. Re-reading a page is
        /// harmless: a message is one row per trace ID and a recipient one row per message and address.
        /// </summary>
        public IngestResult CommitPage(WorkItem item, List<TraceRow> rows, bool last, bool ordered, long oldestMs, long fetchedMs, long runId)
        {
            var result = new IngestResult();
            long now = Time.NowMs();
            using (SqliteTransaction tx = _db.BeginTransaction())
            {
                if (item.IntervalId == 0)
                {
                    using (SqliteCommand c = Command(@"INSERT INTO collection_interval(run_id, signature_id, start_ms, end_ms, covered_from_ms, status, started_ms, collected_ms)
VALUES($r, $s, $a, $b, $b, 'Running', $st, $c) RETURNING interval_id;", tx))
                    {
                        c.Parameters.AddWithValue("$r", runId); c.Parameters.AddWithValue("$s", item.SignatureId);
                        c.Parameters.AddWithValue("$a", item.StartMs); c.Parameters.AddWithValue("$b", item.EndMs);
                        c.Parameters.AddWithValue("$st", item.StartedMs > 0 ? item.StartedMs : now); c.Parameters.AddWithValue("$c", fetchedMs);
                        item.IntervalId = (long)c.ExecuteScalar();
                    }
                    item.CoveredFromMs = item.EndMs;
                }
                long expanded = StatusId("expanded", tx);
                var touched = new HashSet<long>();
                using (SqliteCommand insMsg = Command(@"INSERT INTO message(trace_id, message_id, sender_id, subject, received_ms, first_seen_ms, updated_ms)
VALUES($t, $m, $s, $j, $r, $n, $n) ON CONFLICT(trace_id) DO NOTHING RETURNING message_key;", tx))
                using (SqliteCommand getMsg = Command("SELECT message_key FROM message WHERE trace_id = $t;", tx))
                using (SqliteCommand insRcp = Command("INSERT INTO recipient(message_key, address_id, status_id) VALUES($m, $a, $s) ON CONFLICT(message_key, address_id) DO NOTHING;", tx))
                using (SqliteCommand updRcp = Command("UPDATE recipient SET status_id = $s WHERE message_key = $m AND address_id = $a AND status_id <> $s;", tx))
                using (SqliteCommand count = Command(@"UPDATE message SET
    recipient_rows = (SELECT COUNT(*) FROM recipient WHERE message_key = $m),
    list_rows      = (SELECT COUNT(*) FROM recipient WHERE message_key = $m AND status_id = $x),
    updated_ms     = $n
WHERE message_key = $m;", tx))
                {
                    SqliteParameter mT = insMsg.Parameters.Add("$t", SqliteType.Blob), mM = insMsg.Parameters.Add("$m", SqliteType.Text);
                    SqliteParameter mS = insMsg.Parameters.Add("$s", SqliteType.Integer), mJ = insMsg.Parameters.Add("$j", SqliteType.Text);
                    SqliteParameter mR = insMsg.Parameters.Add("$r", SqliteType.Integer);
                    insMsg.Parameters.AddWithValue("$n", now);
                    SqliteParameter gT = getMsg.Parameters.Add("$t", SqliteType.Blob);
                    SqliteParameter iM = insRcp.Parameters.Add("$m", SqliteType.Integer), iA = insRcp.Parameters.Add("$a", SqliteType.Integer), iS = insRcp.Parameters.Add("$s", SqliteType.Integer);
                    SqliteParameter uM = updRcp.Parameters.Add("$m", SqliteType.Integer), uA = updRcp.Parameters.Add("$a", SqliteType.Integer), uS = updRcp.Parameters.Add("$s", SqliteType.Integer);
                    SqliteParameter cM = count.Parameters.Add("$m", SqliteType.Integer);
                    count.Parameters.AddWithValue("$x", expanded);
                    count.Parameters.AddWithValue("$n", now);
                    var keys = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
                    foreach (TraceRow row in rows)
                    {
                        if (string.IsNullOrEmpty(row.Id) || string.IsNullOrWhiteSpace(row.Recipient)) continue;
                        long key;
                        if (!keys.TryGetValue(row.Id, out key))
                        {
                            byte[] traceKey = TraceKey(row.Id);
                            mT.Value = traceKey; mM.Value = (object)row.MessageId ?? DBNull.Value; mS.Value = AddressId(row.Sender, tx);
                            mJ.Value = (object)row.Subject ?? DBNull.Value; mR.Value = row.ReceivedMs;
                            object id = insMsg.ExecuteScalar();
                            if (id != null && !(id is DBNull)) { key = Convert.ToInt64(id, CultureInfo.InvariantCulture); result.NewMessages++; }
                            else { gT.Value = traceKey; key = Convert.ToInt64(getMsg.ExecuteScalar(), CultureInfo.InvariantCulture); }
                            keys[row.Id] = key;
                        }
                        long address = AddressId(row.Recipient, tx), status = StatusId(row.Status, tx);
                        iM.Value = key; iA.Value = address; iS.Value = status;
                        if (insRcp.ExecuteNonQuery() == 1) { result.NewRows++; touched.Add(key); continue; }
                        uM.Value = key; uA.Value = address; uS.Value = status;
                        if (updRcp.ExecuteNonQuery() == 1) { result.UpdatedRows++; touched.Add(key); }
                    }
                    foreach (long key in touched) { cM.Value = key; count.ExecuteNonQuery(); }
                }

                long from = item.CoveredFromMs;
                if (last) from = item.StartMs;
                else if (ordered && rows.Count > 0) from = Math.Min(from, Math.Max(item.StartMs, Math.Min(item.EndMs, oldestMs + 1)));
                using (SqliteCommand c = Command(@"UPDATE collection_interval SET covered_from_ms = $f, pages = pages + 1, rows = rows + $n, new_messages = new_messages + $nm,
    status = CASE WHEN $last = 1 THEN 'Completed' ELSE status END, completed_ms = CASE WHEN $last = 1 THEN $now ELSE completed_ms END
WHERE interval_id = $id;", tx))
                {
                    c.Parameters.AddWithValue("$f", from); c.Parameters.AddWithValue("$n", rows.Count); c.Parameters.AddWithValue("$nm", result.NewMessages);
                    c.Parameters.AddWithValue("$last", last ? 1 : 0); c.Parameters.AddWithValue("$now", now); c.Parameters.AddWithValue("$id", item.IntervalId);
                    c.ExecuteNonQuery();
                }
                tx.Commit();
                item.CoveredFromMs = from;
            }
            return result;
        }

        public void FailInterval(WorkItem item, string error)
        {
            if (item.IntervalId == 0) return;
            using (SqliteCommand c = Command("UPDATE collection_interval SET status = 'Failed', error = $e WHERE interval_id = $id;"))
            {
                c.Parameters.AddWithValue("$e", (object)error ?? DBNull.Value);
                c.Parameters.AddWithValue("$id", item.IntervalId);
                c.ExecuteNonQuery();
            }
        }

        // ---- Recipient count and recipient list before expansion ----------------------------

        /// <summary>
        /// Messages of [startMs, endMs) with at least one distribution list whose routes must be read:
        /// count before expansion still unknown (more than 'limit' recipients after expansion), or count
        /// over the limit with a recipient list not rebuilt yet (MaxRoutesPerMessage > 0). Each item lists
        /// the recipients whose route can be read: distribution lists first (the route of a list always
        /// holds the Submit event), then the other recipients not told apart yet, the addresses already
        /// seen as members of the same lists in other messages last (the reading stops as soon as every
        /// recipient the sender addressed is found, so they are rarely read).
        /// </summary>
        public List<CountItem> GetCountItems(long startMs, long endMs, int limit, int maxAttempts, int max, string[] domains)
        {
            bool rebuild = MaxRoutesPerMessage > 0;
            var items = new List<CountItem>();
            var reads = new Dictionary<long, long>();
            string sql = "SELECT m.message_key, m.trace_id, m.count_value, m.count_source, m.route_reads FROM message m" + SenderJoin(domains) + @"
WHERE m.received_ms >= $s AND m.received_ms < $e AND (" + CountPendingSql + " OR " + ListPendingSql + ")" + DomainClause(domains) + @"
ORDER BY m.received_ms DESC" + (max > 0 ? " LIMIT " + max.ToString(CultureInfo.InvariantCulture) : "") + ";";
            using (SqliteCommand c = Command(sql))
            {
                c.Parameters.AddWithValue("$s", startMs); c.Parameters.AddWithValue("$e", endMs);
                c.Parameters.AddWithValue("$l", limit); c.Parameters.AddWithValue("$a", maxAttempts); c.Parameters.AddWithValue("$r", MaxRoutesPerMessage);
                AddDomains(c, domains);
                using (SqliteDataReader r = c.ExecuteReader())
                    while (r.Read())
                    {
                        var item = new CountItem
                        {
                            MessageKey = r.GetInt64(0), TraceId = TraceText((byte[])r.GetValue(1)), Limit = limit, RebuildList = rebuild,
                            Count = NullableLong(r, 2), Source = r.IsDBNull(3) ? null : r.GetString(3)
                        };
                        reads[item.MessageKey] = r.GetInt64(4);
                        items.Add(item);
                    }
            }
            long expanded = Scalar("SELECT COALESCE((SELECT status_id FROM status WHERE name = 'expanded'), -1);");
            if (!rebuild)
            {
                // Count only: a distribution list, then one other recipient in case the first route has no count.
                using (SqliteCommand lists = Command("SELECT a.address FROM recipient r JOIN address a ON a.address_id = r.address_id WHERE r.message_key = $m AND r.status_id = $x ORDER BY a.address LIMIT 1;"))
                using (SqliteCommand other = Command("SELECT a.address FROM recipient r JOIN address a ON a.address_id = r.address_id WHERE r.message_key = $m AND r.status_id <> $x ORDER BY a.address LIMIT 1;"))
                {
                    SqliteParameter pl = lists.Parameters.Add("$m", SqliteType.Integer), po = other.Parameters.Add("$m", SqliteType.Integer);
                    lists.Parameters.AddWithValue("$x", expanded); other.Parameters.AddWithValue("$x", expanded);
                    foreach (CountItem item in items)
                    {
                        pl.Value = item.MessageKey; po.Value = item.MessageKey;
                        object a = lists.ExecuteScalar(), b = other.ExecuteScalar();
                        if (a is string) item.Targets.Add((string)a);
                        if (b is string) item.Targets.Add((string)b);
                    }
                }
                return items;
            }
            using (SqliteCommand known = Command("SELECT COUNT(*), COALESCE(SUM(direct = 1), 0) FROM recipient WHERE message_key = $m AND direct IS NOT NULL;"))
            using (SqliteCommand targets = Command(@"SELECT a.address FROM recipient x JOIN address a ON a.address_id = x.address_id
WHERE x.message_key = $m AND x.direct IS NULL
ORDER BY x.status_id = $x DESC,
    EXISTS (SELECT 1 FROM recipient l
            JOIN recipient l2 ON l2.address_id = l.address_id AND l2.direct = 1 AND l2.message_key <> l.message_key
            JOIN recipient r2 ON r2.message_key = l2.message_key AND r2.address_id = x.address_id AND r2.direct = 0
            WHERE l.message_key = x.message_key AND l.status_id = $x),
    a.address COLLATE NOCASE;"))
            using (SqliteCommand total = Command("SELECT COUNT(*) FROM recipient WHERE message_key = $m;"))
            {
                SqliteParameter pk = known.Parameters.Add("$m", SqliteType.Integer), pt = targets.Parameters.Add("$m", SqliteType.Integer), pa = total.Parameters.Add("$m", SqliteType.Integer);
                targets.Parameters.AddWithValue("$x", expanded);
                foreach (CountItem item in items)
                {
                    pk.Value = item.MessageKey; pt.Value = item.MessageKey; pa.Value = item.MessageKey;
                    long classified = 0;
                    using (SqliteDataReader r = known.ExecuteReader()) if (r.Read()) { classified = r.GetInt64(0); item.DirectKnown = (int)r.GetInt64(1); }
                    item.Unclassified = (int)(Convert.ToInt64(total.ExecuteScalar(), CultureInfo.InvariantCulture) - classified);
                    // The count, when still unknown, is read even when the budget of the message is spent.
                    item.Budget = (int)Math.Max(item.Count.HasValue ? 0 : 2, MaxRoutesPerMessage - reads[item.MessageKey]);
                    using (SqliteDataReader r = targets.ExecuteReader())
                        while (item.Targets.Count < item.Budget && r.Read()) item.Targets.Add(r.GetString(0));
                }
            }
            return items;
        }

        /// <summary>
        /// Saves what the routes of one message told: the count (never replaced once known), the role of
        /// each recipient whose route was read, the recipient list state, and one more attempt (not for a
        /// run stopped in the middle of the message).
        /// </summary>
        public void SaveCount(CountItem item)
        {
            using (SqliteTransaction tx = _db.BeginTransaction())
            {
                using (SqliteCommand c = Command(@"UPDATE message SET
    count_value    = COALESCE(count_value, $v),
    count_source   = CASE WHEN count_value IS NULL THEN $src ELSE count_source END,
    count_state    = CASE WHEN COALESCE(count_value, $v) IS NOT NULL THEN 'Done' WHEN $st = 'Cancelled' THEN count_state ELSE $st END,
    count_attempts = count_attempts + $inc,
    count_ms       = $now,
    count_error    = $e,
    route_reads    = route_reads + $rq,
    list_state     = COALESCE($ls, list_state)
WHERE message_key = $m;", tx))
                {
                    c.Parameters.AddWithValue("$v", item.Count.HasValue ? (object)item.Count.Value : DBNull.Value);
                    c.Parameters.AddWithValue("$src", (object)item.Source ?? DBNull.Value);
                    c.Parameters.AddWithValue("$st", item.State);
                    c.Parameters.AddWithValue("$inc", item.State == "Cancelled" ? 0 : 1);
                    c.Parameters.AddWithValue("$now", Time.NowMs());
                    c.Parameters.AddWithValue("$e", (object)item.Error ?? DBNull.Value);
                    c.Parameters.AddWithValue("$rq", item.Requests);
                    c.Parameters.AddWithValue("$ls", (object)item.ListState ?? DBNull.Value);
                    c.Parameters.AddWithValue("$m", item.MessageKey);
                    c.ExecuteNonQuery();
                }
                // A role is only known once the event that carries the count is known (Submit or Receive).
                if (item.Source != null && item.Routes.Count > 0)
                {
                    using (SqliteCommand c = Command("UPDATE recipient SET direct = $d WHERE message_key = $m AND address_id = (SELECT address_id FROM address WHERE address = $a);", tx))
                    {
                        SqliteParameter pd = c.Parameters.Add("$d", SqliteType.Integer), pa = c.Parameters.Add("$a", SqliteType.Text);
                        c.Parameters.AddWithValue("$m", item.MessageKey);
                        foreach (KeyValuePair<string, CountEvidence> route in item.Routes)
                        {
                            pd.Value = CountItem.IsDirect(route.Value, item.Source) ? 1 : 0;
                            pa.Value = route.Key;
                            c.ExecuteNonQuery();
                        }
                    }
                }
                // Every recipient the sender addressed is known: the others were added by the expansion.
                if (item.ListState == "Done")
                {
                    using (SqliteCommand c = Command("UPDATE recipient SET direct = 0 WHERE message_key = $m AND direct IS NULL;", tx))
                    {
                        c.Parameters.AddWithValue("$m", item.MessageKey);
                        c.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        public PeriodSummary GetPeriodSummary(long startMs, long endMs, int limit, int maxAttempts, string[] domains)
        {
            var s = new PeriodSummary();
            string sql = @"SELECT COUNT(*),
    COALESCE(SUM(m.recipient_rows > $l), 0),
    COALESCE(SUM(m.recipient_rows > $l AND m.list_rows > 0), 0),
    COALESCE(SUM(m.recipient_rows > $l AND m.list_rows > 0 AND m.count_value IS NOT NULL), 0),
    COALESCE(SUM(" + CountPendingSql + @"), 0),
    COALESCE(SUM(m.recipient_rows > $l AND m.list_rows > 0 AND m.count_value IS NULL AND (m.count_state = 'NoCount' OR m.count_attempts >= $a)), 0),
    COALESCE(SUM(" + CountSql + @" > $l), 0),
    COALESCE(SUM(m.list_rows > 0 AND m.count_value > $l AND m.list_state = 'Done'), 0),
    COALESCE(SUM(" + ListPendingSql + @"), 0),
    COALESCE(SUM(m.list_rows > 0 AND m.count_value > $l AND COALESCE(m.list_state, '') <> 'Done' AND NOT " + ListPendingSql + @"), 0)
FROM message m" + SenderJoin(domains) + " WHERE m.received_ms >= $s AND m.received_ms < $e" + DomainClause(domains) + ";";
            using (SqliteCommand c = Command(sql))
            {
                c.Parameters.AddWithValue("$s", startMs); c.Parameters.AddWithValue("$e", endMs);
                c.Parameters.AddWithValue("$l", limit); c.Parameters.AddWithValue("$a", maxAttempts); c.Parameters.AddWithValue("$r", MaxRoutesPerMessage);
                AddDomains(c, domains);
                using (SqliteDataReader r = c.ExecuteReader())
                    if (r.Read())
                    {
                        s.Messages = r.GetInt64(0); s.OverLimit = r.GetInt64(1); s.WithLists = r.GetInt64(2); s.Measured = r.GetInt64(3);
                        s.Pending = r.GetInt64(4); s.Unmeasurable = r.GetInt64(5); s.Matches = r.GetInt64(6);
                        s.ListsRebuilt = r.GetInt64(7); s.ListsPending = r.GetInt64(8); s.ListsNotRebuilt = r.GetInt64(9);
                    }
            }
            return s;
        }

        // ---- Compaction and retention ------------------------------------------------------

        const string NotReported = "(m.recipient_rows <= $l OR (m.list_rows > 0 AND m.count_value IS NOT NULL AND m.count_value <= $l))";

        /// <summary>
        /// Keeps only what a report can show. In ranges whose collection is final (settled), the
        /// messages sent to 'limit' recipients or fewer are deleted with their recipients: the
        /// database keeps the history of the messages over the limit (and of those still to be
        /// measured), not the whole message trace of the tenant.
        /// </summary>
        public long Compact(IEnumerable<TimeRange> settled, int limit)
        {
            long deleted = 0;
            using (SqliteTransaction tx = _db.BeginTransaction())
            {
                using (SqliteCommand rcp = Command("DELETE FROM recipient WHERE message_key IN (SELECT m.message_key FROM message m WHERE m.received_ms >= $s AND m.received_ms < $e AND " + NotReported + ");", tx))
                using (SqliteCommand msg = Command("DELETE FROM message AS m WHERE m.received_ms >= $s AND m.received_ms < $e AND " + NotReported + ";", tx))
                {
                    SqliteParameter rs = rcp.Parameters.Add("$s", SqliteType.Integer), re = rcp.Parameters.Add("$e", SqliteType.Integer);
                    SqliteParameter ms = msg.Parameters.Add("$s", SqliteType.Integer), me = msg.Parameters.Add("$e", SqliteType.Integer);
                    rcp.Parameters.AddWithValue("$l", limit); msg.Parameters.AddWithValue("$l", limit);
                    foreach (TimeRange range in Coverage.Merge(settled))
                    {
                        rs.Value = range.Start; re.Value = range.End; ms.Value = range.Start; me.Value = range.End;
                        rcp.ExecuteNonQuery();
                        deleted += msg.ExecuteNonQuery();
                    }
                }
                string current = null;
                using (SqliteCommand c = Command("SELECT value FROM metadata WHERE key = 'compaction_limit';", tx)) { object o = c.ExecuteScalar(); current = o as string; }
                int previous;
                if (current == null || !int.TryParse(current, NumberStyles.Integer, CultureInfo.InvariantCulture, out previous) || limit < previous)
                    SetMetadata("compaction_limit", limit.ToString(CultureInfo.InvariantCulture), tx);
                tx.Commit();
            }
            if (deleted > 0) Exec("PRAGMA incremental_vacuum;");
            return deleted;
        }

        /// <summary>Deletes everything older than cutoffMs and trims coverage accordingly.</summary>
        public PurgeResult PurgeBefore(long cutoffMs)
        {
            var p = new PurgeResult();
            using (SqliteTransaction tx = _db.BeginTransaction())
            {
                Func<string, long> run = sql => { using (SqliteCommand c = Command(sql, tx)) { c.Parameters.AddWithValue("$c", cutoffMs); return c.ExecuteNonQuery(); } };
                p.Rows = run("DELETE FROM recipient WHERE message_key IN (SELECT message_key FROM message WHERE received_ms < $c);");
                p.Messages = run("DELETE FROM message WHERE received_ms < $c;");
                p.Intervals = run("DELETE FROM collection_interval WHERE end_ms <= $c;");
                run("UPDATE collection_interval SET start_ms = $c WHERE start_ms < $c AND end_ms > $c;");
                run("UPDATE collection_interval SET covered_from_ms = $c WHERE covered_from_ms < $c AND end_ms > $c;");
                tx.Commit();
            }
            if (p.Messages + p.Rows > 0) Exec("PRAGMA incremental_vacuum;");
            return p;
        }

        // ---- Statistics -----------------------------------------------------

        public StoreStatistics GetStatistics(int limit, int maxAttempts)
        {
            var s = new StoreStatistics();
            s.FileBytes = FileSize(DatabasePath) + FileSize(DatabasePath + "-wal");
            s.SchemaVersion = (int)Scalar("PRAGMA user_version;");
            s.CreatedByVersion = GetMetadata("created_by_version");
            int cl;
            if (int.TryParse(GetMetadata("compaction_limit"), NumberStyles.Integer, CultureInfo.InvariantCulture, out cl)) s.CompactionLimit = cl;
            s.Runs = Scalar("SELECT COUNT(*) FROM run;");
            using (SqliteCommand c = Command(@"SELECT COUNT(*), COALESCE(SUM(" + CountSql + @" > $l), 0),
    COALESCE(SUM(" + CountPendingSql + @"), 0),
    COALESCE(SUM(m.recipient_rows > $l AND m.list_rows > 0 AND m.count_value IS NULL AND (m.count_state = 'NoCount' OR m.count_attempts >= $a)), 0),
    COALESCE(SUM(" + ListPendingSql + @"), 0)
FROM message m;"))
            {
                c.Parameters.AddWithValue("$l", limit); c.Parameters.AddWithValue("$a", maxAttempts); c.Parameters.AddWithValue("$r", MaxRoutesPerMessage);
                using (SqliteDataReader r = c.ExecuteReader())
                    if (r.Read()) { s.Messages = r.GetInt64(0); s.Matches = r.GetInt64(1); s.Pending = r.GetInt64(2); s.Unmeasurable = r.GetInt64(3); s.ListsPending = r.GetInt64(4); }
            }
            using (SqliteCommand c = Command("SELECT COUNT(*), MIN(covered_from_ms), MAX(end_ms), MAX(completed_ms) FROM collection_interval WHERE covered_from_ms < end_ms;"))
            using (SqliteDataReader r = c.ExecuteReader())
                if (r.Read())
                {
                    s.Intervals = r.GetInt64(0);
                    s.FirstCoveredMs = NullableLong(r, 1);
                    s.LastCoveredMs = NullableLong(r, 2);
                    s.LastSuccessfulCollectionMs = NullableLong(r, 3);
                }
            return s;
        }

        /// <summary>Recipient rows stored (reads the whole table: diagnostics and tests only).</summary>
        public long CountRecipientRows() { return Scalar("SELECT COUNT(*) FROM recipient;"); }

        public List<DayStatistics> GetDailyStatistics(string[] signatures, string[] domains, DateTime firstLocalDate, int days, TimeZoneInfo zone, long settlingMs, int limit, int maxAttempts)
        {
            List<TimeRange> covered = GetCoverage(signatures, 0);
            List<TimeRange> final = GetCoverage(signatures, settlingMs);
            var list = new List<DayStatistics>();
            string sql = @"SELECT COUNT(DISTINCT CASE WHEN " + CountSql + @" > $l THEN COALESCE(NULLIF(m.message_id, ''), hex(m.trace_id)) END),
    COALESCE(SUM(" + CountPendingSql + " OR " + ListPendingSql + @"), 0)
FROM message m" + SenderJoin(domains) + " WHERE m.received_ms >= $s AND m.received_ms < $e" + DomainClause(domains) + ";";
            using (SqliteCommand c = Command(sql))
            {
                SqliteParameter ps = c.Parameters.Add("$s", SqliteType.Integer), pe = c.Parameters.Add("$e", SqliteType.Integer);
                c.Parameters.AddWithValue("$l", limit); c.Parameters.AddWithValue("$a", maxAttempts); c.Parameters.AddWithValue("$r", MaxRoutesPerMessage);
                AddDomains(c, domains);
                for (int i = 0; i < days; i++)
                {
                    DateTime day = firstLocalDate.Date.AddDays(i);
                    long start = Coverage.LocalToUnixMs(day, zone), end = Coverage.LocalToUnixMs(day.AddDays(1), zone);
                    ps.Value = start; pe.Value = end;
                    var d = new DayStatistics { LocalDate = day, StartMs = start, EndMs = end };
                    using (SqliteDataReader r = c.ExecuteReader())
                        if (r.Read()) { d.Matches = r.GetInt64(0); d.Pending = r.GetInt64(1); }
                    d.CoveredMs = (end - start) - Coverage.Gaps(covered, start, end).Sum(g => g.Length);
                    d.FinalCoveredMs = (end - start) - Coverage.Gaps(final, start, end).Sum(g => g.Length);
                    list.Add(d);
                }
            }
            return list;
        }

        // ---- Report ---------------------------------------------------------

        /// <summary>
        /// Messages of the period over the limit, one row per Message ID (a message seen under two trace
        /// IDs, for example on its way back from an on-premises server, is reported once: its first trace).
        /// </summary>
        static string ReportSql(string[] domains, string columns)
        {
            return @"WITH c AS (
    SELECT m.message_key, COALESCE(NULLIF(m.message_id, ''), lower(hex(m.trace_id))) AS mid, m.received_ms, m.sender_id, m.subject,
           m.list_rows, m.recipient_rows, " + CountSql + @" AS cnt,
           CASE WHEN m.list_rows = 0 THEN 'Trace' ELSE m.count_source END AS src, m.list_state
    FROM message m" + SenderJoin(domains) + @"
    WHERE m.received_ms >= $s AND m.received_ms < $e AND m.recipient_rows > $l" + DomainClause(domains) + @"
), r AS (
    SELECT c.*, ROW_NUMBER() OVER (PARTITION BY c.mid ORDER BY c.received_ms, c.message_key) AS rn FROM c WHERE c.cnt > $l
)
SELECT " + columns + @" FROM r LEFT JOIN address a ON a.address_id = r.sender_id WHERE r.rn = 1 ORDER BY r.received_ms, r.message_key;";
        }

        void AddReportParameters(SqliteCommand c, ReportRequest q)
        {
            c.Parameters.AddWithValue("$s", q.StartMs);
            c.Parameters.AddWithValue("$e", q.EndMs);
            c.Parameters.AddWithValue("$l", q.Limit);
            AddDomains(c, q.SenderDomains);
        }

        /// <summary>
        /// Builds and writes the report inside one read transaction, so the row plan (first pass)
        /// and the written rows (second pass) see exactly the same data.
        /// </summary>
        public ReportResult WriteReport(ReportRequest q)
        {
            using (SqliteTransaction tx = _db.BeginTransaction(true))
            {
                var firstTimes = new List<long>();
                using (SqliteCommand c = Command(ReportSql(q.SenderDomains, "r.received_ms"), tx))
                {
                    AddReportParameters(c, q);
                    using (SqliteDataReader r = c.ExecuteReader())
                        while (r.Read()) firstTimes.Add(r.GetInt64(0));
                }
                List<ReportPart> plan = ReportPlanner.Plan(firstTimes, q);
                using (SqliteCommand c = Command(ReportSql(q.SenderDomains, "r.message_key, r.mid, r.received_ms, r.cnt, r.src, r.list_rows, r.recipient_rows, a.address, r.subject, r.list_state"), tx))
                using (SqliteCommand people = Command(@"SELECT a.address, s.name = 'expanded', x.direct FROM recipient x JOIN address a ON a.address_id = x.address_id JOIN status s ON s.status_id = x.status_id
WHERE x.message_key = $m ORDER BY s.name = 'expanded' DESC, a.address COLLATE NOCASE;", tx))
                {
                    AddReportParameters(c, q);
                    SqliteParameter pm = people.Parameters.Add("$m", SqliteType.Integer);
                    using (SqliteDataReader r = c.ExecuteReader())
                    {
                        IEnumerable<ReportRow> rows = ReadRows(r, q.IncludeRecipientDetails ? people : null, pm, q.MaxRecipientsListed);
                        ReportResult result = ReportWriter.Write(rows, plan, q);
                        tx.Commit();
                        return result;
                    }
                }
            }
        }

        /// <summary>
        /// Report rows with their recipient list: as the sender addressed it (no distribution list, or
        /// list rebuilt from the routes), otherwise every address of the trace after expansion.
        /// </summary>
        static IEnumerable<ReportRow> ReadRows(SqliteDataReader r, SqliteCommand people, SqliteParameter key, int maxListed)
        {
            while (r.Read())
            {
                var row = new ReportRow
                {
                    MessageId = r.GetString(1),
                    FirstMs = r.GetInt64(2),
                    RecipientCount = r.IsDBNull(3) ? (long?)null : r.GetInt64(3),
                    CountSource = r.IsDBNull(4) ? "" : r.GetString(4),
                    Lists = (int)r.GetInt64(5),
                    TraceRecipients = r.GetInt64(6),
                    Sender = r.IsDBNull(7) ? "" : r.GetString(7),
                    Subject = r.IsDBNull(8) ? "" : r.GetString(8)
                };
                row.AfterExpansion = row.Lists > 0 && (r.IsDBNull(9) || r.GetString(9) != "Done");
                if (people != null)
                {
                    key.Value = r.GetInt64(0);
                    using (SqliteDataReader p = people.ExecuteReader())
                        while (p.Read())
                        {
                            if (row.Lists > 0 && !row.AfterExpansion && (p.IsDBNull(2) || p.GetInt64(2) != 1)) continue;
                            if (maxListed > 0 && row.Recipients.Count >= maxListed) { row.NotListed++; continue; }
                            if (p.GetInt64(1) == 1 && row.Recipients.Count == row.ListsShown) row.ListsShown++;
                            row.Recipients.Add(p.GetString(0));
                        }
                }
                yield return row;
            }
        }
    }

    // -------------------------------------------------------------------------
    // 4. Report planning and writing
    // -------------------------------------------------------------------------

    public sealed class ReportRow
    {
        public string MessageId, Sender, Subject, CountSource;
        public long FirstMs, TraceRecipients;
        public long? RecipientCount;
        public int Lists;                                    // distribution lists expanded (rows 'expanded' of the trace)
        public List<string> Recipients = new List<string>(); // recipient addresses, distribution lists first
        public int ListsShown;                               // distribution lists at the start of Recipients
        public bool AfterExpansion;                          // Recipients = every address of the trace after expansion (list before expansion not known)
        public long NotListed;                               // addresses beyond MaxRecipientsListed
    }

    public sealed class ReportRequest
    {
        public long StartMs, EndMs;                 // report period [start, end)
        public TimeZoneInfo Zone;
        public string TimeZoneLabel;                // shown in column headers, e.g. "Europe/Paris"
        public int Limit = 25;                      // messages with MORE recipients than this are reported
        public string[] SenderDomains = new string[0];
        public string SplitBy = "Week";             // Rows | Day | Week
        public int MaxRowsPerFile = 500000;
        public bool IncludeRecipientDetails = true;
        public int MaxRecipientsListed = 500;
        public bool WriteCsv = true, WriteHtml = true;
        public string CsvDelimiter = ";";
        public string OutputDirectory;
        public string FilePrefix = "RecipientLimit";
        public string HtmlTemplatePath;
        public string Title = "Messages with more than 25 recipients";
        public string ScopeLabel, ToolVersion, RangeName, CoverageNote;
        public double CoveragePercent = 100;
        public long Unmeasured;                     // messages with distribution lists whose count is not known yet
        public int HtmlChunkRows = 20000;
    }

    public sealed class ReportPart
    {
        public int Index;                 // 1-based position in the whole report
        public long FirstRow, RowCount;   // position in the ordered message list
        public long StartMs, EndMs;       // period shown for this file
        public long GroupStartMs, GroupEndMs;   // day or week (or whole period) the file belongs to; used in the file name
        public string GroupLabel;         // e.g. "2026-09-01 to 2026-09-07"
        public int PartInGroup = 1, PartsInGroup = 1;
        public string BaseName;           // file name without extension
    }

    public sealed class ReportFile
    {
        public string Path, Kind;
        public long Rows, Bytes;
        public int PartIndex;
        public string Label;
    }

    public sealed class ReportResult
    {
        public long Messages, Senders, RecipientTotal, WithLists, MessagesWithoutCount, ListedAfterExpansion;
        public List<ReportFile> Files = new List<ReportFile>();
        public List<ReportPart> Parts = new List<ReportPart>();
    }

    public static class ReportPlanner
    {
        /// <summary>
        /// Decides how many files to write. There is no split while the report fits in
        /// MaxRowsPerFile rows. Above that, rows are grouped by local day or local week
        /// (Monday to Sunday), or simply cut every MaxRowsPerFile rows. A day or week that
        /// is still too large is cut into balanced numbered parts.
        /// </summary>
        public static List<ReportPart> Plan(IList<long> orderedFirstMs, ReportRequest q)
        {
            var parts = new List<ReportPart>();
            long total = orderedFirstMs.Count;
            int max = Math.Max(1, q.MaxRowsPerFile);
            if (total <= max)
            {
                parts.Add(new ReportPart { FirstRow = 0, RowCount = total, StartMs = q.StartMs, EndMs = q.EndMs, GroupStartMs = q.StartMs, GroupEndMs = q.EndMs, GroupLabel = PeriodLabel(q.StartMs, q.EndMs, q.Zone) });
            }
            else if (string.Equals(q.SplitBy, "Rows", StringComparison.OrdinalIgnoreCase))
            {
                int count = (int)((total + max - 1) / max);
                for (int i = 0; i < count; i++)
                {
                    long first = (long)i * max, rows = Math.Min(max, total - first);
                    long start = i == 0 ? q.StartMs : orderedFirstMs[(int)first];
                    long end = i == count - 1 ? q.EndMs : orderedFirstMs[(int)(first + rows)];
                    parts.Add(new ReportPart { FirstRow = first, RowCount = rows, StartMs = start, EndMs = end, GroupStartMs = q.StartMs, GroupEndMs = q.EndMs, GroupLabel = PeriodLabel(start, end, q.Zone), PartInGroup = i + 1, PartsInGroup = count });
                }
            }
            else
            {
                bool week = string.Equals(q.SplitBy, "Week", StringComparison.OrdinalIgnoreCase);
                long row = 0;
                while (row < total)
                {
                    DateTime groupStartLocal = GroupStart(orderedFirstMs[(int)row], q.Zone, week);
                    long groupEndMs = Coverage.LocalToUnixMs(groupStartLocal.AddDays(week ? 7 : 1), q.Zone);
                    long groupStartMs = Math.Max(q.StartMs, Coverage.LocalToUnixMs(groupStartLocal, q.Zone));
                    long clippedEnd = Math.Min(q.EndMs, groupEndMs);
                    long first = row;
                    while (row < total && orderedFirstMs[(int)row] < groupEndMs) row++;
                    long rows = row - first;
                    int pieces = (int)((rows + max - 1) / max);
                    long size = (rows + pieces - 1) / pieces;
                    string label = PeriodLabel(groupStartMs, clippedEnd, q.Zone);
                    for (int i = 0; i < pieces; i++)
                    {
                        long pf = first + i * size, pr = Math.Min(size, rows - i * size);
                        parts.Add(new ReportPart
                        {
                            FirstRow = pf, RowCount = pr,
                            StartMs = i == 0 ? groupStartMs : orderedFirstMs[(int)pf],
                            EndMs = i == pieces - 1 ? clippedEnd : orderedFirstMs[(int)(pf + pr)],
                            GroupStartMs = groupStartMs, GroupEndMs = clippedEnd,
                            GroupLabel = label, PartInGroup = i + 1, PartsInGroup = pieces
                        });
                    }
                }
            }
            for (int i = 0; i < parts.Count; i++)
            {
                ReportPart p = parts[i];
                p.Index = i + 1;
                string name = q.FilePrefix + "_" + FileLabel(p.GroupStartMs, p.GroupEndMs, q.Zone);
                if (parts.Count > 1 && string.Equals(q.SplitBy, "Rows", StringComparison.OrdinalIgnoreCase))
                    name = q.FilePrefix + "_" + FileLabel(q.StartMs, q.EndMs, q.Zone) + "_part" + p.PartInGroup.ToString("00") + "of" + p.PartsInGroup.ToString("00");
                else if (p.PartsInGroup > 1)
                    name += "_part" + p.PartInGroup + "of" + p.PartsInGroup;
                p.BaseName = name;
            }
            return parts;
        }

        static DateTime GroupStart(long unixMs, TimeZoneInfo zone, bool week)
        {
            DateTime local = TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).DateTime.Date;
            if (!week) return local;
            int shift = ((int)local.DayOfWeek + 6) % 7;   // Monday = 0
            return local.AddDays(-shift);
        }

        static DateTime Local(long unixMs, TimeZoneInfo zone)
        {
            return TimeZoneInfo.ConvertTime(DateTimeOffset.FromUnixTimeMilliseconds(unixMs), zone).DateTime;
        }

        /// <summary>Whole days are shown with an inclusive last day; other bounds show times.</summary>
        public static string PeriodLabel(long startMs, long endMs, TimeZoneInfo zone)
        {
            DateTime s = Local(startMs, zone), e = Local(endMs, zone);
            if (s.TimeOfDay == TimeSpan.Zero && e.TimeOfDay == TimeSpan.Zero)
            {
                DateTime last = e.AddDays(-1);
                return last <= s ? s.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                                 : s.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + " to " + last.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            return s.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + " to " + e.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        public static string FileLabel(long startMs, long endMs, TimeZoneInfo zone)
        {
            return PeriodLabel(startMs, endMs, zone).Replace(" to ", "_to_").Replace(" ", "_").Replace(":", "");
        }
    }

    public static class ReportWriter
    {
        static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

        public static ReportResult Write(IEnumerable<ReportRow> rows, List<ReportPart> plan, ReportRequest q)
        {
            var result = new ReportResult { Parts = plan };
            Directory.CreateDirectory(q.OutputDirectory);
            string[] template = null;
            if (q.WriteHtml)
            {
                string text = File.ReadAllText(q.HtmlTemplatePath, Encoding.UTF8);
                int a = text.IndexOf("%%CHUNKS%%", StringComparison.Ordinal), b = text.IndexOf("%%META%%", StringComparison.Ordinal);
                if (a < 0 || b < a || a != text.LastIndexOf("%%CHUNKS%%", StringComparison.Ordinal) || b != text.LastIndexOf("%%META%%", StringComparison.Ordinal))
                    throw new InvalidDataException("The HTML template must contain %%CHUNKS%% then %%META%%, exactly once each: " + q.HtmlTemplatePath);
                template = new[] { text.Substring(0, a), text.Substring(a + 10, b - a - 10), text.Substring(b + 8) };
            }
            var allSenders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (IEnumerator<ReportRow> e = rows.GetEnumerator())
            {
                foreach (ReportPart part in plan)
                {
                    CsvOut csv = q.WriteCsv ? new CsvOut(Path.Combine(q.OutputDirectory, part.BaseName + ".csv"), q) : null;
                    HtmlOut html = q.WriteHtml ? new HtmlOut(Path.Combine(q.OutputDirectory, part.BaseName + ".html"), q, template) : null;
                    try
                    {
                        for (long i = 0; i < part.RowCount; i++)
                        {
                            if (!e.MoveNext()) throw new InvalidOperationException("The database returned fewer rows than planned; the report is not written.");
                            ReportRow row = e.Current;
                            result.Messages++;
                            if (row.Lists > 0) result.WithLists++;
                            if (row.AfterExpansion) result.ListedAfterExpansion++;
                            if (row.RecipientCount.HasValue) result.RecipientTotal += row.RecipientCount.Value; else result.MessagesWithoutCount++;
                            allSenders.Add(row.Sender ?? "");
                            if (csv != null) csv.Add(row);
                            if (html != null) html.Add(row);
                        }
                        if (csv != null) { csv.Close(); result.Files.Add(new ReportFile { Path = csv.FilePath, Kind = "CSV", Rows = part.RowCount, Bytes = new FileInfo(csv.FilePath).Length, PartIndex = part.Index, Label = part.GroupLabel }); }
                        if (html != null)
                        {
                            html.Close(part, plan);
                            result.Files.Add(new ReportFile { Path = html.FilePath, Kind = "HTML", Rows = part.RowCount, Bytes = new FileInfo(html.FilePath).Length, PartIndex = part.Index, Label = part.GroupLabel });
                        }
                    }
                    finally
                    {
                        if (csv != null) csv.Dispose();
                        if (html != null) html.Dispose();
                    }
                }
                if (e.MoveNext()) throw new InvalidOperationException("The database returned more rows than planned; the report is not written.");
            }
            result.Senders = allSenders.Count;
            return result;
        }

        /// <summary>
        /// Recipient list of a CSV cell: "; " separated, with a note when the list was cut or when it is
        /// the list after expansion (list before expansion not known).
        /// </summary>
        public static string RecipientText(ReportRow r)
        {
            string list = string.Join("; ", r.Recipients);
            if (r.NotListed > 0) list += "; (+" + r.NotListed.ToString(CultureInfo.InvariantCulture) + " more)";
            return r.AfterExpansion ? list + "; " + AfterExpansionNote : list;
        }

        public const string AfterExpansionNote = "(addresses after distribution list expansion)";

        // ---- CSV ------------------------------------------------------------

        sealed class CsvOut : IDisposable
        {
            readonly StreamWriter _w;
            readonly ReportRequest _q;
            public string FilePath;
            public CsvOut(string path, ReportRequest q)
            {
                FilePath = path;
                _q = q;
                _w = new StreamWriter(new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 1 << 16), new UTF8Encoding(true));
                var header = new List<string> { "Received time (" + q.TimeZoneLabel + ")", "Sender" };
                if (q.IncludeRecipientDetails) header.Add("Recipients");
                header.AddRange(new[] { "Subject", "Recipient count", "Distribution lists", "Message ID" });
                _w.WriteLine(string.Join(q.CsvDelimiter, header.Select(h => Text.SafeCsv(h, q.CsvDelimiter))));
            }
            public void Add(ReportRow r)
            {
                string d = _q.CsvDelimiter;
                _w.Write(Time.FormatLocal(r.FirstMs, _q.Zone)); _w.Write(d);
                _w.Write(Text.SafeCsv(r.Sender, d)); _w.Write(d);
                if (_q.IncludeRecipientDetails) { _w.Write(Text.SafeCsv(RecipientText(r), d)); _w.Write(d); }
                _w.Write(Text.SafeCsv(r.Subject, d)); _w.Write(d);
                if (r.RecipientCount.HasValue) _w.Write(r.RecipientCount.Value.ToString(CultureInfo.InvariantCulture));
                _w.Write(d);
                _w.Write(r.Lists.ToString(CultureInfo.InvariantCulture)); _w.Write(d);
                _w.WriteLine(Text.SafeCsv(r.MessageId, d));
            }
            public void Close() { _w.Flush(); _w.Dispose(); }
            public void Dispose() { _w.Dispose(); }
        }

        // ---- HTML -----------------------------------------------------------
        // The HTML file is self-contained. Rows are written in compressed chunks
        // (JSON -> gzip -> base64) inside <script> blocks; the page decompresses them
        // in the browser. Senders, subjects and recipients are stored once in
        // dictionaries and referenced by number, which keeps large files small.

        sealed class HtmlOut : IDisposable
        {
            readonly StreamWriter _w;
            readonly ReportRequest _q;
            readonly string[] _template;
            public string FilePath;
            readonly Dictionary<string, int> _senders = new Dictionary<string, int>(StringComparer.Ordinal);
            readonly Dictionary<string, int> _subjects = new Dictionary<string, int>(StringComparer.Ordinal);
            readonly Dictionary<string, int> _people = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            readonly List<string> _newSenders = new List<string>(), _newSubjects = new List<string>(), _newPeople = new List<string>();
            readonly List<long> _t = new List<long>(), _n = new List<long>(), _x = new List<long>(), _rx = new List<long>();
            readonly List<int> _s = new List<int>(), _j = new List<int>(), _l = new List<int>(), _c = new List<int>(), _rl = new List<int>(), _ri = new List<int>(), _rd = new List<int>(), _ra = new List<int>();
            readonly List<string> _m = new List<string>();
            long _rows, _chunks, _recipientTotal, _withLists, _afterExpansion;

            public HtmlOut(string path, ReportRequest q, string[] template)
            {
                FilePath = path;
                _q = q;
                _template = template;
                _w = new StreamWriter(new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 1 << 16), new UTF8Encoding(false));
                _w.Write(template[0]);
            }

            static int Index(Dictionary<string, int> map, List<string> added, string value)
            {
                value = value ?? "";
                int i;
                if (!map.TryGetValue(value, out i)) { i = map.Count; map.Add(value, i); added.Add(value); }
                return i;
            }

            public void Add(ReportRow r)
            {
                _rows++;
                // Local wall-clock seconds: the page formats them as UTC, so the browser's own
                // time zone never changes what is displayed.
                _t.Add((r.FirstMs + (long)_q.Zone.GetUtcOffset(DateTimeOffset.FromUnixTimeMilliseconds(r.FirstMs)).TotalMilliseconds) / 1000);
                _s.Add(Index(_senders, _newSenders, r.Sender));
                _j.Add(Index(_subjects, _newSubjects, r.Subject));
                _n.Add(r.RecipientCount.HasValue ? r.RecipientCount.Value : -1);
                if (r.RecipientCount.HasValue) _recipientTotal += r.RecipientCount.Value;
                _l.Add(r.Lists);
                if (r.Lists > 0) _withLists++;
                _x.Add(r.TraceRecipients);
                _c.Add(r.CountSource == "Submit" ? 1 : r.CountSource == "Receive" ? 2 : 0);
                _m.Add(r.MessageId);
                if (_q.IncludeRecipientDetails)
                {
                    _rl.Add(r.Recipients.Count);
                    _rx.Add(r.NotListed);
                    _rd.Add(r.ListsShown);
                    _ra.Add(r.AfterExpansion ? 1 : 0);
                    if (r.AfterExpansion) _afterExpansion++;
                    foreach (string p in r.Recipients) _ri.Add(Index(_people, _newPeople, p));
                }
                if (_t.Count >= _q.HtmlChunkRows) Flush();
            }

            void Flush()
            {
                if (_t.Count == 0) return;
                var chunk = new Dictionary<string, object>
                {
                    { "ds", _newSenders }, { "dj", _newSubjects }, { "t", _t }, { "s", _s }, { "j", _j }, { "n", _n }, { "m", _m },
                    { "l", _l }, { "x", _x }, { "c", _c }
                };
                if (_q.IncludeRecipientDetails) { chunk["dr"] = _newPeople; chunk["rl"] = _rl; chunk["ri"] = _ri; chunk["rx"] = _rx; chunk["rd"] = _rd; chunk["ra"] = _ra; }
                byte[] json = JsonSerializer.SerializeToUtf8Bytes(chunk, JsonOptions);
                using (var buffer = new MemoryStream())
                {
                    using (var gzip = new GZipStream(buffer, CompressionLevel.Optimal, true)) gzip.Write(json, 0, json.Length);
                    _w.Write("<script type=\"application/x-rlr-chunk\">");
                    _w.Write(Convert.ToBase64String(buffer.GetBuffer(), 0, (int)buffer.Length));
                    _w.Write("</script>\n");
                }
                _chunks++;
                foreach (var list in new List<string>[] { _newSenders, _newSubjects, _newPeople, _m }) list.Clear();
                _t.Clear(); _n.Clear(); _x.Clear(); _rx.Clear(); _s.Clear(); _j.Clear(); _l.Clear(); _c.Clear(); _rl.Clear(); _ri.Clear(); _rd.Clear(); _ra.Clear();
            }

            public void Close(ReportPart part, List<ReportPart> plan)
            {
                Flush();
                var meta = new Dictionary<string, object>
                {
                    { "title", _q.Title },
                    { "limit", _q.Limit },
                    { "scope", _q.ScopeLabel },
                    { "rangeName", _q.RangeName },
                    { "timeZone", _q.TimeZoneLabel },
                    { "periodStart", Time.FormatLocal(_q.StartMs, _q.Zone) },
                    { "periodEnd", Time.FormatLocal(_q.EndMs, _q.Zone) },
                    { "fileStart", Time.FormatLocal(part.StartMs, _q.Zone) },
                    { "fileEnd", Time.FormatLocal(part.EndMs, _q.Zone) },
                    { "fileLabel", part.GroupLabel },
                    { "partIndex", part.Index }, { "partCount", plan.Count },
                    { "partInGroup", part.PartInGroup }, { "partsInGroup", part.PartsInGroup },
                    { "rows", _rows }, { "chunks", _chunks },
                    { "recipientTotal", _recipientTotal }, { "withLists", _withLists },
                    { "includeRecipients", _q.IncludeRecipientDetails },
                    { "afterExpansion", _afterExpansion },
                    { "coveragePercent", Math.Round(_q.CoveragePercent, 2) },
                    { "coverageNote", _q.CoverageNote },
                    { "unmeasured", _q.Unmeasured },
                    { "generated", DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
                    { "toolVersion", _q.ToolVersion },
                    { "siblings", plan.Select(p => new Dictionary<string, object> {
                        { "file", p.BaseName + ".html" }, { "label", p.GroupLabel + (p.PartsInGroup > 1 ? " (part " + p.PartInGroup + " of " + p.PartsInGroup + ")" : "") },
                        { "rows", p.RowCount }, { "index", p.Index } }).ToList() }
                };
                string json = JsonSerializer.Serialize(meta, JsonOptions).Replace("</", "<\\/");
                _w.Write(_template[1]);
                _w.Write(json);
                _w.Write(_template[2]);
                _w.Flush();
                _w.Dispose();
            }

            public void Dispose() { _w.Dispose(); }
        }
    }
}
