// =============================================================================
//  Recipient Limit Report - tests: in-memory Microsoft Graph message trace API
// -----------------------------------------------------------------------------
//  Author  : Nicolas Fabert
//  Version : 1.0.1
//
//  An HttpMessageHandler given to the engine (CollectorOptions.Handler). It answers like the real
//  API measured in the lab (2026-10-02 and 2026-10-05):
//    - $filter keeps ONE value per property (the last one); '*@domain' matches a sender domain;
//    - newest first; $top 1-5000; @odata.nextLink with $skiptoken;
//    - receivedDateTime: both bounds required ('ge' and 'le'), at most 10 days;
//    - every row of a message has the same receivedDateTime;
//    - a distribution list is one row with the status 'expanded', each member one more row;
//    - getDetailsByRecipient: Receive (no RcptCount), Submit (RcptCount = recipients before
//      expansion), Expand DL (RcptCount = members of the list), DLP rule (RcptCount = recipients
//      after expansion), Deliver / Fail. A message received by SMTP has no Submit event: its
//      Receive event carries the RcptCount of the SMTP envelope. The route of a member added
//      by the expansion of a list has neither Receive nor Submit (it starts after the expansion).
//  Failures can be injected: 429 (Retry-After), 401, 500, 403, missing service principal.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace RlrTests
{
    public sealed class FakeRow
    {
        public string Id, MessageId, Sender, Recipient, Subject, Status = "delivered";
        public bool Direct = true;      // addressed by the sender (false: member added by the expansion of a list)
        public DateTimeOffset Received;
    }

    public sealed class FakeMessage
    {
        public string Id, MessageId;
        public int Envelope;            // recipients before expansion
        public string Submission = "Mailbox";   // Mailbox (Submit event) | Smtp (Receive with RcptCount) | None (no count)
    }

    public sealed class FakeGraph : HttpMessageHandler
    {
        public readonly List<FakeRow> Rows = new List<FakeRow>();
        public readonly Dictionary<string, FakeMessage> Messages = new Dictionary<string, FakeMessage>();
        public readonly List<string> Requests = new List<string>();
        public int ThrottleNext, UnauthorizedNext, ServerErrorNext, RetryAfterSeconds = 1, DelayMs;
        public bool Forbidden, MissingServicePrincipal;
        public int Served;

        /// <summary>
        /// Adds one message. 'direct' are the recipients addressed one by one; each list is addressed
        /// once and expanded to its members (rows with the status 'expanded' + one row per member).
        /// Returns the trace ID.
        /// </summary>
        public string AddMessage(string sender, string[] direct, DateTimeOffset received, string subject, Dictionary<string, string[]> lists = null, string status = "delivered", string submission = "Mailbox", string messageId = null)
        {
            string id = Guid.NewGuid().ToString();
            messageId = messageId ?? "<" + Guid.NewGuid().ToString("N") + "@contoso.com>";
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Action<string, string, bool> add = (rcpt, st, d) => { if (seen.Add(rcpt)) Rows.Add(new FakeRow { Id = id, MessageId = messageId, Sender = sender, Recipient = rcpt, Subject = subject, Received = received, Status = st, Direct = d }); };
            if (lists != null) foreach (var l in lists) { add(l.Key, "expanded", true); }
            foreach (string r in direct) add(r, status, true);
            if (lists != null) foreach (var l in lists) foreach (string m in l.Value) add(m, status, false);
            Messages[id] = new FakeMessage { Id = id, MessageId = messageId, Envelope = direct.Length + (lists == null ? 0 : lists.Count), Submission = submission };
            return id;
        }

        static HttpResponseMessage Json(HttpStatusCode code, string json)
        {
            return new HttpResponseMessage(code) { Content = new StringContent(json, Encoding.UTF8, "application/json") };
        }

        static HttpResponseMessage Error(HttpStatusCode code, string message)
        {
            return Json(code, JsonSerializer.Serialize(new { error = new { code = code.ToString(), message = message } }));
        }

        static string Mep(string name, string type, string value) { return "<root><MEP Name=\"" + name + "\" " + type + "=\"" + value + "\" /><MEP Name=\"SequenceNumber\" Long=\"0\" /></root>"; }

        List<object> Route(FakeRow t)
        {
            FakeMessage m = Messages[t.Id];
            int after = Rows.Count(r => r.Id == t.Id);
            var list = new List<object>();
            Action<int, string, string> add = (seconds, ev, data) =>
                list.Add(new { id = t.Id, messageId = t.MessageId, dateTime = t.Received.AddSeconds(seconds).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture), @event = ev, action = "", description = ev, data = data });
            if (t.Direct)
            {
                add(0, "Receive", m.Submission == "Smtp" ? Mep("RcptCount", "Integer", m.Envelope.ToString(CultureInfo.InvariantCulture)) : "<root><MEP Name=\"ServerHostName\" String=\"EXCH01\" /></root>");
                if (m.Submission == "Mailbox") add(1, "Submit", Mep("RcptCount", "Integer", m.Envelope.ToString(CultureInfo.InvariantCulture)));
            }
            if (t.Status == "expanded") { add(1, "Expand DL", Mep("RcptCount", "Integer", "4")); add(2, "Drop", "<root />"); }
            add(2, "DLP rule", Mep("RcptCount", "Integer", after.ToString(CultureInfo.InvariantCulture)));
            if (t.Status == "delivered") add(3, "Deliver", Mep("RcptCount", "Integer", "1"));
            if (t.Status == "failed") add(3, "Fail", Mep("RcptCount", "Integer", "1"));
            return list;
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            string url = Uri.UnescapeDataString(request.RequestUri.PathAndQuery);
            lock (Requests) Requests.Add(url);
            if (DelayMs > 0) await Task.Delay(DelayMs, ct);
            if (Forbidden) return Error(HttpStatusCode.Forbidden, "Authorization_RequestDenied");
            if (MissingServicePrincipal) return Error(HttpStatusCode.Unauthorized, "Service principal-less authentication failed: The service principal for App ID 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373 was not found.");
            if (Interlocked.Decrement(ref ThrottleNext) >= 0)
            {
                var r = Error((HttpStatusCode)429, "Your recent queries have surpassed the permitted limit, please try again later.");
                r.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(RetryAfterSeconds));
                return r;
            }
            if (Interlocked.Decrement(ref UnauthorizedNext) >= 0) return Error(HttpStatusCode.Unauthorized, "InvalidAuthenticationToken: token expired");
            if (Interlocked.Decrement(ref ServerErrorNext) >= 0) return Error(HttpStatusCode.InternalServerError, "Transient failure");

            Match detail = Regex.Match(url, @"/messageTraces/([^/]+)/getDetailsByRecipient\(recipientAddress='((?:[^']|'')*)'\)");
            if (detail.Success)
            {
                string id = detail.Groups[1].Value, rcpt = detail.Groups[2].Value.Replace("''", "'");
                FakeRow t = Rows.FirstOrDefault(x => x.Id == id && string.Equals(x.Recipient, rcpt, StringComparison.OrdinalIgnoreCase));
                if (t == null) return Json(HttpStatusCode.OK, "{\"value\":[]}");
                return Json(HttpStatusCode.OK, JsonSerializer.Serialize(new { value = Route(t) }));
            }

            string query = request.RequestUri.Query.TrimStart('?');
            var args = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string part in query.Split('&')) { int eq = part.IndexOf('='); if (eq > 0) args[Uri.UnescapeDataString(part.Substring(0, eq))] = Uri.UnescapeDataString(part.Substring(eq + 1)); }
            string filter = args.ContainsKey("$filter") ? args["$filter"] : "";
            int top = args.ContainsKey("$top") ? int.Parse(args["$top"], CultureInfo.InvariantCulture) : 1000;
            if (top < 1 || top > 5000) return Error(HttpStatusCode.BadRequest, "ResultSize is invalid");
            int skip = args.ContainsKey("$skiptoken") ? int.Parse(args["$skiptoken"], CultureInfo.InvariantCulture) : 0;
            Match ge = Regex.Match(filter, @"receivedDateTime ge (\S+)"), le = Regex.Match(filter, @"receivedDateTime le (\S+)");
            if (!ge.Success || !le.Success) return Error(HttpStatusCode.BadRequest, "EndDate is required when a StartDate is entered.");
            DateTimeOffset start = DateTimeOffset.Parse(ge.Groups[1].Value, CultureInfo.InvariantCulture), end = DateTimeOffset.Parse(le.Groups[1].Value, CultureInfo.InvariantCulture);
            if (end - start > TimeSpan.FromDays(10).Add(TimeSpan.FromSeconds(1))) return Error(HttpStatusCode.BadRequest, "The interval between StartDate and EndDate can't be longer than 10 days.");
            Match sender = Regex.Match(filter, @"senderAddress eq '((?:[^']|'')*)'");
            IEnumerable<FakeRow> rows = Rows.Where(r => r.Received >= start && r.Received <= end);
            if (sender.Success)
            {
                string s = sender.Groups[1].Value.Replace("''", "'");
                rows = s.StartsWith("*@") ? rows.Where(r => r.Sender.EndsWith(s.Substring(1), StringComparison.OrdinalIgnoreCase)) : rows.Where(r => string.Equals(r.Sender, s, StringComparison.OrdinalIgnoreCase));
            }
            List<FakeRow> all = rows.OrderByDescending(r => r.Received).ThenBy(r => r.Id, StringComparer.Ordinal).ThenBy(r => r.Recipient, StringComparer.Ordinal).ToList();
            List<FakeRow> page = all.Skip(skip).Take(top).ToList();
            Interlocked.Increment(ref Served);
            var body = new Dictionary<string, object>
            {
                { "value", page.Select(r => new Dictionary<string, object> {
                    { "id", r.Id }, { "messageId", r.MessageId }, { "status", r.Status }, { "receivedDateTime", r.Received.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture) },
                    { "recipientAddress", r.Recipient }, { "senderAddress", r.Sender }, { "subject", r.Subject }, { "size", 1000 }, { "fromIP", "10.0.0.1" }, { "toIP", "" } }).ToList() }
            };
            if (skip + top < all.Count)
                body["@odata.nextLink"] = request.RequestUri.GetLeftPart(UriPartial.Path) + "?$filter=" + Uri.EscapeDataString(filter) + "&$top=" + top + "&$skiptoken=" + (skip + top);
            return Json(HttpStatusCode.OK, JsonSerializer.Serialize(body));
        }
    }
}
