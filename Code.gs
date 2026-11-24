/**
 * Text-to-Email Forwarder (prototype)
 *
 * A Google Apps Script web app. An iPhone Shortcuts automation POSTs each
 * incoming text here as JSON; this script checks the sender against an
 * allowlist and emails the text to FORWARD_TO.
 *
 * Expected request body (JSON):
 *   { "secret": "...", "sender": "+1 (650) 555-0101" or "Arjun", "body": "message text" }
 *
 * Response (JSON): { ok, forwarded, reason? }
 *
 * Reply path: a second Shortcut POSTs { "secret": "...", "action": "replies" }.
 * The script looks in Gmail for a new reply to a "Texts from <name>" email,
 * and answers with { ok, has_reply, to, label, text } so the phone can text
 * it to that friend. Each reply is handed out once.
 *
 * Chat website: Chat.gs and Index.html add a private chat page at the web
 * app URL. Incoming texts are saved there, and messages typed on the page
 * go out through the same "replies" request.
 *
 * Settings (numbers to forward, destination email, secret) live in Config.gs.
 */

// ─────────────────────────── ENTRY POINTS ──────────────────────────

/**
 * Opening the web app URL shows the chat page.
 * Add ?health=1 to the URL for the plain "is it live" check.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.health) {
    return jsonOut_({ ok: true, service: 'text-to-email-forwarder' });
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Chats')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called by the iPhone Shortcuts for incoming texts and for replies. */
function doPost(e) {
  var raw = e && e.postData && e.postData.contents;
  // One request at a time, so two triggers firing together cannot
  // forward the same text twice or hand out the same reply twice.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return jsonOut_({ ok: false, forwarded: false, reason: 'busy' });
  }
  try {
    return jsonOut_(handleRequest(raw, CONFIG, realDeps_()));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Manual test: pick this function in the Apps Script editor and press Run.
 * It pretends the first number in Config.gs just texted, and should put
 * an email in the FORWARD_TO inbox. The result shows in the Execution log.
 */
function testForward() {
  var first = normalizeEntry(CONFIG.SENDERS[0]);
  var raw = JSON.stringify({
    secret: CONFIG.SHARED_SECRET,
    sender: first.phone,
    body: 'Test message sent at ' + new Date().toISOString()
  });
  var result = handleRequest(raw, CONFIG, realDeps_());
  console.log(JSON.stringify(result));
  return result;
}

/**
 * Manual test for the reply path: reply to a "Texts from ..." email first,
 * then Run this. It shows what the phone would be told to send, WITHOUT
 * marking the reply as handled, so the phone can still pick it up.
 */
function testReplies() {
  var deps = realDeps_();
  deps.markHandled = function () {};
  var result = nextReply(CONFIG, deps);
  console.log(JSON.stringify(result));
  return result;
}

function realDeps_() {
  var deps = baseDeps_();
  var chat = chatDeps_(); // from Chat.gs: store, newId, wake
  for (var k in chat) deps[k] = chat[k];
  return deps;
}

function baseDeps_() {
  return {
    findReplies: function (fromAddress) {
      var query = 'subject:"Texts from" from:' + fromAddress + ' newer_than:2d';
      var threads = GmailApp.search(query, 0, 20);
      var out = [];
      threads.forEach(function (thread) {
        thread.getMessages().forEach(function (m) {
          out.push({
            id: m.getId(),
            from: m.getFrom(),
            subject: m.getSubject(),
            body: m.getPlainBody(),
            time: m.getDate().getTime()
          });
        });
      });
      return out;
    },
    isHandled: function (id) {
      return handledIds_().indexOf(id) !== -1;
    },
    markHandled: function (id) {
      var ids = handledIds_();
      ids.push(id);
      if (ids.length > 300) ids = ids.slice(ids.length - 300);
      PropertiesService.getScriptProperties().setProperty('handled_ids', JSON.stringify(ids));
    },
    sendEmail: function (to, subject, body, fromName) {
      MailApp.sendEmail(to, subject, body, { name: fromName });
    },
    cacheGet: function (key) {
      return CacheService.getScriptCache().get(key);
    },
    cachePut: function (key, seconds) {
      CacheService.getScriptCache().put(key, '1', seconds);
    },
    hash: function (text) {
      var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
      return Utilities.base64EncodeWebSafe(bytes);
    },
    now: function () {
      return new Date();
    },
    log: function (line) {
      console.log(line);
    }
  };
}

function handledIds_() {
  var raw = PropertiesService.getScriptProperties().getProperty('handled_ids');
  try {
    return raw ? JSON.parse(raw) : [];
  } catch (err) {
    return [];
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ──────────────────────────── CORE LOGIC ───────────────────────────
// Pure functions: no Google services are touched directly, so this part
// can be unit-tested outside Apps Script.

/**
 * @param {string} raw     Raw POST body.
 * @param {object} config  CONFIG-shaped object.
 * @param {object} deps    { sendEmail, cacheGet, cachePut, hash, now, log,
 *                           findReplies, isHandled, markHandled,
 *                           store?, newId?, wake? }
 * @return {{ok: boolean, forwarded: boolean, reason?: string}}
 */
function handleRequest(raw, config, deps) {
  var payload;
  try {
    payload = JSON.parse(raw);
  } catch (err) {
    return { ok: false, forwarded: false, reason: 'bad_json' };
  }
  if (!payload || typeof payload !== 'object') {
    return { ok: false, forwarded: false, reason: 'bad_json' };
  }

  // Refuse to run with the placeholder secret, so a half-configured
  // deployment can never forward anything.
  if (!config.SHARED_SECRET || config.SHARED_SECRET.indexOf('CHANGE-ME') === 0) {
    return { ok: false, forwarded: false, reason: 'secret_not_configured' };
  }
  if (String(payload.secret || '') !== config.SHARED_SECRET) {
    return { ok: false, forwarded: false, reason: 'unauthorized' };
  }

  if (payload.action === 'replies') {
    // Messages typed on the chat page go first, then replies sent by email.
    if (deps.store) {
      try {
        var out = nextOutgoing(config, deps);
        if (out.has_reply) return out;
      } catch (err) {
        deps.log('chat outbox failed: ' + err);
      }
    }
    return nextReply(config, deps);
  }

  var senderRaw = String(payload.sender || '').trim();
  if (!senderRaw) {
    return { ok: false, forwarded: false, reason: 'missing_sender' };
  }

  var match = findSender(senderRaw, config.SENDERS);
  if (!match) {
    deps.log('skipped: sender not on allowlist');
    return { ok: true, forwarded: false, reason: 'sender_not_allowed' };
  }

  var text = String(payload.body == null ? '' : payload.body).trim();
  if (!text) {
    text = '[No text in this message. It may be a photo, attachment or reaction.]';
  }

  var dedupeKey = 'fwd:' + deps.hash(match.label + '\n' + text);
  if (config.DEDUPE_SECONDS > 0 && deps.cacheGet(dedupeKey)) {
    deps.log('skipped: duplicate from ' + match.label);
    return { ok: true, forwarded: false, reason: 'duplicate' };
  }

  // Save to the chat page (when Chat.gs is installed).
  var stored = false;
  if (deps.store) {
    try {
      recordIncoming(match, text, deps);
      stored = true;
    } catch (err) {
      deps.log('chat save failed: ' + err);
    }
  }

  // Email as well, unless Config says ALSO_EMAIL: false.
  var wantEmail = config.ALSO_EMAIL !== false && !!config.FORWARD_TO;
  if (wantEmail) {
    var email = buildEmail(match.label, text, deps.now());
    try {
      deps.sendEmail(config.FORWARD_TO, email.subject, email.body, email.fromName);
    } catch (err) {
      deps.log('send failed: ' + err);
      if (!stored) return { ok: false, forwarded: false, reason: 'send_failed' };
    }
  } else if (!stored) {
    return { ok: false, forwarded: false, reason: 'nowhere_to_deliver' };
  }

  // Mark as done only after it was delivered, so a failed send can be retried.
  if (config.DEDUPE_SECONDS > 0) {
    deps.cachePut(dedupeKey, config.DEDUPE_SECONDS);
  }
  deps.log('forwarded text from ' + match.label);
  return { ok: true, forwarded: true };
}

/** Last 10 digits of anything phone-like; '' if fewer than 7 digits. */
function normalizePhone(value) {
  var digits = String(value || '').replace(/\D/g, '');
  if (digits.length < 7) return '';
  return digits.length > 10 ? digits.slice(-10) : digits;
}

/** True when the string is a phone number rather than a contact name. */
function looksLikePhone(value) {
  return /^[\d\s()+\-.]+$/.test(String(value || '').trim()) && normalizePhone(value) !== '';
}

/**
 * A Config.gs entry is either a bare number ('+16505550101') or an object
 * ({ phone, label, aliases }). Returns the object form either way.
 */
function normalizeEntry(entry) {
  if (typeof entry === 'string') {
    return { label: entry, phone: entry, aliases: [] };
  }
  var e = entry || {};
  return { label: e.label || e.phone || '', phone: e.phone || '', aliases: e.aliases || [] };
}

/**
 * Finds the allowlist entry for a sender string, which may be a phone
 * number in any format or a contact name. Returns the entry or null.
 */
function findSender(senderRaw, allowed) {
  var list = (allowed || []).map(normalizeEntry);
  var i, j;

  if (looksLikePhone(senderRaw)) {
    var phone = normalizePhone(senderRaw);
    for (i = 0; i < list.length; i++) {
      if (normalizePhone(list[i].phone) === phone) return list[i];
    }
    return null;
  }

  var name = senderRaw.toLowerCase().replace(/\s+/g, ' ').trim();
  for (i = 0; i < list.length; i++) {
    var names = [list[i].label].concat(list[i].aliases || []);
    for (j = 0; j < names.length; j++) {
      if (String(names[j]).toLowerCase().replace(/\s+/g, ' ').trim() === name) return list[i];
    }
  }
  return null;
}

/**
 * One fixed subject per sender, so Gmail groups each friend's texts
 * into a single conversation.
 */
function buildEmail(label, text, when) {
  return {
    subject: 'Texts from ' + label,
    fromName: label + ' (text)',
    body: text + '\n\n-- \nForwarded text from ' + label + ' at ' + when.toISOString()
  };
}

// ──────────────────────────── REPLY PATH ───────────────────────────

/**
 * Finds the oldest unhandled reply and returns it for the phone to text.
 * A reply counts only if it is from the address the texts are forwarded to,
 * its subject is "Re: Texts from <name>", <name> is in Config, and it is
 * recent. Each reply is marked handled before it is returned, so it can
 * never be texted twice.
 *
 * @return {{ok: boolean, has_reply: boolean, to?: string, label?: string, text?: string}}
 */
function nextReply(config, deps) {
  var replyFrom = String(config.REPLY_FROM || config.FORWARD_TO || '').trim();
  if (!replyFrom) {
    return { ok: false, has_reply: false, reason: 'reply_from_not_configured' };
  }
  var maxAgeMin = config.REPLY_MAX_AGE_MINUTES > 0 ? config.REPLY_MAX_AGE_MINUTES : 120;
  var cutoff = deps.now().getTime() - maxAgeMin * 60 * 1000;

  var messages;
  try {
    messages = deps.findReplies(replyFrom) || [];
  } catch (err) {
    deps.log('reply lookup failed: ' + err);
    return { ok: false, has_reply: false, reason: 'lookup_failed' };
  }
  messages.sort(function (a, b) {
    return a.time - b.time;
  });

  for (var i = 0; i < messages.length; i++) {
    var m = messages[i];
    if (deps.isHandled(m.id)) continue;
    if (emailAddress(m.from) !== emailAddress(replyFrom)) continue;
    if (m.time < cutoff) continue;

    var label = labelFromSubject(m.subject);
    if (!label) continue; // not a reply to one of our emails

    // Our own forwarded emails carry this footer unquoted. Never treat one
    // as a reply, or a friend's text would be sent straight back to them.
    if (isOwnForward(m.body)) continue;

    var entry = findSender(label, config.SENDERS);
    var text = stripQuoted(m.body);

    if (!entry || !normalizePhone(entry.phone)) {
      deps.markHandled(m.id);
      deps.log('reply skipped: no number for "' + label + '"');
      continue;
    }
    if (!text) {
      deps.markHandled(m.id);
      deps.log('reply skipped: empty reply to ' + entry.label);
      continue;
    }

    deps.markHandled(m.id);
    deps.log('reply handed to phone for ' + entry.label);
    return { ok: true, has_reply: true, to: entry.phone, label: entry.label, text: text };
  }
  return { ok: true, has_reply: false };
}

/** 'Name <a@b.com>' -> 'a@b.com', lower-cased. */
function emailAddress(value) {
  var s = String(value || '');
  var m = s.match(/<([^>]+)>/);
  return (m ? m[1] : s).trim().toLowerCase();
}

/** 'Re: RE: Texts from Arjun' -> 'Arjun'. '' unless it is a reply. */
function labelFromSubject(subject) {
  var s = String(subject || '').trim();
  if (!/^(re|aw|sv)\s*:/i.test(s)) return '';
  s = s.replace(/^((re|aw|sv|fwd?)\s*:\s*)+/i, '');
  var m = s.match(/^Texts from (.+)$/i);
  return m ? m[1].trim() : '';
}

/** True when the unquoted part of the email holds our forwarding footer. */
function isOwnForward(body) {
  var lines = String(body || '').replace(/\r/g, '').split('\n');
  for (var i = 0; i < lines.length; i++) {
    if (/^Forwarded text from .+ at \d{4}-\d{2}-\d{2}T/.test(lines[i].trim())) return true;
  }
  return false;
}

/**
 * Keeps only what the person typed: cuts the email at the first sign of
 * quoted history or a signature, and drops "Sent from my iPhone" lines.
 */
function stripQuoted(body) {
  var lines = String(body || '').replace(/\r/g, '').split('\n');
  var kept = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var t = line.trim();
    if (/^>/.test(t)) break;
    if (/^--\s*$/.test(line)) break; // signature delimiter
    if (/^-{2,}\s*(original|forwarded) message/i.test(t)) break;
    if (/^_{5,}$/.test(t)) break;
    if (/^On\b.*\bwrote:$/i.test(t)) break;
    // Gmail wraps a long "On <date> <name> wrote:" line across two or three lines.
    if (/^On\b/i.test(t)) {
      var joined = t;
      var hit = false;
      for (var j = i + 1; j <= i + 2 && j < lines.length; j++) {
        if (!lines[j].trim()) break; // a wrapped header never spans a blank line
        joined += ' ' + lines[j].trim();
        if (/\bwrote:$/i.test(joined)) {
          hit = true;
          break;
        }
      }
      if (hit) break;
    }
    if (/^Sent from my /i.test(t)) continue;
    kept.push(line);
  }
  return kept.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

// Lets the test file load this script in Node. Ignored by Apps Script.
if (typeof module !== 'undefined') {
  module.exports = {
    handleRequest: handleRequest,
    nextReply: nextReply,
    emailAddress: emailAddress,
    labelFromSubject: labelFromSubject,
    isOwnForward: isOwnForward,
    stripQuoted: stripQuoted,
    normalizeEntry: normalizeEntry,
    normalizePhone: normalizePhone,
    looksLikePhone: looksLikePhone,
    findSender: findSender,
    buildEmail: buildEmail
  };
}
