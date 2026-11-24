/**
 * Chat website for the Text Forwarder.
 *
 * Adds a private chat page (Index.html) on top of Code.gs:
 *   - incoming texts are saved to a Google Sheet, one chat per friend
 *   - messages typed on the page are queued, and the iPhone reply Shortcut
 *     picks them up through the existing { action: "replies" } request
 *   - the page can rename friends, add reactions and send GIF links
 *
 * Only numbers listed in Config.gs can ever be texted from the page.
 * The page is protected by SITE_PASSCODE (or SHARED_SECRET if that is unset).
 */

var CHAT_COLS = ['id', 'thread', 'dir', 'text', 'ts', 'status', 'reaction', 'kind'];
var CHAT_REACTIONS = ['👍', '❤️', '😂', '😮', '😢', '🙏'];
var CHAT_MAX_ROWS = 2000;
var CHAT_MAX_TEXT = 1000;
var CHAT_MAX_QUEUE = 50;

// ─────────────────── FUNCTIONS THE PAGE CALLS ───────────────────
// Each one checks the passcode first. Writes take the script lock so two
// requests cannot interleave.

function apiState(pass, since) {
  return siteState(pass, since, CONFIG, realDeps_());
}

function apiSend(pass, thread, text, kind) {
  return withLock_(function () {
    return siteSend(pass, thread, text, kind, CONFIG, realDeps_());
  });
}

function apiReact(pass, id, emoji) {
  return withLock_(function () {
    return siteReact(pass, id, emoji, CONFIG, realDeps_());
  });
}

function apiRename(pass, thread, name) {
  return withLock_(function () {
    return siteRename(pass, thread, name, CONFIG, realDeps_());
  });
}

function apiGifs(pass, query) {
  if (!siteAuthOk(pass, CONFIG)) return { ok: false, reason: 'bad_passcode' };
  if (!CONFIG.GIPHY_API_KEY) return { ok: false, reason: 'gifs_off' };
  var q = String(query || '').trim().slice(0, 60);
  var url =
    'https://api.giphy.com/v1/gifs/' + (q ? 'search' : 'trending') +
    '?api_key=' + encodeURIComponent(CONFIG.GIPHY_API_KEY) +
    '&limit=18&rating=g' + (q ? '&q=' + encodeURIComponent(q) : '');
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, reason: 'gif_lookup_failed' };
    return { ok: true, gifs: gifResults(JSON.parse(res.getContentText())) };
  } catch (err) {
    return { ok: false, reason: 'gif_lookup_failed' };
  }
}

/**
 * Run this once from the editor after pasting the files. It creates the
 * Google Sheet that stores the chats and prints the two things you need:
 * the ntfy topic for the iPhone and where the Sheet lives.
 */
function setupChat() {
  var deps = realDeps_();
  deps.store.version();
  var info = {
    sheet: chatSheet_().getUrl(),
    ntfy_topic: ntfyTopic(CONFIG, deps.hash),
    passcode_is: CONFIG.SITE_PASSCODE ? 'SITE_PASSCODE from Config' : 'SHARED_SECRET from Config'
  };
  console.log(JSON.stringify(info, null, 2));
  return info;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return { ok: false, reason: 'busy' };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────── GOOGLE SHEET STORAGE ─────────────────────

var chatSheetCache_ = null;

function chatSheet_() {
  if (chatSheetCache_) return chatSheetCache_;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('chat_sheet_id');
  var ss = null;
  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (err) {
      ss = null;
    }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Text Forwarder Chats');
    props.setProperty('chat_sheet_id', ss.getId());
    var first = ss.getSheets()[0];
    first.setName('messages');
    first.appendRow(CHAT_COLS);
    ss.insertSheet('names').appendRow(['key', 'name']);
  }
  chatSheetCache_ = ss;
  return ss;
}

/**
 * Storage used by the chat logic. Free text is saved with a "t:" prefix so
 * the Sheet can never read a message such as "=SUM(...)" as a formula.
 */
function chatDeps_() {
  var props = PropertiesService.getScriptProperties();
  var store = {
    listMessages: function () {
      var sh = chatSheet_().getSheetByName('messages');
      var last = sh.getLastRow();
      if (last < 2) return [];
      var start = Math.max(2, last - CHAT_MAX_ROWS + 1);
      var rows = sh.getRange(start, 1, last - start + 1, CHAT_COLS.length).getValues();
      return rows.map(rowToMessage).filter(function (m) {
        return m.id;
      });
    },
    appendMessage: function (m) {
      chatSheet_().getSheetByName('messages').appendRow(messageToRow(m));
    },
    updateMessage: function (id, patch) {
      var sh = chatSheet_().getSheetByName('messages');
      var last = sh.getLastRow();
      if (last < 2) return false;
      var ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(id)) {
          if (patch.status != null) sh.getRange(i + 2, 6).setValue(patch.status);
          if (patch.reaction != null) sh.getRange(i + 2, 7).setValue('t:' + patch.reaction);
          return true;
        }
      }
      return false;
    },
    getNames: function () {
      var sh = chatSheet_().getSheetByName('names');
      var last = sh.getLastRow();
      var out = {};
      if (last < 2) return out;
      sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
        var name = String(r[1] || '').replace(/^t:/, '');
        if (r[0] && name) out[String(r[0])] = name;
      });
      return out;
    },
    setName: function (key, name) {
      var sh = chatSheet_().getSheetByName('names');
      var last = sh.getLastRow();
      if (last >= 2) {
        var keys = sh.getRange(2, 1, last - 1, 1).getValues();
        for (var i = 0; i < keys.length; i++) {
          if (String(keys[i][0]) === key) {
            sh.getRange(i + 2, 2).setValue('t:' + name);
            return;
          }
        }
      }
      sh.appendRow([key, 't:' + name]);
    },
    version: function () {
      chatSheet_();
      return Number(props.getProperty('chat_version') || 0);
    },
    bump: function () {
      props.setProperty('chat_version', String(Number(props.getProperty('chat_version') || 0) + 1));
    }
  };
  return {
    store: store,
    newId: function () {
      return Utilities.getUuid();
    },
    wake: function () {
      var deps = realDeps_();
      UrlFetchApp.fetch('https://ntfy.sh/' + ntfyTopic(CONFIG, deps.hash), {
        method: 'post',
        payload: 'Message waiting to send',
        muteHttpExceptions: true
      });
    }
  };
}

function messageToRow(m) {
  return [m.id, m.thread, m.dir, 't:' + m.text, m.ts, m.status, 't:' + (m.reaction || ''), m.kind || 'text'];
}

function rowToMessage(r) {
  return {
    id: String(r[0] || ''),
    thread: String(r[1] || ''),
    dir: String(r[2] || ''),
    text: String(r[3] == null ? '' : r[3]).replace(/^t:/, ''),
    ts: Number(r[4]) || 0,
    status: String(r[5] || ''),
    reaction: String(r[6] == null ? '' : r[6]).replace(/^t:/, ''),
    kind: String(r[7] || 'text')
  };
}

// ────────────────────────── CHAT LOGIC ──────────────────────────
// Pure functions: storage and network arrive through `deps`, so this part
// is unit-tested outside Apps Script.

/** True when `pass` is the site passcode. Never true for the placeholder. */
function siteAuthOk(pass, config) {
  var want = String(config.SITE_PASSCODE || config.SHARED_SECRET || '');
  if (!want || want.indexOf('CHANGE-ME') === 0) return false;
  return String(pass == null ? '' : pass) === want;
}

/** '+1 (650) 555-0101' -> 'p:6505550101'. '' if not a phone number. */
function threadKey(phone) {
  var p = normalizePhone(phone);
  return p ? 'p:' + p : '';
}

/** Friends from Config, with any names chosen on the page applied. */
function friendList(config, names) {
  var seen = {};
  var out = [];
  (config.SENDERS || []).map(normalizeEntry).forEach(function (e) {
    var key = threadKey(e.phone);
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push({ key: key, phone: e.phone, defaultName: e.label, name: (names && names[key]) || e.label });
  });
  return out;
}

function friendByKey(key, config) {
  var list = (config.SENDERS || []).map(normalizeEntry);
  for (var i = 0; i < list.length; i++) {
    if (threadKey(list[i].phone) && threadKey(list[i].phone) === key) return list[i];
  }
  return null;
}

function isGifUrl(text) {
  return /^https:\/\/([a-z0-9-]+\.)*giphy\.com\/\S+$/i.test(String(text || '').trim());
}

/** Saves a text that arrived on the phone into that friend's chat. */
function recordIncoming(entry, text, deps) {
  deps.store.appendMessage({
    id: deps.newId(),
    thread: threadKey(entry.phone),
    dir: 'in',
    text: text,
    ts: deps.now().getTime(),
    status: 'received',
    reaction: '',
    kind: isGifUrl(text) ? 'gif' : 'text'
  });
  deps.store.bump();
}

/** Everything the page needs to draw itself, or "nothing changed". */
function siteState(pass, since, config, deps) {
  if (!siteAuthOk(pass, config)) return { ok: false, reason: 'bad_passcode' };
  var version = deps.store.version();
  if (since != null && since !== '' && String(since) === String(version)) {
    return { ok: true, changed: false, version: version };
  }
  return {
    ok: true,
    changed: true,
    version: version,
    friends: friendList(config, deps.store.getNames()),
    messages: deps.store.listMessages(),
    reactions: CHAT_REACTIONS,
    gifs: !!config.GIPHY_API_KEY
  };
}

/** Queues a message typed on the page and nudges the phone to send it. */
function siteSend(pass, thread, text, kind, config, deps) {
  if (!siteAuthOk(pass, config)) return { ok: false, reason: 'bad_passcode' };
  if (!friendByKey(thread, config)) return { ok: false, reason: 'unknown_friend' };
  var t = String(text == null ? '' : text).trim();
  if (!t) return { ok: false, reason: 'empty' };
  if (t.length > CHAT_MAX_TEXT) return { ok: false, reason: 'too_long' };
  var k = kind === 'gif' ? 'gif' : 'text';
  if (k === 'gif' && !isGifUrl(t)) return { ok: false, reason: 'bad_gif' };

  var queued = deps.store.listMessages().filter(function (m) {
    return m.dir === 'out' && m.status === 'queued';
  }).length;
  if (queued >= CHAT_MAX_QUEUE) return { ok: false, reason: 'queue_full' };

  var id = queueOutgoing(thread, t, k, deps);
  return { ok: true, id: id };
}

function queueOutgoing(thread, text, kind, deps) {
  var id = deps.newId();
  deps.store.appendMessage({
    id: id,
    thread: thread,
    dir: 'out',
    text: text,
    ts: deps.now().getTime(),
    status: 'queued',
    reaction: '',
    kind: kind
  });
  deps.store.bump();
  try {
    deps.wake();
  } catch (err) {
    deps.log('wake failed: ' + err);
  }
  return id;
}

/** Sets or clears the reaction on one message. */
function siteReact(pass, id, emoji, config, deps) {
  if (!siteAuthOk(pass, config)) return { ok: false, reason: 'bad_passcode' };
  var e = String(emoji || '');
  if (e && CHAT_REACTIONS.indexOf(e) === -1) return { ok: false, reason: 'bad_reaction' };
  var msg = null;
  deps.store.listMessages().forEach(function (m) {
    if (m.id === id) msg = m;
  });
  if (!msg) return { ok: false, reason: 'unknown_message' };
  if (!deps.store.updateMessage(id, { reaction: e })) return { ok: false, reason: 'unknown_message' };
  deps.store.bump();

  // Friends cannot see a reaction made on the page, so optionally text it.
  if (config.SEND_REACTIONS === true && e && msg.dir === 'in' && friendByKey(msg.thread, config)) {
    var snippet = msg.kind === 'gif' ? 'your GIF' : '"' + msg.text.slice(0, 60) + (msg.text.length > 60 ? '…' : '') + '"';
    queueOutgoing(msg.thread, e + ' to ' + snippet, 'text', deps);
  }
  return { ok: true };
}

/** Sets the name shown for a friend. An empty name restores the default. */
function siteRename(pass, thread, name, config, deps) {
  if (!siteAuthOk(pass, config)) return { ok: false, reason: 'bad_passcode' };
  if (!friendByKey(thread, config)) return { ok: false, reason: 'unknown_friend' };
  var n = String(name == null ? '' : name).replace(/\s+/g, ' ').trim();
  if (n.length > 40) return { ok: false, reason: 'too_long' };
  deps.store.setName(thread, n);
  deps.store.bump();
  return { ok: true, name: n };
}

/**
 * Hands the phone the oldest queued message from the page, in the same
 * shape as an email reply: { ok, has_reply, to, label, text }.
 * A message is marked sent before it is returned, so it is texted at most
 * once. Messages queued for longer than the reply age limit are expired
 * instead of being sent late.
 */
function nextOutgoing(config, deps) {
  var maxAgeMin = config.REPLY_MAX_AGE_MINUTES > 0 ? config.REPLY_MAX_AGE_MINUTES : 120;
  var cutoff = deps.now().getTime() - maxAgeMin * 60 * 1000;
  var queued = deps.store.listMessages().filter(function (m) {
    return m.dir === 'out' && m.status === 'queued';
  });
  var changed = false;
  for (var i = 0; i < queued.length; i++) {
    var m = queued[i];
    var friend = friendByKey(m.thread, config);
    if (!friend) {
      deps.store.updateMessage(m.id, { status: 'failed' });
      changed = true;
      continue;
    }
    if (m.ts < cutoff) {
      deps.store.updateMessage(m.id, { status: 'expired' });
      changed = true;
      continue;
    }
    deps.store.updateMessage(m.id, { status: 'sent' });
    deps.store.bump();
    if (i < queued.length - 1) {
      try {
        deps.wake(); // more are waiting: nudge the phone again
      } catch (err) {
        deps.log('wake failed: ' + err);
      }
    }
    deps.log('page message handed to phone for ' + friend.label);
    return { ok: true, has_reply: true, to: friend.phone, label: friend.label, text: m.text };
  }
  if (changed) deps.store.bump();
  return { ok: true, has_reply: false };
}

/** The private ntfy topic the phone listens on. Derived from the secret. */
function ntfyTopic(config, hashFn) {
  if (config.NTFY_TOPIC) return String(config.NTFY_TOPIC);
  var h = String(hashFn('ntfy:' + config.SHARED_SECRET)).toLowerCase().replace(/[^a-z0-9]/g, '');
  return 'tf' + h.slice(0, 24);
}

/** Picks the fields the page needs out of a GIPHY API response. */
function gifResults(body) {
  var data = (body && body.data) || [];
  var out = [];
  data.forEach(function (g) {
    if (!g || !/^[A-Za-z0-9]+$/.test(String(g.id || ''))) return;
    var small = g.images && (g.images.fixed_width_small || g.images.fixed_width);
    var preview = small && small.url;
    if (!isGifUrl(preview)) return;
    out.push({ preview: preview, url: 'https://media.giphy.com/media/' + g.id + '/giphy.gif' });
  });
  return out;
}

// Lets the test file load this script in Node. Ignored by Apps Script.
if (typeof module !== 'undefined') {
  var chatExports_ = {
    siteAuthOk: siteAuthOk,
    threadKey: threadKey,
    friendList: friendList,
    friendByKey: friendByKey,
    isGifUrl: isGifUrl,
    recordIncoming: recordIncoming,
    siteState: siteState,
    siteSend: siteSend,
    siteReact: siteReact,
    siteRename: siteRename,
    nextOutgoing: nextOutgoing,
    ntfyTopic: ntfyTopic,
    gifResults: gifResults,
    messageToRow: messageToRow,
    rowToMessage: rowToMessage,
    CHAT_REACTIONS: CHAT_REACTIONS
  };
  for (var chatKey_ in chatExports_) module.exports[chatKey_] = chatExports_[chatKey_];
}
