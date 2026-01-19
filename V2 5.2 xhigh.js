/**
 * SmartEmailProcessor v2.0
 * Google Apps Script (Gmail -> Gemini -> Calendar/Tasks + 智能标签 + 每日日报)
 *
 * 核心设计：
 * - 代码层：提取/清洗/去噪/幂等/标签体系/执行(Calendar/Tasks)/日志与日报
 * - AI层：理解与决策（分类/优先级/时间地点/关键链接/二维码与图片信息/下一步建议）
 *
 * 使用前：
 * - Apps Script -> Services 开启 “Tasks API”(可选但推荐，写入 Google Tasks 需要)
 * - 运行 setGeminiApiKey('你的Key') 或在 CONFIG.GEMINI.API_KEY 里写死
 * - 运行一次 setupSmartEmailProcessor()
 */

const CONFIG = {
  VERSION: '2.0.0',

  GEMINI: {
    API_KEY: '', // 推荐留空/保持占位，用 setGeminiApiKey() 写入 Script Properties
    MODEL_NAME: 'gemini-3-flash-preview',
    API_ROOT: 'https://generativelanguage.googleapis.com/v1beta/models',

    TEMPERATURE: 0.2,

    // Gemini 3 推荐：thinkingLevel
    THINKING_LEVEL: 'HIGH', // LOW | HIGH

    // 多模态分辨率（全局）
    MEDIA_RESOLUTION: 'MEDIA_RESOLUTION_HIGH', // LOW | MEDIUM | HIGH | ULTRA_HIGH

    USE_RESPONSE_SCHEMA: true,
  },

  GMAIL: {
    SOURCE_LABEL: 'PolyU', // 你用于收集 Outlook 转发邮件的标签
    AI_ROOT_LABEL: 'AI',

    STATUS_LABELS: {
      PROCESSING: 'AI/⏳ 处理中',
      PROCESSED: 'AI/✅ 已处理',
      REVIEW: 'AI/⚠️ 待复核',
      ERROR: 'AI/❌ 处理失败',
      SKIPPED: 'AI/🧹 规则跳过',
    },

    CATEGORY_LABELS: {
      EVENT: 'AI/📅 日程',
      TASK: 'AI/✅ 待办',
      REMINDER: 'AI/🔔 提醒',
      PROMO: 'AI/📢 推广',
      INFO: 'AI/📄 信息',
    },

    PRIORITY_LABELS: {
      HIGH: 'AI/🔴 高优先',
      MEDIUM: 'AI/🟡 中优先',
      LOW: 'AI/🟢 低优先',
    },

    SYNC_LABELS: {
      CALENDAR: 'AI/🔁 已同步/Calendar',
      TASKS: 'AI/🔁 已同步/Tasks',
    },

    TRAIT_LABELS: {
      HAS_QR: 'AI/🧾 含二维码/票据',
      HAS_CODE: 'AI/🔢 含验证码/编号',
      HAS_MEETING_LINK: 'AI/🎥 含会议链接',
      HAS_PAYMENT: 'AI/💳 含支付/金额',
      HAS_ATTACHMENT: 'AI/📎 含附件',
    },

    EXTRA_LABEL_PREFIX: 'AI/🏷️ ', // AI 建议的语义标签会挂在这里
  },

  PROCESSING: {
    // API 不支持 batch：每次触发只处理 1 个线程（用幂等保证不会重复扣费）
    MAX_THREADS_SCAN: 30, // 从来源标签中扫描多少个线程，寻找“下一封要处理的”
    MAX_BODY_CHARS: 22000,
    MAX_HTML_SNIPPET_CHARS: 5000,
    MAX_LINKS: 40,

    // 多模态：图片 + PDF 统一当 media
    MAX_MEDIA_ITEMS: 8,
    MAX_IMAGE_ITEMS: 5,
    MAX_PDF_ITEMS: 2,
    MAX_MEDIA_BYTES_EACH: 6 * 1024 * 1024,
    MAX_TOTAL_MEDIA_BYTES: 14 * 1024 * 1024,
    MIN_IMAGE_BYTES: 8 * 1024, // 小于这个大概率是 logo

    MAX_CONVERSATION_MESSAGES: 2,
    MAX_CONVERSATION_SNIPPET_CHARS: 800,

    MIN_CONFIDENCE_TO_AUTO_CREATE_EVENT: 0.65,
    ALWAYS_CREATE_TASK_WHEN_REQUIRES_REVIEW: true,

    ENABLE_HEURISTIC_PROMO_BYPASS: true, // 极明显 newsletter 可跳过 AI（省请求次数）
    PROMO_BYPASS_THRESHOLD: 0.92,

    DRY_RUN: false, // true：只打标签/写日志，不创建 Calendar/Tasks
    MARK_READ_AFTER_PROCESSING: false,
    ARCHIVE_AFTER_PROCESSING: false,
  },

  DAILY_REPORT: {
    ENABLED: true,
    RECIPIENT_EMAIL: 'your_gmail@gmail.com',
    HOUR: 22,

    LOG_SPREADSHEET_NAME: 'SmartEmailProcessor Log',
    LOG_SHEET_TAB: 'log',
    MAX_ROWS_READ: 2000,
    MAX_ITEMS_IN_PROMPT: 60,

    TEMPERATURE: 0.6,
    THINKING_LEVEL: 'HIGH',
  },
};

const CACHE = {
  labels: {},
  log: { ssId: null, sheet: null },
};

/**
 * =========================
 * 一键初始化（建议只跑一次）
 * =========================
 */
function setupSmartEmailProcessor() {
  ensureLabelsExist_();
  ensureLogSheet_();
  setupTriggers();
  Logger.log('✅ 初始化完成：标签 + 日志表 + 触发器');
}


/**
 * 保存 Gemini API Key 到 Script Properties（更安全，避免写进代码）
 */
function setGeminiApiKey(apiKey) {
  const key = String(apiKey || '').trim();
  if (key.length < 20) throw new Error('apiKey 为空或太短');
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('✅ 已保存 GEMINI_API_KEY 到 Script Properties');
}

/**
 * 触发器：每 5 分钟处理 1 个线程；每天 22 点发送日报
 */
function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) ScriptApp.deleteTrigger(t);

  ScriptApp.newTrigger('processEmails').timeBased().everyMinutes(5).create();

  if (CONFIG.DAILY_REPORT.ENABLED) {
    ScriptApp.newTrigger('sendDailyReport')
      .timeBased()
      .atHour(CONFIG.DAILY_REPORT.HOUR)
      .everyDays(1)
      .create();
  }

  Logger.log('✅ 触发器已设置');
}

/**
 * =========================
 * 主流程：每次处理 1 个线程
 * =========================
 */
function processEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('⏭️ 上一次还没跑完，跳过本次触发');
    return;
  }

  let thread = null;

  try {
    ensureLabelsExist_();
    ensureLogSheet_();

    thread = getNextThreadToProcess_();
    if (!thread) {
      Logger.log('✓ 没有新邮件需要处理');
      return;
    }

    // 取线程最新一封作为本次处理对象
    const messages = thread.getMessages();
    const latest = messages[messages.length - 1];
    const threadId = thread.getId();
    const latestMessageId = latest.getId();
    const tz = Session.getScriptTimeZone();
    const now = new Date();

    // 幂等：同一线程同一 messageId 只处理一次（避免重复扣费/重复创建日程）
    const lastProcessed = getLastProcessedMessageId_(threadId);
    if (lastProcessed === latestMessageId) {
      Logger.log(`⏭️ 幂等跳过（已处理过这封）：${latest.getSubject()}`);
      return;
    }

    addLabel_(thread, CONFIG.GMAIL.AI_ROOT_LABEL);
    addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.PROCESSING);

    // 若线程有新回复：先清掉上一轮 AI 分类/优先级/同步等标签，再重新贴
    clearAiDerivedLabels_(thread);

    const email = extractEmailData_(thread, messages, tz);

    // （可选）极明显的推广/Newsletter：直接规则跳过 AI，省请求次数
    const bypass = maybeBypassAi_(email);
    if (bypass.bypassed) {
      const exec = applyResultActions_(thread, email, bypass.result, tz);
      addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.SKIPPED);
      addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.PROCESSED);

      appendLogRow_(email, bypass.result, exec);
      setLastProcessedMessageId_(threadId, latestMessageId);

      Logger.log(`🧹 规则跳过AI：${email.subject}`);
      return;
    }

    const aiRaw = callGeminiForEmail_(email, tz, now);
    const result = normalizeAiResult_(aiRaw, email);

    const exec = applyResultActions_(thread, email, result, tz);

    addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.PROCESSED);

    appendLogRow_(email, result, exec);
    setLastProcessedMessageId_(threadId, latestMessageId);

    if (CONFIG.PROCESSING.MARK_READ_AFTER_PROCESSING) thread.markRead();
    if (CONFIG.PROCESSING.ARCHIVE_AFTER_PROCESSING) thread.moveToArchive();

    Logger.log(`✅ 已处理：${email.subject}`);
  } catch (e) {
    const err = (e && e.stack) ? e.stack : String(e);
    Logger.log(`❌ 处理失败：${err}`);

    if (thread) {
      try {
        addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.ERROR);
        addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.REVIEW);
      } catch (_) {}
    }
  } finally {
    if (thread) {
      try {
        removeLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.PROCESSING);
      } catch (_) {}
    }
    lock.releaseLock();
  }
}

/**
 * =========================
 * 每日回顾（额外一次 AI 调用）
 * =========================
 */
function sendDailyReport() {
  if (!CONFIG.DAILY_REPORT.ENABLED) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const tz = Session.getScriptTimeZone();
    const today = new Date();
    const dateStr = Utilities.formatDate(today, tz, 'yyyy/MM/dd EEEE');

    const items = getTodayLogEntries_();
    if (items.length === 0) {
      Logger.log('今日无处理邮件，跳过日报');
      return;
    }

    const summary = callGeminiForDailySummary_(items, dateStr, tz);

    const subject = `📊 [智能日报] ${dateStr}（${items.length}封已处理）`;
    GmailApp.sendEmail(CONFIG.DAILY_REPORT.RECIPIENT_EMAIL, subject, summary.plainText, {
      htmlBody: summary.htmlBody,
    });

    Logger.log('✅ 智能日报已发送');
  } catch (e) {
    Logger.log(`❌ 日报失败：${(e && e.stack) ? e.stack : String(e)}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * =========================
 * 数据提取层（代码负责）
 * =========================
 */
function extractEmailData_(thread, messages, timeZone) {
  const latest = messages[messages.length - 1];

  const subject = latest.getSubject() || '';
  const from = latest.getFrom() || '';
  const receivedAt = latest.getDate();
  const messageId = latest.getId();
  const threadId = thread.getId();
  const permalink = thread.getPermalink();

  const plainRaw = latest.getPlainBody() || '';
  const htmlRaw = latest.getBody() || '';

  const cleanedBody = cleanPlainBody_(plainRaw);

  const extractedLinks = extractLinks_(htmlRaw, plainRaw, CONFIG.PROCESSING.MAX_LINKS);
  const extractedPatterns = extractKeyPatterns_(cleanedBody);

  const media = collectMediaParts_(latest);
  const conversationContext = extractConversationContext_(messages.slice(0, -1), timeZone);

  return {
    threadId,
    messageId,
    subject,
    from,
    receivedAt,
    receivedAtStr: Utilities.formatDate(receivedAt, timeZone, 'yyyy-MM-dd HH:mm:ss'),

    body: cleanedBody.substring(0, CONFIG.PROCESSING.MAX_BODY_CHARS),
    htmlSnippet: (htmlRaw || '').substring(0, CONFIG.PROCESSING.MAX_HTML_SNIPPET_CHARS),

    extractedLinks,
    extractedPatterns,

    mediaParts: media.parts, // 给 Gemini 的 parts（含 descriptor + inlineData）
    mediaManifest: media.manifest, // 给 prompt 的“顺序说明”
    otherAttachments: media.otherAttachments,

    conversationContext,

    threadMessageCount: messages.length,
    permalink,
  };
}

function cleanPlainBody_(text) {
  if (!text) return '';
  let t = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  t = t.replace(/[\u200B-\u200D\uFEFF]/g, '');

  const cutMarkers = [
    /^-{2,}\s*Original Message\s*-{2,}$/gmi,
    /^-{2,}\s*Forwarded message\s*-{2,}$/gmi,
    /^Begin forwarded message:/gmi,
    /^_{10,}$/gm,
    /^On .+ wrote:$/gmi,
    /^\s*(From|Sent|To|Cc|Subject)\s*[:：]/gmi,
    /^\s*(发件人|发送时间|收件人|抄送|主题)\s*[:：]/gmi,
  ];

  let cutIndex = null;
  for (const re of cutMarkers) {
    re.lastIndex = 0;
    const m = re.exec(t);
    if (m && m.index > 300) cutIndex = cutIndex === null ? m.index : Math.min(cutIndex, m.index);
  }
  if (cutIndex !== null) t = t.substring(0, cutIndex);

  const sigMarkers = [
    /\n--\s*\n/,
    /\nSent from my (iPhone|iPad|Android).*/i,
    /\n此邮件为自动(发送|生成).*/i,
  ];
  let sigCut = null;
  for (const re of sigMarkers) {
    const m = re.exec(t);
    if (m && m.index > 200) sigCut = sigCut === null ? m.index : Math.min(sigCut, m.index);
  }
  if (sigCut !== null) t = t.substring(0, sigCut);

  const lines = t.split('\n').map(l => l.replace(/[ \t]+$/g, ''));
  const out = [];
  let blank = 0;

  for (const line of lines) {
    if (line.trim() === '') {
      blank += 1;
      if (blank <= 2) out.push('');
    } else {
      blank = 0;
      out.push(line);
    }
  }

  return out.join('\n').trim();
}

function extractLinks_(html, plain, maxLinks) {
  const results = [];
  const seen = new Set();

  function add(url, text, source) {
    const cleaned = normalizeUrl_(url);
    if (!cleaned || !isValidUrl_(cleaned)) return;
    const dedupKey = cleaned;
    if (seen.has(dedupKey)) return;
    seen.add(dedupKey);

    const anchorText = (text || '').trim();
    const domain = extractDomain_(cleaned);
    const type = classifyLinkType_(cleaned, anchorText);
    const score = scoreLink_(cleaned, anchorText, type);

    results.push({
      url: cleaned,
      text: anchorText || domain,
      domain,
      type,
      score,
      source,
    });
  }

  // HTML anchor links
  const anchorRe = /<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = anchorRe.exec(html || '')) !== null && results.length < maxLinks * 3) {
    const url = m[1];
    const rawText = stripHtmlTags_(m[2] || '');
    const text = decodeHtmlEntities_(rawText).trim();
    add(url, text, 'HTML_ANCHOR');
  }

  // Bare URLs (plain + html)
  const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi;
  const combined = `${plain || ''}\n${stripHtmlTags_(html || '')}`;
  while ((m = urlRe.exec(combined)) !== null && results.length < maxLinks * 4) {
    add(m[0], '', 'BARE_URL');
  }

  results.sort((a, b) => b.score - a.score);

  const kept = [];
  for (const l of results) {
    if (kept.length >= maxLinks) break;
    kept.push(l);
  }
  return kept;
}

function normalizeUrl_(url) {
  if (!url) return '';
  let u = String(url).trim();
  u = u.replace(/^<|>$/g, '');
  u = u.replace(/[)\].,;:!?"']+$/g, ''); // 末尾标点
  return u;
}

function isValidUrl_(url) {
  if (!url) return false;
  const u = String(url).trim();
  if (!/^https?:\/\//i.test(u)) return false;
  if (/^https?:\/\/(localhost|127\.0\.0\.1)/i.test(u)) return false;
  if (/javascript:/i.test(u)) return false;
  return true;
}

function extractDomain_(url) {
  try {
    const m = String(url).match(/^https?:\/\/([^\/?#]+)/i);
    return m ? m[1] : url.substring(0, 40);
  } catch (_) {
    return String(url).substring(0, 40);
  }
}

function classifyLinkType_(url, text) {
  const u = String(url).toLowerCase();
  const t = String(text || '').toLowerCase();

  if (u.includes('unsubscribe') || t.includes('取消订阅') || t.includes('unsubscribe')) return 'UNSUBSCRIBE';
  if (u.includes('zoom.') || u.includes('teams.microsoft') || u.includes('meet.google') || u.includes('webex') || u.includes('tencentmeeting') || t.includes('zoom') || t.includes('teams')) {
    return 'MEETING';
  }
  if (u.includes('calendar') || u.includes('calendly') || u.endsWith('.ics') || t.includes('add to calendar') || t.includes('日程')) return 'CALENDAR';
  if (u.includes('pay') || u.includes('invoice') || u.includes('billing') || u.includes('stripe') || u.includes('paypal') || t.includes('付款') || t.includes('支付') || t.includes('invoice')) return 'PAYMENT';
  if (u.includes('docs.google') || u.includes('drive.google') || u.includes('dropbox') || u.includes('onedrive') || u.includes('sharepoint') || u.includes('notion.so') || u.includes('figma.com')) return 'DOCUMENT';
  if (u.includes('register') || u.includes('signup') || u.includes('registration') || u.includes('rsvp') || u.includes('eventbrite') || t.includes('register') || t.includes('报名') || t.includes('注册')) return 'REGISTRATION';
  if (u.includes('login') || u.includes('verify') || u.includes('otp') || u.includes('reset') || t.includes('verify') || t.includes('验证码')) return 'AUTH';
  if (u.includes('facebook.com') || u.includes('twitter.com') || u.includes('instagram.com') || u.includes('linkedin.com') || u.includes('youtube.com') || u.includes('tiktok.com')) return 'SOCIAL';

  return 'GENERAL';
}

function scoreLink_(url, text, type) {
  const u = String(url).toLowerCase();
  const t = String(text || '').toLowerCase();

  const baseByType = {
    MEETING: 100,
    PAYMENT: 95,
    AUTH: 90,
    DOCUMENT: 85,
    CALENDAR: 80,
    REGISTRATION: 70,
    GENERAL: 30,
    SOCIAL: 5,
    UNSUBSCRIBE: -50,
  };

  let score = baseByType[type] != null ? baseByType[type] : 20;

  const goodKw = ['join', 'register', 'rsvp', 'ticket', 'invoice', 'pay', 'verify', 'confirm', 'download', 'open', 'view', '报名', '注册', '加入', '验证', '确认', '查看'];
  for (const kw of goodKw) {
    if (t.includes(kw) || u.includes(kw)) {
      score += 6;
      break;
    }
  }

  if (u.includes('utm_') || u.includes('utm-') || u.includes('tracking') || u.includes('trk=')) score -= 20;
  if (u.includes('mailto:') || u.includes('tel:')) score -= 100;

  if (text && text.trim().length >= 8) score += 5;

  return score;
}

function stripHtmlTags_(html) {
  return String(html || '').replace(/<[^>]*>/g, ' ');
}

function decodeHtmlEntities_(text) {
  return String(text || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function extractKeyPatterns_(text) {
  const t = String(text || '');
  const patterns = {
    referenceNumbers: [],
    codes: [],
    dates: [],
    times: [],
    amounts: [],
    emails: [],
    phones: [],
  };

  const pushUniq = (arr, v) => {
    const val = String(v || '').trim();
    if (!val) return;
    if (arr.indexOf(val) === -1) arr.push(val);
  };

  const emailRe = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
  let m;
  while ((m = emailRe.exec(t)) !== null) pushUniq(patterns.emails, m[0]);

  const phoneRe = /(\+\d{1,3}[\s-]?)?(\(?\d{2,4}\)?[\s-]?)?\d{3,4}[\s-]?\d{3,4}/g;
  while ((m = phoneRe.exec(t)) !== null) {
    const p = m[0].replace(/\s+/g, ' ').trim();
    if (p.length >= 8) pushUniq(patterns.phones, p);
  }

  const amountRe = /(?:HK\$|\$|¥|£|€)\s*[\d,]+(?:\.\d+)?/g;
  while ((m = amountRe.exec(t)) !== null) pushUniq(patterns.amounts, m[0]);

  const refRe = /(?:(?:order|confirmation|reference|booking|ticket|invoice|case|ref|no|#|订单|确认|预订|发票|单号|编号|参考)\s*[:#：]?\s*)([A-Z0-9][A-Z0-9\-]{5,30})/gi;
  while ((m = refRe.exec(t)) !== null) pushUniq(patterns.referenceNumbers, m[1]);

  const otpRe = /(?:OTP|code|验证码|驗證碼|verification code)\s*[:：]?\s*([0-9]{4,8})/gi;
  while ((m = otpRe.exec(t)) !== null) pushUniq(patterns.codes, m[1]);

  const codeRe = /\b([A-Z]{2,5}[0-9]{4,12})\b/g;
  while ((m = codeRe.exec(t)) !== null) pushUniq(patterns.codes, m[1]);

  const dateRe = /\b(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/g;
  while ((m = dateRe.exec(t)) !== null) pushUniq(patterns.dates, m[1]);

  const timeRe = /\b(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\b/g;
  while ((m = timeRe.exec(t)) !== null) pushUniq(patterns.times, m[1]);

  return patterns;
}

function extractConversationContext_(previousMessages, timeZone) {
  const msgs = previousMessages || [];
  if (msgs.length === 0) return null;

  const n = Math.min(CONFIG.PROCESSING.MAX_CONVERSATION_MESSAGES, msgs.length);
  const last = msgs.slice(-n);

  return last.map(msg => {
    const body = cleanPlainBody_(msg.getPlainBody() || '').substring(0, CONFIG.PROCESSING.MAX_CONVERSATION_SNIPPET_CHARS);
    return {
      from: msg.getFrom() || '',
      date: Utilities.formatDate(msg.getDate(), timeZone, 'yyyy-MM-dd HH:mm:ss'),
      snippet: body,
    };
  });
}

function collectMediaParts_(message) {
  const attachments = message.getAttachments({ includeInlineImages: true, includeAttachments: true }) || [];

  const imageCandidates = [];
  const pdfCandidates = [];
  const otherAttachments = [];

  for (const att of attachments) {
    const mimeType = att.getContentType() || '';
    const name = att.getName() || '';
    const size = att.getSize() || 0;

    const base = { att, mimeType, name, size };

    if (mimeType.startsWith('image/')) {
      const score = scoreImageAttachment_(name, mimeType, size);
      imageCandidates.push({ ...base, score, kind: 'image' });
    } else if (mimeType === 'application/pdf' || /\.pdf$/i.test(name)) {
      const score = scorePdfAttachment_(name, size);
      pdfCandidates.push({ ...base, score, kind: 'pdf' });
    } else {
      otherAttachments.push({
        filename: name || '(unnamed)',
        mimeType: mimeType || 'unknown',
        sizeKB: Math.round(size / 1024),
      });
    }
  }

  imageCandidates.sort((a, b) => b.score - a.score);
  pdfCandidates.sort((a, b) => b.score - a.score);

  const selected = [];
  const maxImages = CONFIG.PROCESSING.MAX_IMAGE_ITEMS;
  const maxPdfs = CONFIG.PROCESSING.MAX_PDF_ITEMS;

  for (const c of imageCandidates) {
    if (selected.filter(x => x.kind === 'image').length >= maxImages) break;
    if (c.size < CONFIG.PROCESSING.MIN_IMAGE_BYTES) continue;
    if (c.size > CONFIG.PROCESSING.MAX_MEDIA_BYTES_EACH) continue;
    selected.push(c);
    if (selected.length >= CONFIG.PROCESSING.MAX_MEDIA_ITEMS) break;
  }

  for (const c of pdfCandidates) {
    if (selected.filter(x => x.kind === 'pdf').length >= maxPdfs) break;
    if (c.size > CONFIG.PROCESSING.MAX_MEDIA_BYTES_EACH) continue;
    selected.push(c);
    if (selected.length >= CONFIG.PROCESSING.MAX_MEDIA_ITEMS) break;
  }

  let totalBytes = 0;
  const parts = [];
  const manifest = [];

  let index = 0;
  for (const item of selected) {
    if (totalBytes + item.size > CONFIG.PROCESSING.MAX_TOTAL_MEDIA_BYTES) break;

    index += 1;
    totalBytes += item.size;

    const desc = {
      index,
      kind: item.kind,
      filename: item.name || '(unnamed)',
      mimeType: item.mimeType,
      sizeKB: Math.round(item.size / 1024),
      score: item.score,
    };
    manifest.push(desc);

    parts.push({
      text: `MEDIA #${index}: kind=${desc.kind}, filename=${desc.filename}, mimeType=${desc.mimeType}, sizeKB=${desc.sizeKB}`,
    });

    parts.push({
      inlineData: {
        mimeType: item.mimeType,
        data: Utilities.base64Encode(item.att.getBytes()),
      },
    });
  }

  return { parts, manifest, otherAttachments };
}

function scoreImageAttachment_(name, mimeType, size) {
  const n = String(name || '').toLowerCase();
  let score = 30;

  const strong = ['qr', 'qrcode', 'barcode', 'ticket', 'pass', 'boarding', 'receipt', 'invoice', 'bill', 'payment', '二维码', '票', '凭证', '发票', '收据', '付款'];
  for (const kw of strong) if (n.includes(kw)) score += 80;

  const noise = ['logo', 'icon', 'banner', 'facebook', 'twitter', 'linkedin', 'instagram', 'wechat', 'whatsapp'];
  for (const kw of noise) if (n.includes(kw)) score -= 60;

  if (mimeType === 'image/png') score += 10;
  if (size >= 200 * 1024 && size <= 1500 * 1024) score += 20;
  if (size < 20 * 1024) score -= 40;

  return score;
}

function scorePdfAttachment_(name, size) {
  const n = String(name || '').toLowerCase();
  let score = 50;

  const strong = ['invoice', 'receipt', 'ticket', 'boarding', 'statement', 'bill', '发票', '收据', '票', '账单', '对账单', '付款'];
  for (const kw of strong) if (n.includes(kw)) score += 80;

  if (size >= 200 * 1024 && size <= 5 * 1024 * 1024) score += 20;
  return score;
}

/**
 * =========================
 * AI 调用层（单封邮件）
 * =========================
 */
function callGeminiForEmail_(email, timeZone, now) {
  const prompt = buildEmailPrompt_(email, timeZone, now);
  const parts = [{ text: prompt }];

  if (email.mediaParts && email.mediaParts.length > 0) {
    parts.push(...email.mediaParts);
  }

  const genCfg = {
    responseMimeType: 'application/json',
    temperature: CONFIG.GEMINI.TEMPERATURE,
    thinkingConfig: { thinkingLevel: CONFIG.GEMINI.THINKING_LEVEL },
  };

  if (CONFIG.GEMINI.MEDIA_RESOLUTION) genCfg.mediaResolution = CONFIG.GEMINI.MEDIA_RESOLUTION;
  if (CONFIG.GEMINI.USE_RESPONSE_SCHEMA) genCfg.responseSchema = getEmailResponseSchema_();

  return fetchGeminiJson_(parts, genCfg);
}

function buildEmailPrompt_(email, timeZone, now) {
  const linksText = (email.extractedLinks || [])
    .map(l => `- [${l.type}] score=${l.score} text="${l.text}" url=${l.url}`)
    .join('\n');

  const patterns = email.extractedPatterns || {};
  const mediaManifest = email.mediaManifest || [];

  return `
SYSTEM ROLE (strict):
You are an ultra-intelligent executive assistant. You MUST deeply understand the email (and any attached images/PDFs), then make decisions.
Output MUST be valid JSON that matches the provided schema. Do NOT output any extra text.

Context:
- Current time: ${Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm:ss')}
- Timezone: ${timeZone}

Email (latest message in thread):
- Subject: ${email.subject}
- From: ${email.from}
- ReceivedAt: ${email.receivedAtStr}
- Thread messages: ${email.threadMessageCount}
- Gmail permalink: ${email.permalink}

Cleaned plain body (high-signal, may still contain forwarded content):
${email.body}

HTML snippet (for layout/table hints; may be noisy):
${email.htmlSnippet}

Extracted links (already de-duplicated and pre-scored):
${linksText || '(none)'}

Extracted patterns (pre-heuristics; do NOT trust blindly):
${JSON.stringify(patterns, null, 2)}

Conversation context (previous messages, if any):
${email.conversationContext ? JSON.stringify(email.conversationContext, null, 2) : '(none)'}

Attached media manifest (the actual media parts follow this text, in the same order):
${mediaManifest.length ? JSON.stringify(mediaManifest, null, 2) : '(none)'}
Other attachments (not included as media):
${(email.otherAttachments && email.otherAttachments.length) ? JSON.stringify(email.otherAttachments, null, 2) : '(none)'}

Your mission (high thinking):
1) Understand the TRUE intent and what the recipient should do.
2) Categorize the email into one of:
   - EVENT: specific date+time (meeting/webinar/flight/appointment) and should go to calendar.
   - TASK: action required but no fixed time slot.
   - REMINDER: important to remember (status/confirmation) but low action.
   - PROMO: marketing/newsletter.
   - INFO: informational only.
3) Decide priority: HIGH / MEDIUM / LOW.
4) Be selective with links:
   - Choose up to 3 truly relevant links, ordered by importance.
   - Prefer actionable links (join/confirm/pay/verify/download) and ignore tracking/unsubscribe unless it's actually needed.
5) Analyze images/PDFs:
   - If there is a QR code, try to read it; if not possible, describe its likely purpose precisely.
   - Extract important numbers, reference IDs, tickets, amounts, schedules, dates/times.
6) Provide concrete next steps (specific, actionable).

Time rules:
- Use RFC3339 with timezone offset, e.g. 2026-01-11T09:30:00+08:00
- If only a date is known, set deadline to 17:00 local time for that date.
- If timezone is mentioned in email, respect it; otherwise default to ${timeZone}.
- If ambiguous, set requiresReview=true. Prefer TASK over INFO to avoid missing.

Output language:
- reasoning / understanding / description / nextSteps: Chinese (简洁、可执行)
- title: Chinese or English ok, but short

Remember: JSON only.`;
}

function getEmailResponseSchema_() {
  return {
    type: 'object',
    properties: {
      category: { type: 'string', enum: ['EVENT', 'TASK', 'REMINDER', 'PROMO', 'INFO'] },
      priority: { type: 'string', enum: ['HIGH', 'MEDIUM', 'LOW'] },
      confidence: { type: 'number', description: '0.0-1.0' },
      requiresReview: { type: 'boolean' },

      reasoning: { type: 'string' },
      understanding: { type: 'string' },

      traits: {
        type: 'object',
        properties: {
          hasQr: { type: 'boolean' },
          hasMeetingLink: { type: 'boolean' },
          hasPayment: { type: 'boolean' },
          hasCode: { type: 'boolean' },
          hasAttachment: { type: 'boolean' },
        },
      },

      meta: {
        type: 'object',
        properties: {
          title: { type: 'string' },

          startTime: { type: 'string', nullable: true, description: 'RFC3339 with offset' },
          endTime: { type: 'string', nullable: true, description: 'RFC3339 with offset' },
          deadline: { type: 'string', nullable: true, description: 'RFC3339 with offset (task deadline)' },
          allDay: { type: 'boolean' },
          location: { type: 'string', nullable: true },

          description: { type: 'string' },
          actionRequired: { type: 'string', nullable: true },

          nextSteps: {
            type: 'array',
            items: { type: 'string' },
            maxItems: 6,
          },

          keyLinks: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                url: { type: 'string' },
                purpose: { type: 'string' },
              },
              required: ['url', 'purpose'],
            },
            maxItems: 3,
          },

          keyNumbers: {
            type: 'array',
            items: { type: 'string' },
            maxItems: 12,
          },

          imageInsights: { type: 'string', nullable: true },
          qrContent: { type: 'string', nullable: true, description: 'decoded QR content if possible' },
        },
        required: ['title', 'allDay', 'description', 'nextSteps', 'keyLinks', 'keyNumbers'],
      },

      suggestedLabels: {
        type: 'array',
        items: { type: 'string' },
        maxItems: 6,
      },
    },
    required: ['category', 'priority', 'confidence', 'requiresReview', 'reasoning', 'understanding', 'meta'],
  };
}

function fetchGeminiJson_(parts, generationConfig) {
  const apiKey = getApiKey_();
  const url = `${CONFIG.GEMINI.API_ROOT}/${encodeURIComponent(CONFIG.GEMINI.MODEL_NAME)}:generateContent`;

  const payload = {
    contents: [{ role: 'user', parts }],
    generationConfig,
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects: true,
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  let json;
  try {
    json = JSON.parse(text);
  } catch (_) {
    throw new Error(`Gemini 返回非 JSON（HTTP ${code}）：${text.slice(0, 800)}`);
  }

  if (json.error) throw new Error(json.error.message || JSON.stringify(json.error));
  if (code < 200 || code >= 300) throw new Error(`Gemini HTTP ${code}：${text.slice(0, 800)}`);

  const outText = json?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!outText) throw new Error('Gemini 返回缺少 candidates[0].content.parts[0].text');

  try {
    return JSON.parse(outText);
  } catch (e) {
    throw new Error(`Gemini 输出无法 JSON.parse：${String(e)}\n原始输出(前800字)：${outText.slice(0, 800)}`);
  }
}

function getApiKey_() {
  const fromConfig = String(CONFIG.GEMINI.API_KEY || '').trim();
  if (fromConfig && fromConfig !== 'YOUR_API_KEY_HERE') return fromConfig;

  const fromProps = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (fromProps && fromProps.trim()) return fromProps.trim();

  throw new Error('未配置 Gemini API Key：请运行 setGeminiApiKey("...") 或在 CONFIG.GEMINI.API_KEY 写入');
}

/**
 * =========================
 * AI 结果归一化 + 执行层
 * =========================
 */
function normalizeAiResult_(raw, email) {
  const allowedCategory = { EVENT: true, TASK: true, REMINDER: true, PROMO: true, INFO: true };
  const allowedPriority = { HIGH: true, MEDIUM: true, LOW: true };

  const r = raw && typeof raw === 'object' ? raw : {};
  const meta = r.meta && typeof r.meta === 'object' ? r.meta : {};

  const category = allowedCategory[String(r.category || '').toUpperCase()] ? String(r.category).toUpperCase() : 'TASK';
  const priority = allowedPriority[String(r.priority || '').toUpperCase()] ? String(r.priority).toUpperCase() : 'MEDIUM';

  let confidence = Number(r.confidence);
  if (Number.isNaN(confidence)) confidence = 0.6;
  confidence = Math.max(0, Math.min(1, confidence));

  const requiresReview = typeof r.requiresReview === 'boolean' ? r.requiresReview : (confidence < 0.55);

  const nextSteps = Array.isArray(meta.nextSteps) ? meta.nextSteps.filter(Boolean).slice(0, 6) : [];
  const keyLinks = Array.isArray(meta.keyLinks) ? meta.keyLinks.slice(0, 3) : [];
  const keyNumbers = Array.isArray(meta.keyNumbers) ? meta.keyNumbers.filter(Boolean).slice(0, 12) : [];

  const titleFallback = (email.subject || '').substring(0, 80) || '邮件事项';

  return {
    category,
    priority,
    confidence,
    requiresReview,

    reasoning: String(r.reasoning || '').trim().substring(0, 500),
    understanding: String(r.understanding || '').trim().substring(0, 300),

    traits: r.traits && typeof r.traits === 'object' ? r.traits : {},

    meta: {
      title: String(meta.title || titleFallback).trim().substring(0, 80),
      startTime: meta.startTime ? String(meta.startTime).trim() : null,
      endTime: meta.endTime ? String(meta.endTime).trim() : null,
      deadline: meta.deadline ? String(meta.deadline).trim() : null,
      allDay: typeof meta.allDay === 'boolean' ? meta.allDay : false,
      location: meta.location ? String(meta.location).trim().substring(0, 200) : null,

      description: String(meta.description || '').trim().substring(0, 5000),
      actionRequired: meta.actionRequired ? String(meta.actionRequired).trim().substring(0, 300) : null,

      nextSteps,
      keyLinks,
      keyNumbers,

      imageInsights: meta.imageInsights ? String(meta.imageInsights).trim().substring(0, 2000) : null,
      qrContent: meta.qrContent ? String(meta.qrContent).trim().substring(0, 500) : null,
    },

    suggestedLabels: Array.isArray(r.suggestedLabels) ? r.suggestedLabels.filter(Boolean).slice(0, 6) : [],
  };
}

function applyResultActions_(thread, email, result, timeZone) {
  applySmartLabels_(thread, email, result, { action: 'NONE', calendarEventId: '', taskId: '' });

  const notes = buildNotesFromResult_(result, email);
  const exec = { action: 'NONE', calendarEventId: '', taskId: '', notesPreview: notes.substring(0, 800) };

  if (CONFIG.PROCESSING.DRY_RUN) return exec;

  const shouldForceTask = result.requiresReview && CONFIG.PROCESSING.ALWAYS_CREATE_TASK_WHEN_REQUIRES_REVIEW;

  if (result.category === 'EVENT' && !shouldForceTask) {
    const start = parseDateTimeSafe_(result.meta.startTime);
    if (start && result.confidence >= CONFIG.PROCESSING.MIN_CONFIDENCE_TO_AUTO_CREATE_EVENT) {
      const end = parseDateTimeSafe_(result.meta.endTime) || new Date(start.getTime() + 60 * 60 * 1000);

      try {
        const eventId = createCalendarEvent_(result, start, end, notes);
        exec.action = 'CALENDAR';
        exec.calendarEventId = eventId || '';
        addLabel_(thread, CONFIG.GMAIL.SYNC_LABELS.CALENDAR);
        return exec;
      } catch (e) {
        Logger.log(`⚠️ 日历创建失败，降级为 Task：${String(e)}`);
      }
    }
  }

  const shouldCreateTask =
    shouldForceTask ||
    result.category === 'TASK' ||
    (result.category === 'REMINDER' && result.priority !== 'LOW') ||
    (result.category === 'EVENT' && (!result.meta.startTime || result.confidence < CONFIG.PROCESSING.MIN_CONFIDENCE_TO_AUTO_CREATE_EVENT));

  if (shouldCreateTask) {
    try {
      const taskId = createGoogleTask_(result, notes);
      exec.action = 'TASKS';
      exec.taskId = taskId || '';
      addLabel_(thread, CONFIG.GMAIL.SYNC_LABELS.TASKS);
      return exec;
    } catch (e) {
      Logger.log(`⚠️ Tasks 写入失败：${String(e)}`);
      addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.REVIEW);
    }
  }

  return exec;
}

function buildNotesFromResult_(result, email) {
  const m = result.meta || {};
  const lines = [];

  lines.push(`📌 解析理解：${result.understanding || '(无)'}`);
  lines.push(`🎯 分类/优先级：${result.category} / ${result.priority}（置信度 ${result.confidence}）`);
  if (result.reasoning) lines.push(`🧠 理由：${result.reasoning}`);

  lines.push('');
  if (m.description) {
    lines.push('📝 摘要：');
    lines.push(m.description);
    lines.push('');
  }

  if (m.actionRequired) {
    lines.push(`⚡ 需要行动：${m.actionRequired}`);
    lines.push('');
  }

  if (m.startTime || m.deadline) {
    lines.push(`⏰ 时间信息：startTime=${m.startTime || 'null'} endTime=${m.endTime || 'null'} deadline=${m.deadline || 'null'} allDay=${String(m.allDay)}`);
  }

  if (m.location) lines.push(`📍 地点/链接：${m.location}`);

  if (m.keyNumbers && m.keyNumbers.length) lines.push(`🔢 关键编号：${m.keyNumbers.join(' | ')}`);

  if (m.qrContent) {
    lines.push(`🧾 二维码内容：${m.qrContent}`);
  } else if (m.imageInsights) {
    lines.push(`🖼️ 图片/PDF要点：${m.imageInsights}`);
  }

  if (m.keyLinks && m.keyLinks.length) {
    lines.push('');
    lines.push('🔗 重要链接：');
    for (const l of m.keyLinks) {
      if (!l || !l.url) continue;
      lines.push(`- ${l.purpose || 'link'}: ${l.url}`);
    }
  }

  if (m.nextSteps && m.nextSteps.length) {
    lines.push('');
    lines.push('📋 建议下一步：');
    for (const step of m.nextSteps) lines.push(`□ ${step}`);
  }

  lines.push('');
  lines.push(`📧 原邮件：${email.permalink}`);
  lines.push(`👤 发件人：${email.from}`);
  lines.push(`🗓️ 收到时间：${email.receivedAtStr}`);
  return lines.join('\n');
}

function parseDateTimeSafe_(value) {
  if (!value) return null;
  const s = String(value).trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(`${s}T09:00:00`);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

function createCalendarEvent_(result, start, end, notes) {
  const cal = CalendarApp.getDefaultCalendar();
  const title = decorateTitle_(result.meta.title, result.priority, result.requiresReview);

  if (result.meta.allDay) {
    const startDate = new Date(start);
    startDate.setHours(0, 0, 0, 0);

    const endDate = new Date(end);
    endDate.setHours(0, 0, 0, 0);

    if (endDate.getTime() <= startDate.getTime()) endDate.setDate(endDate.getDate() + 1);

    const ev = cal.createAllDayEvent(title, startDate, endDate, {
      location: result.meta.location || '',
      description: notes,
    });
    addEventReminders_(ev, result.priority);
    return ev.getId();
  }

  const ev = cal.createEvent(title, start, end, {
    location: result.meta.location || '',
    description: notes,
  });
  addEventReminders_(ev, result.priority);
  return ev.getId();
}

function addEventReminders_(event, priority) {
  try {
    event.removeAllReminders();
  } catch (_) {}

  if (priority === 'HIGH') {
    event.addPopupReminder(1440);
    event.addPopupReminder(60);
  } else if (priority === 'MEDIUM') {
    event.addPopupReminder(120);
    event.addPopupReminder(30);
  } else {
    event.addPopupReminder(30);
  }
}

function isTasksServiceAvailable_() {
  return typeof Tasks !== 'undefined' && Tasks.Tasks && typeof Tasks.Tasks.insert === 'function';
}

function createGoogleTask_(result, notes) {
  if (!isTasksServiceAvailable_()) throw new Error('未启用 Advanced Google services: Tasks API');

  const title = decorateTitle_(result.meta.title, result.priority, result.requiresReview);

  const task = {
    title,
    notes: String(notes || '').substring(0, 8000),
  };

  const due = normalizeTaskDueRfc3339_(result.meta.deadline || result.meta.startTime);
  if (due) task.due = due;

  const inserted = Tasks.Tasks.insert(task, '@default');
  return inserted && inserted.id ? inserted.id : '';
}

function normalizeTaskDueRfc3339_(value) {
  if (!value) return '';
  const s = String(value).trim();
  if (!s) return '';

  if (s.includes('T')) {
    const d = new Date(s);
    return Number.isNaN(d.getTime()) ? '' : d.toISOString();
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(`${s}T09:00:00`);
    return Number.isNaN(d.getTime()) ? '' : d.toISOString();
  }

  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? '' : d.toISOString();
}

function decorateTitle_(title, priority, requiresReview) {
  let t = String(title || '').trim();
  if (!t) t = '邮件事项';

  const p = priority === 'HIGH' ? '🔴' : priority === 'MEDIUM' ? '🟡' : '🟢';
  const r = requiresReview ? '⚠️' : '';
  return `${p}${r} ${t}`.substring(0, 100);
}

/**
 * =========================
 * 规则跳过（可选，省 API 次数）
 * =========================
 */
function maybeBypassAi_(email) {
  if (!CONFIG.PROCESSING.ENABLE_HEURISTIC_PROMO_BYPASS) return { bypassed: false };

  const subj = (email.subject || '').toLowerCase();
  const body = (email.body || '').toLowerCase();
  const from = (email.from || '').toLowerCase();
  const links = email.extractedLinks || [];

  let score = 0;
  if (from.includes('noreply') || from.includes('no-reply')) score += 1;
  if (subj.includes('newsletter') || subj.includes('unsubscribe') || subj.includes('promo') || subj.includes('sale') || subj.includes('offer')) score += 2;
  if (body.includes('unsubscribe') || body.includes('manage preferences') || body.includes('取消订阅') || body.includes('退订')) score += 2;
  if (links.filter(l => l.type === 'UNSUBSCRIBE').length >= 1) score += 2;
  if (links.length >= 18) score += 1;

  const prob = Math.min(1, score / 8);
  if (prob < CONFIG.PROCESSING.PROMO_BYPASS_THRESHOLD) return { bypassed: false };

  const result = {
    category: 'PROMO',
    priority: 'LOW',
    confidence: prob,
    requiresReview: false,
    reasoning: '规则判定为明显推广/Newsletter，跳过AI以节省请求次数',
    understanding: '推广或订阅类内容（非关键事务）',
    traits: { hasAttachment: (email.mediaManifest && email.mediaManifest.length > 0) },
    meta: {
      title: (email.subject || '推广邮件').substring(0, 80),
      startTime: null,
      endTime: null,
      deadline: null,
      allDay: false,
      location: null,
      description: '该邮件被规则快速判定为推广/Newsletter。如需更精细处理，可关闭 ENABLE_HEURISTIC_PROMO_BYPASS。',
      actionRequired: null,
      nextSteps: [],
      keyLinks: [],
      keyNumbers: [],
      imageInsights: null,
      qrContent: null,
    },
    suggestedLabels: ['newsletter'],
  };

  return { bypassed: true, result };
}

/**
 * =========================
 * Gmail 标签体系
 * =========================
 */
function ensureLabelsExist_() {
  getOrCreateLabel_(CONFIG.GMAIL.AI_ROOT_LABEL);

  Object.values(CONFIG.GMAIL.STATUS_LABELS).forEach(getOrCreateLabel_);
  Object.values(CONFIG.GMAIL.CATEGORY_LABELS).forEach(getOrCreateLabel_);
  Object.values(CONFIG.GMAIL.PRIORITY_LABELS).forEach(getOrCreateLabel_);
  Object.values(CONFIG.GMAIL.SYNC_LABELS).forEach(getOrCreateLabel_);
  Object.values(CONFIG.GMAIL.TRAIT_LABELS).forEach(getOrCreateLabel_);
}

function getOrCreateLabel_(name) {
  if (CACHE.labels[name]) return CACHE.labels[name];

  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);

  CACHE.labels[name] = label;
  return label;
}

function addLabel_(thread, name) {
  const label = getOrCreateLabel_(name);
  thread.addLabel(label);
}

function removeLabel_(thread, name) {
  const label = GmailApp.getUserLabelByName(name);
  if (label) thread.removeLabel(label);
}

function clearAiDerivedLabels_(thread) {
  const toRemove = new Set([
    ...Object.values(CONFIG.GMAIL.STATUS_LABELS).filter(n => n !== CONFIG.GMAIL.STATUS_LABELS.PROCESSING),
    ...Object.values(CONFIG.GMAIL.CATEGORY_LABELS),
    ...Object.values(CONFIG.GMAIL.PRIORITY_LABELS),
    ...Object.values(CONFIG.GMAIL.SYNC_LABELS),
    ...Object.values(CONFIG.GMAIL.TRAIT_LABELS),
  ]);

  const labels = thread.getLabels();
  for (const l of labels) {
    const name = l.getName();
    if (toRemove.has(name)) {
      thread.removeLabel(l);
      continue;
    }
    if (name.startsWith(CONFIG.GMAIL.EXTRA_LABEL_PREFIX)) {
      thread.removeLabel(l);
    }
  }
}

function applySmartLabels_(thread, email, result) {
  const catLabel = CONFIG.GMAIL.CATEGORY_LABELS[result.category];
  const priLabel = CONFIG.GMAIL.PRIORITY_LABELS[result.priority];

  if (catLabel) addLabel_(thread, catLabel);
  if (priLabel) addLabel_(thread, priLabel);

  if (result.requiresReview || result.confidence < 0.55) addLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.REVIEW);

  const traits = result.traits || {};
  const patterns = email.extractedPatterns || {};
  const hasCode = Boolean(traits.hasCode) || (patterns.referenceNumbers && patterns.referenceNumbers.length) || (patterns.codes && patterns.codes.length);
  const hasPayment = Boolean(traits.hasPayment) || (patterns.amounts && patterns.amounts.length);
  const hasMeetingLink = Boolean(traits.hasMeetingLink) || (email.extractedLinks || []).some(l => l.type === 'MEETING');
  const hasAttachment = Boolean(traits.hasAttachment) || (email.mediaManifest && email.mediaManifest.length) || (email.otherAttachments && email.otherAttachments.length);
  const hasQr = Boolean(traits.hasQr) || Boolean(result.meta && (result.meta.qrContent || (result.meta.imageInsights && /qr|二维码/i.test(result.meta.imageInsights))));

  if (hasQr) addLabel_(thread, CONFIG.GMAIL.TRAIT_LABELS.HAS_QR);
  if (hasCode) addLabel_(thread, CONFIG.GMAIL.TRAIT_LABELS.HAS_CODE);
  if (hasMeetingLink) addLabel_(thread, CONFIG.GMAIL.TRAIT_LABELS.HAS_MEETING_LINK);
  if (hasPayment) addLabel_(thread, CONFIG.GMAIL.TRAIT_LABELS.HAS_PAYMENT);
  if (hasAttachment) addLabel_(thread, CONFIG.GMAIL.TRAIT_LABELS.HAS_ATTACHMENT);

  if (Array.isArray(result.suggestedLabels)) {
    const extras = result.suggestedLabels.slice(0, 3).map(sanitizeLabelSegment_).filter(Boolean);
    for (const seg of extras) addLabel_(thread, `${CONFIG.GMAIL.EXTRA_LABEL_PREFIX}${seg}`);
  }
}

function sanitizeLabelSegment_(s) {
  let x = String(s || '').trim();
  if (!x) return '';
  x = x.replace(/[\/\\]/g, '／');
  x = x.replace(/[\[\]]/g, '');
  x = x.replace(/\s+/g, ' ');
  x = x.substring(0, 30);
  return x;
}

/**
 * =========================
 * 线程选择 + 幂等状态
 * =========================
 */
function getNextThreadToProcess_() {
  const source = GmailApp.getUserLabelByName(CONFIG.GMAIL.SOURCE_LABEL);
  if (!source) throw new Error(`找不到来源标签：${CONFIG.GMAIL.SOURCE_LABEL}`);

  const threads = source.getThreads(0, CONFIG.PROCESSING.MAX_THREADS_SCAN) || [];
  for (const thread of threads) {
    if (threadHasLabel_(thread, CONFIG.GMAIL.STATUS_LABELS.PROCESSING)) continue;

    const messages = thread.getMessages();
    if (!messages || messages.length === 0) continue;

    const latestId = messages[messages.length - 1].getId();
    const last = getLastProcessedMessageId_(thread.getId());
    if (last !== latestId) return thread;
  }

  return null;
}

function threadHasLabel_(thread, labelName) {
  const labels = thread.getLabels();
  for (const l of labels) if (l.getName() === labelName) return true;
  return false;
}

function getLastProcessedMessageId_(threadId) {
  return PropertiesService.getScriptProperties().getProperty(`thread:${threadId}:lastMessageId`) || '';
}

function setLastProcessedMessageId_(threadId, messageId) {
  PropertiesService.getScriptProperties().setProperty(`thread:${threadId}:lastMessageId`, String(messageId));
}

/**
 * =========================
 * 日志（Google Sheet）
 * =========================
 */
function ensureLogSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('LOG_SHEET_ID');
  let ss;

  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create(CONFIG.DAILY_REPORT.LOG_SPREADSHEET_NAME);
    ssId = ss.getId();
    props.setProperty('LOG_SHEET_ID', ssId);
  }

  let sheet = ss.getSheetByName(CONFIG.DAILY_REPORT.LOG_SHEET_TAB);
  if (!sheet) {
    const first = ss.getSheets()[0] || ss.insertSheet();
    first.setName(CONFIG.DAILY_REPORT.LOG_SHEET_TAB);
    sheet = first;
  }

  if (sheet.getLastRow() === 0) {
    const header = [
      'processedAt',
      'threadId',
      'messageId',
      'subject',
      'from',
      'receivedAt',
      'category',
      'priority',
      'confidence',
      'requiresReview',
      'actionTaken',
      'startTime',
      'endTime',
      'deadline',
      'calendarEventId',
      'taskId',
      'keyLinks',
      'keyNumbers',
      'hasMedia',
      'permalink',
      'understanding',
      'notesPreview',
      'rawJson',
    ];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }

  CACHE.log.ssId = ssId;
  CACHE.log.sheet = sheet;
  return sheet;
}

function appendLogRow_(email, result, exec) {
  const sheet = ensureLogSheet_();

  const keyLinks = (result.meta.keyLinks || [])
    .map(l => `${l.purpose || 'link'}: ${l.url}`)
    .join(' | ');
  const keyNumbers = (result.meta.keyNumbers || []).join(' | ');
  const hasMedia = Boolean(email.mediaManifest && email.mediaManifest.length);

  const row = [
    new Date(),
    email.threadId,
    email.messageId,
    email.subject,
    email.from,
    email.receivedAt,
    result.category,
    result.priority,
    result.confidence,
    result.requiresReview,
    exec.action || 'NONE',
    result.meta.startTime || '',
    result.meta.endTime || '',
    result.meta.deadline || '',
    exec.calendarEventId || '',
    exec.taskId || '',
    keyLinks,
    keyNumbers,
    hasMedia,
    email.permalink,
    result.understanding || '',
    exec.notesPreview || '',
    JSON.stringify(result).substring(0, 5000),
  ];

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function getTodayLogEntries_() {
  const sheet = ensureLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const startRow = Math.max(2, lastRow - CONFIG.DAILY_REPORT.MAX_ROWS_READ + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayStart = new Date(today);
  todayStart.setHours(0, 0, 0, 0);

  const items = [];
  for (const row of values) {
    const processedAt = row[0]; // Date
    if (!(processedAt instanceof Date)) continue;
    if (processedAt < todayStart) continue;

    items.push({
      processedAt: Utilities.formatDate(processedAt, tz, 'HH:mm'),
      subject: row[3],
      from: row[4],
      category: row[6],
      priority: row[7],
      confidence: row[8],
      requiresReview: row[9],
      actionTaken: row[10],
      startTime: row[11],
      deadline: row[13],
      permalink: row[19],
      understanding: row[20],
    });

    if (items.length >= CONFIG.DAILY_REPORT.MAX_ITEMS_IN_PROMPT) break;
  }

  return items;
}

/**
 * =========================
 * 日报 AI（一次调用）
 * =========================
 */
function callGeminiForDailySummary_(items, dateStr, timeZone) {
  const prompt = buildDailySummaryPrompt_(items, dateStr, timeZone);

  const genCfg = {
    responseMimeType: 'application/json',
    temperature: CONFIG.DAILY_REPORT.TEMPERATURE,
    thinkingConfig: { thinkingLevel: CONFIG.DAILY_REPORT.THINKING_LEVEL },
    responseSchema: getDailySummarySchema_(),
  };

  const out = fetchGeminiJson_([{ text: prompt }], genCfg);

  if (!out || typeof out.plainText !== 'string' || typeof out.htmlBody !== 'string') {
    return {
      plainText: `智能日报 ${dateStr}\n\n共处理 ${items.length} 封邮件。\n` + items.map(i => `- [${i.category}/${i.priority}] ${i.subject}`).join('\n'),
      htmlBody: `<h2>智能日报 ${dateStr}</h2><p>共处理 <strong>${items.length}</strong> 封邮件。</p><ul>${items.map(i => `<li>[${i.category}/${i.priority}] ${escapeHtml_(i.subject)}</li>`).join('')}</ul>`,
    };
  }

  return out;
}

function buildDailySummaryPrompt_(items, dateStr, timeZone) {
  return `
You are generating a daily report for an email automation system. Output MUST be valid JSON matching the schema. No extra text.

Context:
- Date: ${dateStr}
- Timezone: ${timeZone}
- Total processed emails: ${items.length}

Data (each item is one processed email):
${JSON.stringify(items, null, 2)}

Write a concise but insightful Chinese report that includes:
1) 今日概览（1-2句话）
2) 重点事项（按优先级列出，给出可执行建议）
3) 待复核清单（requiresReview=true 的项目要单独列出，并给出你建议的复核动作）
4) 明日建议（模式/规则优化建议：比如“某类发件人可默认归 PROMO”）

Return JSON:
{
  "plainText": "...",
  "htmlBody": "..."
}

HTML should be clean and scannable (use headings + bullet lists).`;
}

function getDailySummarySchema_() {
  return {
    type: 'object',
    properties: {
      plainText: { type: 'string' },
      htmlBody: { type: 'string' },
    },
    required: ['plainText', 'htmlBody'],
  };
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * =========================
 * 测试/维护工具
 * =========================
 */
function testProcessOneLatest() {
  processEmails();
}

function resetProcessorState() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  for (const k of Object.keys(all)) {
    if (k.startsWith('thread:') && k.endsWith(':lastMessageId')) props.deleteProperty(k);
  }
  Logger.log('✅ 已清空 thread:*:lastMessageId 状态');
}