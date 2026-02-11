/**
 * SmartEmailProcessor V4 (Google Apps Script)
 * Gmail(Outlook forward) -> Gemini -> Calendar / Tasks -> Log(Memory) -> Daily Report
 *
 * ===========================
 * ✅ 新手只需要改这里（SECTION A）
 * ===========================
 * 1) 先运行：V4_setGeminiApiKey_('你的Key')
 * 2) 再运行：V4_setupSmartEmailProcessor()
 * 3) 给你的 Outlook 转发到 Gmail 的邮件打上 Gmail 标签：CONFIG.GMAIL.SOURCE_LABEL（默认 PolyU）
 * 4) 之后触发器会每 5 分钟处理 1 个线程（1 封“最新邮件”），一次可生成多个日历/待办
 *
 * 重要说明（关于“去除上限”）：
 * - V4 不再使用“为了模型能力而设”的启发式上限（比如只给前 N 个链接/图片）
 * - 但仍保留“技术硬限制”（避免 UrlFetch / base64 / API payload 报错），并尽量设得很高
 */

// ============================================================
// SECTION A: 用户画像（你只需要改这里）
// ============================================================
const V4_PROFILE = {
  // 你是谁（越具体越好，AI 的个性化就越准）
  identity: {
    school: 'The Hong Kong Polytechnic University (PolyU)',
    faculty: 'Faculty of Business',
    department: 'Department of Logistics and Maritime Studies (LMS)',
    role: 'PhD student',
    year: 'Year 1-4 (fill in)',
    timezone: Session.getScriptTimeZone(), // 也可写死 'Asia/Hong_Kong'
  },

  // 你的偏好：什么更重要？什么通常可以忽略？
  priorities: {
    focusAreas: [
      'Seminar / workshop / conference / research talk',
      'Supervisor / lab / project related',
      'PhD milestones & deadlines (RPg, annual report, candidature, viva)',
      'Funding, reimbursement, travel, invoice/receipt',
      'Visa, accommodation, admin compliance',
    ],
    deprioritize: [
      'General campus-wide marketing not related to research',
      'Undergraduate-only teaching activities',
      'Mass newsletters (unless clearly relevant to my research/admin duties)',
    ],
  },

  // 学术兴趣：用于过滤海量群发邮件、提升“有价值邮件”的召回率
  research: {
    interests: ['(fill in) e.g., maritime logistics, supply chain, optimization, OR, AI4Ops'],
    currentTopics: ['(fill in) e.g., berth allocation, port resilience, container shipping'],
    keywordsToBoost: ['seminar', 'workshop', 'deadline', 'registration', 'candidature', 'funding'],
    keywordsToDownweight: ['newsletter', 'promotion', 'sale', 'discount', 'unsubscribe'],
  },

  // 你的日程习惯（用于生成更合理的事件/提醒）
  scheduling: {
    defaultEventDurationMinutes: 60,
    defaultDeadlineLocalTime: '17:00', // 仅给了日期时，用当天 17:00 作为截止
    addReminders: true,
  },
};

// ============================================================
// SECTION B: 输出风格预设（推广给 PhD 时可统一在这里调）
// - TERSE: 大量群发/推广 -> 极简、低冗余
// - BALANCED: 默认 -> 结构规整、信息密度高
// - WARM: 个性化/需要行动/重要日程 -> 更贴心但不啰嗦
// ============================================================
const V4_STYLE_PRESETS = {
  TERSE: {
    id: 'TERSE',
    description: '极简、低冗余。只保留结论 + 关键行动/时间/链接。适合群发推广/无关邮件。',
  },
  BALANCED: {
    id: 'BALANCED',
    description: '结构规整、信息密度高。适合大多数行政/学术通知。',
  },
  WARM: {
    id: 'WARM',
    description: '更贴心、带轻微个性化关怀，但仍保持可执行与简洁。适合重要事项/需要行动/针对性邮件。',
  },
};

// ============================================================
// SECTION B2: Prompt / 行为配置（新手可选改）
// - 这里放“提示词偏好、默认语言、冗余度策略”等，避免你去代码深处找 prompt
// ============================================================
const V4_PROMPT_CONFIG = {
  // 输出语言（建议中文，推广给 PhD 更友好）
  outputLanguage: 'Chinese',

  // AI 在“群发推广”场景下的压缩策略：尽量只输出 0-1 个 NOTE/IGNORE
  compressMassMail: true,

  // 每封邮件最多允许多少个 action items（防止极端邮件导致执行过久）
  maxItemsPerEmail: 30,

  // 是否强制在任何“有不确定性”时仍创建一个 TASK（防漏）
  createFallbackTaskWhenAmbiguous: true,

  // 事件标题前缀（可用于区分来源）
  calendarTitlePrefix: '[Email]',
};

// ============================================================
// SECTION C: 系统配置（一般不需要改）
// ============================================================
const V4_CONFIG = {
  VERSION: '4.0.0',

  GEMINI: {
    API_KEY: '', // 推荐留空，用 V4_setGeminiApiKey_ 写入 Script Properties
    MODEL_NAME: 'gemini-3-flash-preview',
    API_ROOT: 'https://generativelanguage.googleapis.com/v1beta/models',
    TEMPERATURE: 0.2,
    THINKING_LEVEL: 'HIGH', // LOW | HIGH
    MEDIA_RESOLUTION: 'MEDIA_RESOLUTION_HIGH',
    USE_RESPONSE_SCHEMA: true,
  },

  GMAIL: {
    SOURCE_LABEL: 'PolyU',
    ROOT_LABEL: 'AI',
    STATUS: {
      PROCESSING: 'AI/⏳ 处理中',
      PROCESSED: 'AI/✅ 已处理',
      REVIEW: 'AI/⚠️ 待复核',
      ERROR: 'AI/❌ 处理失败',
    },
    CATEGORY: {
      ACTIONABLE: 'AI/🎯 需行动',
      EVENT: 'AI/📅 日程',
      TASK: 'AI/✅ 待办',
      INFO: 'AI/📄 信息',
      PROMO: 'AI/📢 推广',
    },
    SYNC: {
      CALENDAR: 'AI/🔁 已同步/Calendar',
      TASKS: 'AI/🔁 已同步/Tasks',
    },
    EXTRA_PREFIX: 'AI/🏷️ ',
  },

  PIPELINE: {
    // 每次触发处理 1 个线程（最新一封邮件）
    MAX_THREADS_SCAN: 40,
    MAX_ITEMS_PER_EMAIL: 30,
    MARK_READ_AFTER_PROCESSING: false,
    ARCHIVE_AFTER_PROCESSING: false,
    DRY_RUN: false,
  },

  // 重要：这里保留的是“技术硬限制”（不是模型能力限制）
  HARD_LIMITS: {
    // 纯文本与 HTML：防止单次 prompt 超过 Apps Script/HTTP payload 的极限
    MAX_BODY_CHARS: 120000,
    MAX_HTML_CHARS: 30000,

    // 多模态：inlineData base64 会膨胀，且 UrlFetch 有 payload 上限
    // 这里设得较大，但依旧需要硬限制来避免运行期错误
    MAX_MEDIA_BYTES_EACH: 10 * 1024 * 1024,
    MAX_TOTAL_MEDIA_BYTES: 28 * 1024 * 1024,
  },

  MEMORY: {
    // Log/Memo 作为“AI 的备忘录”，每次只取当前时间往前两个月
    WINDOW_DAYS: 62,
    SPREADSHEET_NAME: 'SmartEmailProcessor Log V4',
    SHEET_TAB: 'log',
    MAX_ROWS_READ: 5000, // 技术限制：避免每次扫描整张表
  },

  DAILY_REPORT: {
    ENABLED: true,
    RECIPIENT_EMAIL: 'your_gmail@gmail.com',
    HOUR: 22,
    TEMPERATURE: 0.6,
    THINKING_LEVEL: 'HIGH',
    MAX_ITEMS_IN_PROMPT: 180,
  },
};

const V4_CACHE = {
  labels: {},
  log: { ssId: null, sheet: null },
};

// ============================================================
// 初始化 / Key / Trigger
// ============================================================
function V4_setGeminiApiKey_(apiKey) {
  const key = String(apiKey || '').trim();
  if (key.length < 20) throw new Error('apiKey 为空或太短');
  PropertiesService.getScriptProperties().setProperty('V4_GEMINI_API_KEY', key);
  Logger.log('✅ 已保存 V4_GEMINI_API_KEY 到 Script Properties');
}

// Apps Script 运行菜单只会显示“公开入口函数”（不以下划线结尾）。
// 这些 wrapper 让 V4 在编辑器里可直接选择并运行。
function V4_setGeminiApiKey() {
  throw new Error("请在编辑器中手动执行：V4_setGeminiApiKey_('YOUR_API_KEY')");
}

function V4_setupSmartEmailProcessor() {
  V4_setupSmartEmailProcessor_();
}

function V4_processEmails() {
  V4_processEmails_();
}

function V4_sendDailyReport() {
  V4_sendDailyReport_();
}

function V4_testProcessOne() {
  V4_testProcessOne_();
}

function V4_resetProcessorState() {
  V4_resetProcessorState_();
}

function V4_setupSmartEmailProcessor_() {
  V4_ensureLabelsExist_();
  V4_ensureLogSheet_();
  V4_setupTriggers_();
  Logger.log('✅ V4 初始化完成：标签 + 日志表 + 触发器');
}

function V4_setupTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) ScriptApp.deleteTrigger(t);

  ScriptApp.newTrigger('V4_processEmails').timeBased().everyMinutes(5).create();

  if (V4_CONFIG.DAILY_REPORT.ENABLED) {
    ScriptApp.newTrigger('V4_sendDailyReport')
      .timeBased()
      .atHour(V4_CONFIG.DAILY_REPORT.HOUR)
      .everyDays(1)
      .create();
  }
}

// ============================================================
// 主流程：每次处理 1 个线程（并可生成多个日历/待办）
// ============================================================
function V4_processEmails_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('⏭️ V4 上一次还没跑完，跳过本次触发');
    return;
  }

  let thread = null;

  try {
    V4_ensureLabelsExist_();
    V4_ensureLogSheet_();

    thread = V4_getNextThreadToProcess_();
    if (!thread) {
      Logger.log('✓ V4 没有新邮件需要处理');
      return;
    }

    const messages = thread.getMessages();
    if (!messages || messages.length === 0) return;

    const latest = messages[messages.length - 1];
    const threadId = thread.getId();
    const messageId = latest.getId();

    const tz = V4_PROFILE.identity.timezone || Session.getScriptTimeZone();
    const now = new Date();

    // 幂等：同一线程同一 messageId 只处理一次
    const lastProcessed = V4_getLastProcessedMessageId_(threadId);
    if (lastProcessed === messageId) {
      Logger.log(`⏭️ V4 幂等跳过：${latest.getSubject()}`);
      return;
    }

    V4_addLabel_(thread, V4_CONFIG.GMAIL.ROOT_LABEL);
    V4_addLabel_(thread, V4_CONFIG.GMAIL.STATUS.PROCESSING);
    V4_clearAiDerivedLabels_(thread);

    const email = V4_extractEmailData_(thread, messages, tz);

    const memory = V4_getMemorySnapshot_(now, tz);

    const aiRaw = V4_callGeminiForEmail_(email, memory, tz, now);
    const plan = V4_normalizeAiPlan_(aiRaw, email);

    const exec = V4_applyPlan_(thread, email, plan, memory, tz, now);

    V4_appendLogRows_(email, plan, exec, tz);
    V4_setLastProcessedMessageId_(threadId, messageId);

    V4_addLabel_(thread, V4_CONFIG.GMAIL.STATUS.PROCESSED);
    if (V4_CONFIG.PIPELINE.MARK_READ_AFTER_PROCESSING) thread.markRead();
    if (V4_CONFIG.PIPELINE.ARCHIVE_AFTER_PROCESSING) thread.moveToArchive();

    Logger.log(`✅ V4 已处理：${email.subject}`);
  } catch (e) {
    const err = (e && e.stack) ? e.stack : String(e);
    Logger.log(`❌ V4 处理失败：${err}`);
    if (thread) {
      try {
        V4_addLabel_(thread, V4_CONFIG.GMAIL.STATUS.ERROR);
        V4_addLabel_(thread, V4_CONFIG.GMAIL.STATUS.REVIEW);
      } catch (_) {}
    }
  } finally {
    if (thread) {
      try { V4_removeLabel_(thread, V4_CONFIG.GMAIL.STATUS.PROCESSING); } catch (_) {}
    }
    lock.releaseLock();
  }
}

// ============================================================
// 日报：基于“Log/Memo”生成（可选）
// ============================================================
function V4_sendDailyReport_() {
  if (!V4_CONFIG.DAILY_REPORT.ENABLED) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    V4_ensureLogSheet_();

    const tz = V4_PROFILE.identity.timezone || Session.getScriptTimeZone();
    const now = new Date();
    const dateStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd EEEE');

    const items = V4_getTodayLogItems_(now, tz);
    if (items.length === 0) {
      Logger.log('V4 今日无处理邮件，跳过日报');
      return;
    }

    const summary = V4_callGeminiForDailyReport_(items, dateStr, tz);
    const subject = `📊 [V4 智能日报] ${dateStr}（${items.length}条动作/记录）`;

    GmailApp.sendEmail(V4_CONFIG.DAILY_REPORT.RECIPIENT_EMAIL, subject, summary.plainText, {
      htmlBody: summary.htmlBody,
    });
  } catch (e) {
    Logger.log(`❌ V4 日报失败：${(e && e.stack) ? e.stack : String(e)}`);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 邮件数据提取（尽量少启发式，保留技术硬限制）
// ============================================================
function V4_extractEmailData_(thread, messages, timeZone) {
  const latest = messages[messages.length - 1];

  const subject = latest.getSubject() || '';
  const from = latest.getFrom() || '';
  const receivedAt = latest.getDate();
  const messageId = latest.getId();
  const threadId = thread.getId();
  const permalink = thread.getPermalink();

  const plainRaw = latest.getPlainBody() || '';
  const htmlRaw = latest.getBody() || '';

  const cleanedBody = V4_cleanPlainBody_(plainRaw);
  const links = V4_extractLinks_(htmlRaw, plainRaw);
  const patterns = V4_extractKeyPatterns_(cleanedBody);

  const media = V4_collectMediaParts_(latest);
  const conversationContext = V4_extractConversationContext_(messages.slice(0, -1), timeZone);

  return {
    threadId,
    messageId,
    subject,
    from,
    receivedAt,
    receivedAtStr: Utilities.formatDate(receivedAt, timeZone, 'yyyy-MM-dd HH:mm:ss'),
    permalink,

    body: cleanedBody.substring(0, V4_CONFIG.HARD_LIMITS.MAX_BODY_CHARS),
    html: (htmlRaw || '').substring(0, V4_CONFIG.HARD_LIMITS.MAX_HTML_CHARS),

    links,
    patterns,

    mediaParts: media.parts,
    mediaManifest: media.manifest,
    otherAttachments: media.otherAttachments,

    conversationContext,
    threadMessageCount: messages.length,
  };
}

function V4_cleanPlainBody_(text) {
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

function V4_extractLinks_(html, plain) {
  const results = [];
  const seen = new Set();

  function add(url, text, source) {
    const cleaned = V4_normalizeUrl_(url);
    if (!cleaned || !V4_isValidUrl_(cleaned)) return;
    if (seen.has(cleaned)) return;
    seen.add(cleaned);

    const anchorText = (text || '').trim();
    const domain = V4_extractDomain_(cleaned);
    const type = V4_classifyLinkType_(cleaned, anchorText);
    const score = V4_scoreLink_(cleaned, anchorText, type);

    results.push({ url: cleaned, text: anchorText || domain, domain, type, score, source });
  }

  const anchorRe = /<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = anchorRe.exec(html || '')) !== null) {
    const url = m[1];
    const rawText = V4_stripHtmlTags_(m[2] || '');
    const text = V4_decodeHtmlEntities_(rawText).trim();
    add(url, text, 'HTML_ANCHOR');
  }

  const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi;
  const combined = `${plain || ''}\n${V4_stripHtmlTags_(html || '')}`;
  while ((m = urlRe.exec(combined)) !== null) {
    add(m[0], '', 'BARE_URL');
  }

  results.sort((a, b) => b.score - a.score);
  return results;
}

function V4_normalizeUrl_(url) {
  if (!url) return '';
  let u = String(url).trim();
  u = u.replace(/^<|>$/g, '');
  u = u.replace(/[)\].,;:!?"']+$/g, '');
  return u;
}

function V4_isValidUrl_(url) {
  const u = String(url || '').trim();
  if (!/^https?:\/\//i.test(u)) return false;
  if (/javascript:/i.test(u)) return false;
  return true;
}

function V4_extractDomain_(url) {
  try {
    const m = String(url).match(/^https?:\/\/([^\/?#]+)/i);
    return m ? m[1] : String(url).substring(0, 60);
  } catch (_) {
    return String(url).substring(0, 60);
  }
}

function V4_classifyLinkType_(url, text) {
  const u = String(url).toLowerCase();
  const t = String(text || '').toLowerCase();
  if (u.includes('unsubscribe') || t.includes('取消订阅') || t.includes('unsubscribe')) return 'UNSUBSCRIBE';
  if (u.includes('zoom.') || u.includes('teams.microsoft') || u.includes('meet.google') || u.includes('webex') || u.includes('tencentmeeting')) return 'MEETING';
  if (u.includes('calendar') || u.includes('calendly') || u.endsWith('.ics') || t.includes('add to calendar') || t.includes('日程')) return 'CALENDAR';
  if (u.includes('pay') || u.includes('invoice') || u.includes('billing') || u.includes('stripe') || u.includes('paypal')) return 'PAYMENT';
  if (u.includes('docs.google') || u.includes('drive.google') || u.includes('dropbox') || u.includes('onedrive') || u.includes('sharepoint') || u.includes('notion.so') || u.includes('figma.com')) return 'DOCUMENT';
  if (u.includes('register') || u.includes('signup') || u.includes('registration') || u.includes('rsvp') || u.includes('eventbrite') || t.includes('报名') || t.includes('注册')) return 'REGISTRATION';
  if (u.includes('login') || u.includes('verify') || u.includes('otp') || u.includes('reset') || t.includes('验证码')) return 'AUTH';
  return 'GENERAL';
}

function V4_scoreLink_(url, text, type) {
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
    UNSUBSCRIBE: -50,
  };
  let score = baseByType[type] != null ? baseByType[type] : 20;

  const goodKw = ['join', 'register', 'rsvp', 'ticket', 'invoice', 'pay', 'verify', 'confirm', 'download', 'open', 'view', '报名', '注册', '加入', '验证', '确认'];
  for (const kw of goodKw) {
    if (t.includes(kw) || u.includes(kw)) { score += 6; break; }
  }
  if (u.includes('utm_') || u.includes('utm-') || u.includes('tracking') || u.includes('trk=')) score -= 20;
  if (t && t.trim().length >= 8) score += 5;
  return score;
}

function V4_stripHtmlTags_(html) {
  return String(html || '').replace(/<[^>]*>/g, ' ');
}

function V4_decodeHtmlEntities_(text) {
  return String(text || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function V4_extractKeyPatterns_(text) {
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

  let m;
  const emailRe = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
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

  const dateRe = /\b(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/g;
  while ((m = dateRe.exec(t)) !== null) pushUniq(patterns.dates, m[1]);

  const timeRe = /\b(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\b/g;
  while ((m = timeRe.exec(t)) !== null) pushUniq(patterns.times, m[1]);

  return patterns;
}

function V4_extractConversationContext_(previousMessages, timeZone) {
  const msgs = previousMessages || [];
  if (msgs.length === 0) return null;

  // 不做“内容理解上限”，但做“技术硬限制”避免 payload 爆炸
  const max = Math.min(6, msgs.length);
  const last = msgs.slice(-max);

  return last.map(msg => {
    const body = V4_cleanPlainBody_(msg.getPlainBody() || '').substring(0, 2000);
    return {
      from: msg.getFrom() || '',
      date: Utilities.formatDate(msg.getDate(), timeZone, 'yyyy-MM-dd HH:mm:ss'),
      snippet: body,
    };
  });
}

function V4_collectMediaParts_(message) {
  const attachments = message.getAttachments({ includeInlineImages: true, includeAttachments: true }) || [];

  const parts = [];
  const manifest = [];
  const otherAttachments = [];

  let totalBytes = 0;
  let index = 0;

  for (const att of attachments) {
    const mimeType = att.getContentType() || '';
    const name = att.getName() || '(unnamed)';
    const size = att.getSize() || 0;

    const isImage = mimeType.startsWith('image/');
    const isPdf = mimeType === 'application/pdf' || /\.pdf$/i.test(name);

    if (!isImage && !isPdf) {
      otherAttachments.push({ filename: name, mimeType: mimeType || 'unknown', sizeKB: Math.round(size / 1024) });
      continue;
    }

    // 技术硬限制：避免 payload/内存爆炸
    if (size > V4_CONFIG.HARD_LIMITS.MAX_MEDIA_BYTES_EACH) {
      otherAttachments.push({ filename: name, mimeType, sizeKB: Math.round(size / 1024), note: 'too_large_for_inlineData' });
      continue;
    }
    if (totalBytes + size > V4_CONFIG.HARD_LIMITS.MAX_TOTAL_MEDIA_BYTES) {
      otherAttachments.push({ filename: name, mimeType, sizeKB: Math.round(size / 1024), note: 'exceeds_total_media_budget' });
      continue;
    }

    index += 1;
    totalBytes += size;

    manifest.push({
      index,
      kind: isPdf ? 'pdf' : 'image',
      filename: name,
      mimeType,
      sizeKB: Math.round(size / 1024),
    });

    parts.push({ text: `MEDIA #${index}: kind=${isPdf ? 'pdf' : 'image'}, filename=${name}, mimeType=${mimeType}, sizeKB=${Math.round(size / 1024)}` });
    parts.push({
      inlineData: {
        mimeType,
        data: Utilities.base64Encode(att.getBytes()),
      },
    });
  }

  return { parts, manifest, otherAttachments };
}

// ============================================================
// Memory：从 Log 里截取过去两个月，作为 AI 备忘录输入 + 去重依据
// ============================================================
function V4_getMemorySnapshot_(now, timeZone) {
  const sheet = V4_ensureLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { windowDays: V4_CONFIG.MEMORY.WINDOW_DAYS, items: [], memoHints: [] };

  const startRow = Math.max(2, lastRow - V4_CONFIG.MEMORY.MAX_ROWS_READ + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  const windowStart = new Date(now.getTime() - V4_CONFIG.MEMORY.WINDOW_DAYS * 24 * 60 * 60 * 1000);

  // Columns (see V4_ensureLogSheet_ header)
  // 0 processedAt(Date)
  // 1 threadId
  // 2 messageId
  // 3 subject
  // 4 from
  // 5 receivedAt(Date)
  // 6 permalink
  // 7 style
  // 8 overallCategory
  // 9 overallPriority
  // 10 itemIndex
  // 11 itemType
  // 12 requestedAction
  // 13 appliedAction
  // 14 memoryId
  // 15 fingerprint
  // 16 title
  // 17 startTime
  // 18 endTime
  // 19 dueTime
  // 20 allDay
  // 21 location
  // 22 calendarEventId
  // 23 taskId
  // 24 requiresReview
  // 25 confidence
  // 26 userFacingHeadline
  // 27 itemMemo
  // 28 assistantMemo
  const items = [];
  const memoHints = [];

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const processedAt = row[0];
    if (!(processedAt instanceof Date)) continue;
    if (processedAt < windowStart) break;

    const itemType = String(row[11] || '').trim();
    const appliedAction = String(row[13] || '').trim();

    // 备忘录：优先 itemMemo，其次 assistantMemo（减少噪音，但保留可用于去重/更新的线索）
    const itemMemo = String(row[27] || '').trim();
    const assistantMemo = String(row[28] || '').trim();
    const memo = itemMemo || assistantMemo;
    if (memo && (appliedAction !== 'SKIP' || itemType === 'NOTE')) {
      memoHints.push(`[${Utilities.formatDate(processedAt, timeZone, 'yyyy-MM-dd')}] ${memo}`.substring(0, 260));
    }

    items.push({
      processedAt: Utilities.formatDate(processedAt, timeZone, 'yyyy-MM-dd HH:mm:ss'),
      memoryId: String(row[14] || ''),
      fingerprint: String(row[15] || ''),
      itemType,
      appliedAction,
      title: String(row[16] || ''),
      startTime: String(row[17] || ''),
      endTime: String(row[18] || ''),
      dueTime: String(row[19] || ''),
      allDay: Boolean(row[20]),
      location: String(row[21] || ''),
      calendarEventId: String(row[22] || ''),
      taskId: String(row[23] || ''),
      subject: String(row[3] || ''),
      from: String(row[4] || ''),
      permalink: String(row[6] || ''),
    });
  }

  // 去重：同一个 memoryId 只保留最新一条
  const latestByMemoryId = {};
  for (const it of items) {
    if (!it.memoryId) continue;
    if (!latestByMemoryId[it.memoryId]) latestByMemoryId[it.memoryId] = it;
  }

  const compact = Object.values(latestByMemoryId).slice(0, 600);

  return {
    windowDays: V4_CONFIG.MEMORY.WINDOW_DAYS,
    items: compact,
    memoHints: memoHints.slice(0, 120),
  };
}

// ============================================================
// Gemini 调用：输出一个“计划（plan）”，包含多个 items（可多日历/多待办）
// ============================================================
function V4_callGeminiForEmail_(email, memory, timeZone, now) {
  const prompt = V4_buildEmailPrompt_(email, memory, timeZone, now);
  const parts = [{ text: prompt }];
  if (email.mediaParts && email.mediaParts.length) parts.push(...email.mediaParts);

  const genCfg = {
    responseMimeType: 'application/json',
    temperature: V4_CONFIG.GEMINI.TEMPERATURE,
    thinkingConfig: { thinkingLevel: V4_CONFIG.GEMINI.THINKING_LEVEL },
  };
  if (V4_CONFIG.GEMINI.MEDIA_RESOLUTION) genCfg.mediaResolution = V4_CONFIG.GEMINI.MEDIA_RESOLUTION;
  if (V4_CONFIG.GEMINI.USE_RESPONSE_SCHEMA) genCfg.responseSchema = V4_getEmailPlanSchema_();

  return V4_fetchGeminiJson_(parts, genCfg);
}

function V4_buildEmailPrompt_(email, memory, timeZone, now) {
  const profile = V4_PROFILE;
  const styles = Object.values(V4_STYLE_PRESETS).map(s => `- ${s.id}: ${s.description}`).join('\n');

  const links = (email.links || []).map(l => ({
    type: l.type,
    score: l.score,
    text: l.text,
    url: l.url,
    domain: l.domain,
  }));

  const memoryBlock = JSON.stringify({
    windowDays: memory.windowDays,
    memoHints: memory.memoHints || [],
    items: memory.items || [],
  }, null, 2);

  return `
SYSTEM (strict):
You are an autonomous email executive assistant inside an automation. Your job is to produce a deterministic action plan.
Output MUST be valid JSON matching the provided schema. Do NOT output any extra text.

User profile (for personalization & relevance filtering):
${JSON.stringify(profile, null, 2)}

Available output styles (choose one per email):
${styles}

Current context:
- Now: ${Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm:ss')}
- Timezone: ${timeZone}
- Memory window: last ${memory.windowDays} days

AI memory (Log excerpts; treat as YOUR own memo, not user conversation):
${memoryBlock}

Email (latest message in thread):
- Subject: ${email.subject}
- From: ${email.from}
- ReceivedAt: ${email.receivedAtStr}
- Thread message count: ${email.threadMessageCount}
- Gmail permalink: ${email.permalink}

Cleaned plain body:
${email.body}

HTML snippet:
${email.html}

Extracted links (deduped + scored; include as reference, not ground truth):
${JSON.stringify(links, null, 2)}

Extracted patterns (rough hints):
${JSON.stringify(email.patterns || {}, null, 2)}

Conversation context (previous messages):
${email.conversationContext ? JSON.stringify(email.conversationContext, null, 2) : '(none)'}

Attached media manifest (actual media follows this text in the same order):
${email.mediaManifest && email.mediaManifest.length ? JSON.stringify(email.mediaManifest, null, 2) : '(none)'}
Other attachments (not included as media):
${email.otherAttachments && email.otherAttachments.length ? JSON.stringify(email.otherAttachments, null, 2) : '(none)'}

Your mission:
1) Understand the email deeply (including images/PDFs if provided).
2) Decide relevance & whether it is mass-mail/promo vs actionable.
3) Extract ALL concrete arrangements:
   - If the email contains multiple dates/times for multiple sessions, output MULTIPLE calendar items.
   - If the email contains multiple independent action requests, output MULTIPLE tasks.
4) Use AI memory to avoid duplication:
   - If this is a repeated reminder of something already created recently, prefer SKIP or UPDATE rather than CREATE.
   - If details changed (time/location/link), prefer UPDATE the existing item instead of creating duplicates.
5) Choose an output style automatically:
   - Mass promo -> TERSE
   - Normal admin/seminar -> BALANCED
   - Personalized/important -> WARM

Time rules (important):
- Use RFC3339 with timezone offset, e.g. 2026-01-11T09:30:00+08:00
- If only a date is known, set deadline to ${profile.scheduling.defaultDeadlineLocalTime} local time.
- If timezone is explicitly mentioned, respect it; otherwise default to ${timeZone}.
- If ambiguous, set requiresReview=true and still create a TASK (not INFO) to avoid missing important things.

Output language rules:
- userFacing.headline / bullets / nextSteps: ${V4_PROMPT_CONFIG.outputLanguage}
- classification.reasoning / assistantMemo / itemMemo / description: ${V4_PROMPT_CONFIG.outputLanguage}
- title: short; Chinese or English ok
- Avoid unnecessary emojis; keep scannable with short bullets.

Implementation constraints (do NOT mention to user, just comply):
- You may output up to ${V4_PROMPT_CONFIG.maxItemsPerEmail} items.
- If compressMassMail=${String(V4_PROMPT_CONFIG.compressMassMail)}, then for obvious mass promo/newsletter you should output at most 1 NOTE/IGNORE and no calendar/tasks.

Return JSON only.`;
}

function V4_getEmailPlanSchema_() {
  return {
    type: 'object',
    properties: {
      selectedStyle: { type: 'string', enum: ['TERSE', 'BALANCED', 'WARM'] },
      classification: {
        type: 'object',
        properties: {
          overallCategory: { type: 'string', enum: ['ACTIONABLE', 'INFO', 'PROMO'] },
          priority: { type: 'string', enum: ['HIGH', 'MEDIUM', 'LOW'] },
          isMassMail: { type: 'boolean' },
          needsUserAttention: { type: 'boolean' },
          reasoning: { type: 'string' },
        },
        required: ['overallCategory', 'priority', 'isMassMail', 'needsUserAttention', 'reasoning'],
      },
      assistantMemo: { type: 'string', description: 'AI自己的备忘录（用于未来去重/更新决策）' },
      suggestedLabels: { type: 'array', items: { type: 'string' }, maxItems: 8 },
      items: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            itemType: { type: 'string', enum: ['CALENDAR_EVENT', 'TASK', 'NOTE', 'IGNORE'] },
            action: { type: 'string', enum: ['CREATE', 'UPDATE', 'SKIP', 'CANCEL'] },
            title: { type: 'string' },
            confidence: { type: 'number' },
            requiresReview: { type: 'boolean' },

            // For UPDATE/CANCEL, AI can reference memoryId (preferred) or eventId/taskId if known
            target: {
              type: 'object',
              properties: {
                memoryId: { type: 'string', nullable: true },
                calendarEventId: { type: 'string', nullable: true },
                taskId: { type: 'string', nullable: true },
              },
            },

            time: {
              type: 'object',
              properties: {
                startTime: { type: 'string', nullable: true },
                endTime: { type: 'string', nullable: true },
                deadline: { type: 'string', nullable: true },
                allDay: { type: 'boolean' },
              },
              required: ['allDay'],
            },

            location: { type: 'string', nullable: true },
            description: { type: 'string', nullable: true },

            keyLinks: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  purpose: { type: 'string' },
                  url: { type: 'string' },
                },
                required: ['purpose', 'url'],
              },
              maxItems: 6,
            },
            keyNumbers: { type: 'array', items: { type: 'string' }, maxItems: 20 },

            userFacing: {
              type: 'object',
              properties: {
                headline: { type: 'string' },
                bullets: { type: 'array', items: { type: 'string' }, maxItems: 10 },
                nextSteps: { type: 'array', items: { type: 'string' }, maxItems: 10 },
              },
              required: ['headline', 'bullets', 'nextSteps'],
            },

            itemMemo: { type: 'string', nullable: true },
          },
          required: ['itemType', 'action', 'title', 'confidence', 'requiresReview', 'time', 'userFacing', 'keyLinks', 'keyNumbers'],
        },
      },
    },
    required: ['selectedStyle', 'classification', 'assistantMemo', 'items'],
  };
}

function V4_fetchGeminiJson_(parts, generationConfig) {
  const apiKey = V4_getApiKey_();
  const url = `${V4_CONFIG.GEMINI.API_ROOT}/${encodeURIComponent(V4_CONFIG.GEMINI.MODEL_NAME)}:generateContent`;

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

function V4_getApiKey_() {
  const fromConfig = String(V4_CONFIG.GEMINI.API_KEY || '').trim();
  if (fromConfig && fromConfig !== 'YOUR_API_KEY_HERE') return fromConfig;

  const fromProps = PropertiesService.getScriptProperties().getProperty('V4_GEMINI_API_KEY');
  if (fromProps && fromProps.trim()) return fromProps.trim();

  throw new Error('未配置 Gemini API Key：请运行 V4_setGeminiApiKey_("...") 或在 V4_CONFIG.GEMINI.API_KEY 写入');
}

// ============================================================
// 计划归一化 + 执行（多事项）
// ============================================================
function V4_normalizeAiPlan_(raw, email) {
  const r = raw && typeof raw === 'object' ? raw : {};

  const style = ['TERSE', 'BALANCED', 'WARM'].includes(String(r.selectedStyle || '').toUpperCase())
    ? String(r.selectedStyle).toUpperCase()
    : 'BALANCED';

  const cls = r.classification && typeof r.classification === 'object' ? r.classification : {};
  const overallCategory = ['ACTIONABLE', 'INFO', 'PROMO'].includes(String(cls.overallCategory || '').toUpperCase())
    ? String(cls.overallCategory).toUpperCase()
    : 'INFO';
  const priority = ['HIGH', 'MEDIUM', 'LOW'].includes(String(cls.priority || '').toUpperCase())
    ? String(cls.priority).toUpperCase()
    : 'MEDIUM';

  const items = Array.isArray(r.items) ? r.items : [];

  const maxItems = Number(V4_CONFIG.PIPELINE.MAX_ITEMS_PER_EMAIL || V4_PROMPT_CONFIG.maxItemsPerEmail || 30);
  const normalizedItems = items.slice(0, Math.max(1, Math.min(60, maxItems))).map((it, idx) => {
    const o = it && typeof it === 'object' ? it : {};
    const itemType = ['CALENDAR_EVENT', 'TASK', 'NOTE', 'IGNORE'].includes(String(o.itemType || '').toUpperCase())
      ? String(o.itemType).toUpperCase()
      : 'NOTE';
    const action = ['CREATE', 'UPDATE', 'SKIP', 'CANCEL'].includes(String(o.action || '').toUpperCase())
      ? String(o.action).toUpperCase()
      : 'SKIP';

    const time = o.time && typeof o.time === 'object' ? o.time : {};
    const userFacing = o.userFacing && typeof o.userFacing === 'object' ? o.userFacing : {};

    return {
      itemIndex: idx + 1,
      itemType,
      action,
      title: String(o.title || email.subject || '邮件事项').trim().substring(0, 120),
      confidence: V4_clamp01_(Number(o.confidence)),
      requiresReview: typeof o.requiresReview === 'boolean' ? o.requiresReview : false,
      target: o.target && typeof o.target === 'object' ? {
        memoryId: o.target.memoryId ? String(o.target.memoryId) : null,
        calendarEventId: o.target.calendarEventId ? String(o.target.calendarEventId) : null,
        taskId: o.target.taskId ? String(o.target.taskId) : null,
      } : { memoryId: null, calendarEventId: null, taskId: null },
      time: {
        startTime: time.startTime ? String(time.startTime).trim() : null,
        endTime: time.endTime ? String(time.endTime).trim() : null,
        deadline: time.deadline ? String(time.deadline).trim() : null,
        allDay: typeof time.allDay === 'boolean' ? time.allDay : false,
      },
      location: o.location ? String(o.location).trim().substring(0, 300) : null,
      description: o.description ? String(o.description).trim().substring(0, 8000) : null,
      keyLinks: Array.isArray(o.keyLinks)
        ? o.keyLinks
          .filter(Boolean)
          .slice(0, 6)
          .map(x => ({
            purpose: String((x && x.purpose) || 'link').trim().substring(0, 80),
            url: String((x && x.url) || '').trim().substring(0, 500),
          }))
          .filter(x => x.url)
        : [],
      keyNumbers: Array.isArray(o.keyNumbers) ? o.keyNumbers.filter(Boolean).slice(0, 20).map(x => String(x).trim().substring(0, 60)) : [],
      userFacing: {
        headline: String(userFacing.headline || '').trim().substring(0, 400),
        bullets: Array.isArray(userFacing.bullets) ? userFacing.bullets.filter(Boolean).slice(0, 10).map(x => String(x).trim().substring(0, 300)) : [],
        nextSteps: Array.isArray(userFacing.nextSteps) ? userFacing.nextSteps.filter(Boolean).slice(0, 10).map(x => String(x).trim().substring(0, 300)) : [],
      },
      itemMemo: o.itemMemo ? String(o.itemMemo).trim().substring(0, 800) : null,
    };
  });

  return {
    selectedStyle: style,
    classification: {
      overallCategory,
      priority,
      isMassMail: Boolean(cls.isMassMail),
      needsUserAttention: Boolean(cls.needsUserAttention),
      reasoning: String(cls.reasoning || '').trim().substring(0, 800),
    },
    assistantMemo: String(r.assistantMemo || '').trim().substring(0, 1200),
    suggestedLabels: Array.isArray(r.suggestedLabels) ? r.suggestedLabels.filter(Boolean).slice(0, 8).map(x => String(x).trim().substring(0, 30)) : [],
    items: normalizedItems,
    raw: r,
  };
}

function V4_clamp01_(n) {
  const x = Number(n);
  if (Number.isNaN(x)) return 0.6;
  return Math.max(0, Math.min(1, x));
}

function V4_applyPlan_(thread, email, plan, memory, timeZone, now) {
  V4_applyLabelsFromPlan_(thread, plan);

  const memoryIndex = V4_buildMemoryIndex_(memory);
  const results = [];

  for (const item of plan.items) {
    const exec = V4_executeOneItem_(thread, email, item, plan, memoryIndex, timeZone, now);
    results.push(exec);
  }

  // 同步标签（基于实际执行结果）
  const didCalendar = results.some(r => r && (r.appliedAction === 'CREATE' || r.appliedAction === 'UPDATE') && r.calendarEventId);
  const didTasks = results.some(r => r && (r.appliedAction === 'CREATE' || r.appliedAction === 'UPDATE') && r.taskId);
  if (didCalendar) V4_addLabel_(thread, V4_CONFIG.GMAIL.SYNC.CALENDAR);
  if (didTasks) V4_addLabel_(thread, V4_CONFIG.GMAIL.SYNC.TASKS);

  return {
    perItem: results,
  };
}

function V4_buildMemoryIndex_(memory) {
  const idx = { byMemoryId: {}, byEventId: {}, byTaskId: {} };
  const items = (memory && Array.isArray(memory.items)) ? memory.items : [];
  for (const it of items) {
    if (it.memoryId) idx.byMemoryId[it.memoryId] = it;
    if (it.calendarEventId) idx.byEventId[it.calendarEventId] = it;
    if (it.taskId) idx.byTaskId[it.taskId] = it;
  }
  return idx;
}

function V4_executeOneItem_(thread, email, item, plan, memoryIndex, timeZone, now) {
  const computed = V4_computeEntityIds_(item, timeZone);
  const memoryId = computed.memoryId;
  const fingerprint = computed.fingerprint;

  const base = {
    itemIndex: item.itemIndex,
    itemType: item.itemType,
    requestedAction: item.action,
    appliedAction: 'SKIP',
    memoryId,
    fingerprint,
    title: item.title,
    startTime: computed.startTimeStr,
    endTime: computed.endTimeStr,
    dueTime: computed.dueTimeStr,
    allDay: Boolean(item.time && item.time.allDay),
    location: item.location || '',
    calendarEventId: '',
    taskId: '',
    requiresReview: Boolean(item.requiresReview),
    confidence: item.confidence,
    userFacingHeadline: item.userFacing.headline || '',
    itemMemo: item.itemMemo || '',
    assistantMemo: plan.assistantMemo || '',
    error: '',
  };

  if (item.itemType === 'IGNORE') return base;

  // NOTE: 只写入日志/备忘录，不做外部同步
  if (item.itemType === 'NOTE') {
    base.appliedAction = 'NOTE';
    return base;
  }

  // 去重：如果 CREATE 但 memory 中已存在相同 memoryId，则默认 SKIP（除非 AI 明确要求 UPDATE）
  const existing = memoryId && memoryIndex.byMemoryId[memoryId] ? memoryIndex.byMemoryId[memoryId] : null;
  if (item.action === 'CREATE' && existing && (existing.calendarEventId || existing.taskId)) {
    base.appliedAction = 'SKIP_DUPLICATE';
    return base;
  }

  if (V4_CONFIG.PIPELINE.DRY_RUN) {
    base.appliedAction = `DRY_RUN_${item.action}`;
    return base;
  }

  try {
    if (item.itemType === 'CALENDAR_EVENT') {
      const exec = V4_applyCalendarOp_(email, item, plan, computed, existing, memoryIndex, timeZone);
      // 可靠性增强：如果 AI 想建日历但没给出可解析时间，且配置允许，则自动降级为 Task（防漏）
      if (exec && exec.appliedAction === 'SKIP_NO_TIME' && V4_PROMPT_CONFIG.createFallbackTaskWhenAmbiguous) {
        const fallbackItem = {
          ...item,
          itemType: 'TASK',
          action: 'CREATE',
          time: { ...item.time, allDay: false, deadline: item.time && (item.time.deadline || item.time.startTime) ? (item.time.deadline || item.time.startTime) : null },
        };
        const taskExec = V4_applyTaskOp_(email, fallbackItem, plan, computed, existing, memoryIndex, timeZone);
        return {
          ...base,
          appliedAction: `FALLBACK_TASK_${taskExec.appliedAction || 'SKIP'}`,
          taskId: taskExec.taskId || '',
        };
      }
      return { ...base, ...exec };
    }
    if (item.itemType === 'TASK') {
      const exec = V4_applyTaskOp_(email, item, plan, computed, existing, memoryIndex, timeZone);
      return { ...base, ...exec };
    }
    return base;
  } catch (e) {
    base.appliedAction = 'ERROR';
    base.error = String(e && e.stack ? e.stack : e);
    base.requiresReview = true;
    V4_addLabel_(thread, V4_CONFIG.GMAIL.STATUS.REVIEW);
    return base;
  }
}

function V4_computeEntityIds_(item, timeZone) {
  const t = item.time || {};

  const start = V4_parseDateTime_(t.startTime, timeZone);
  const end = V4_parseDateTime_(t.endTime, timeZone);
  const deadline = V4_parseDateTime_(t.deadline, timeZone);

  const startTimeStr = start ? start.toISOString() : (t.startTime || '');
  const endTimeStr = end ? end.toISOString() : (t.endTime || '');
  const dueTimeStr = deadline ? deadline.toISOString() : (t.deadline || '');

  const fpInput = [
    item.itemType,
    V4_normText_(item.title),
    t.allDay ? 'ALLDAY' : 'TIMED',
    start ? start.toISOString() : '',
    end ? end.toISOString() : '',
    deadline ? deadline.toISOString() : '',
    V4_normText_(item.location || ''),
  ].join('|');

  const fingerprint = V4_sha256Hex_(fpInput);
  const memoryId = fingerprint ? fingerprint.substring(0, 12) : '';

  return { start, end, deadline, startTimeStr, endTimeStr, dueTimeStr, fingerprint, memoryId };
}

function V4_normText_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ').substring(0, 200);
}

function V4_sha256Hex_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function V4_parseDateTime_(value, timeZone) {
  if (!value) return null;
  const s = String(value).trim();
  if (!s) return null;

  // Date only: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const localTime = V4_PROFILE.scheduling.defaultDeadlineLocalTime || '17:00';
    const d = new Date(`${s}T${localTime}:00`);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

function V4_applyCalendarOp_(email, item, plan, computed, existing, memoryIndex, timeZone) {
  const requested = item.action;

  // Resolve target for UPDATE/CANCEL
  const targetEventId =
    (item.target && item.target.calendarEventId) ||
    (item.target && item.target.memoryId && memoryIndex.byMemoryId[item.target.memoryId] && memoryIndex.byMemoryId[item.target.memoryId].calendarEventId) ||
    (existing && existing.calendarEventId) ||
    '';

  const notes = V4_buildNotes_(email, item, plan, timeZone);
  const cal = CalendarApp.getDefaultCalendar();

  if (requested === 'CANCEL') {
    if (!targetEventId) return { appliedAction: 'SKIP_NO_TARGET' };
    const ev = cal.getEventById(targetEventId);
    if (!ev) return { appliedAction: 'SKIP_NOT_FOUND' };
    ev.deleteEvent();
    return { appliedAction: 'CANCEL', calendarEventId: targetEventId };
  }

  if (requested === 'UPDATE') {
    if (!targetEventId) return { appliedAction: 'SKIP_NO_TARGET' };
    const ev = cal.getEventById(targetEventId);
    if (!ev) return { appliedAction: 'SKIP_NOT_FOUND' };

    const start = computed.start;
    let end = computed.end || (start ? new Date(start.getTime() + (V4_PROFILE.scheduling.defaultEventDurationMinutes || 60) * 60000) : null);
    if (start && end && end.getTime() <= start.getTime()) {
      end = new Date(start.getTime() + (V4_PROFILE.scheduling.defaultEventDurationMinutes || 60) * 60000);
    }

    if (item.time && item.time.allDay) {
      if (start) {
        const startDate = new Date(start); startDate.setHours(0, 0, 0, 0);
        const endDate = end ? new Date(end) : new Date(startDate);
        endDate.setHours(0, 0, 0, 0);
        if (endDate.getTime() <= startDate.getTime()) endDate.setDate(endDate.getDate() + 1);
        ev.setAllDayDates(startDate, endDate);
      }
    } else {
      if (start && end) ev.setTime(start, end);
    }
    if (item.title) ev.setTitle(item.title.substring(0, 120));
    if (item.location != null) ev.setLocation(item.location);
    ev.setDescription(notes);
    V4_applyEventReminders_(ev, plan.classification.priority);
    return { appliedAction: 'UPDATE', calendarEventId: targetEventId };
  }

  // CREATE
  if (requested === 'CREATE') {
    const start = computed.start;
    if (!start && !(item.time && item.time.allDay)) {
      return { appliedAction: 'SKIP_NO_TIME' };
    }

    const prefix = String(V4_PROMPT_CONFIG.calendarTitlePrefix || '').trim();
    const title = `${prefix ? prefix + ' ' : ''}${item.title}`.substring(0, 120);

    if (item.time && item.time.allDay) {
      const startDate = start ? new Date(start) : new Date(email.receivedAt);
      startDate.setHours(0, 0, 0, 0);
      const endDate = computed.end ? new Date(computed.end) : new Date(startDate);
      endDate.setHours(0, 0, 0, 0);
      if (endDate.getTime() <= startDate.getTime()) endDate.setDate(endDate.getDate() + 1);
      const ev = cal.createAllDayEvent(title, startDate, endDate, { location: item.location || '', description: notes });
      V4_applyEventReminders_(ev, plan.classification.priority);
      return { appliedAction: 'CREATE', calendarEventId: ev.getId() };
    }

    let end = computed.end || new Date(start.getTime() + (V4_PROFILE.scheduling.defaultEventDurationMinutes || 60) * 60000);
    if (end.getTime() <= start.getTime()) {
      end = new Date(start.getTime() + (V4_PROFILE.scheduling.defaultEventDurationMinutes || 60) * 60000);
    }
    const ev = cal.createEvent(title, start, end, { location: item.location || '', description: notes });
    V4_applyEventReminders_(ev, plan.classification.priority);
    return { appliedAction: 'CREATE', calendarEventId: ev.getId() };
  }

  return { appliedAction: 'SKIP' };
}

function V4_applyEventReminders_(event, priority) {
  if (!V4_PROFILE.scheduling.addReminders) return;
  try { event.removeAllReminders(); } catch (_) {}

  try {
    if (priority === 'HIGH') {
      event.addPopupReminder(1440);
      event.addPopupReminder(120);
      event.addPopupReminder(30);
    } else if (priority === 'MEDIUM') {
      event.addPopupReminder(180);
      event.addPopupReminder(30);
    } else {
      event.addPopupReminder(30);
    }
  } catch (_) {}
}

function V4_applyTaskOp_(email, item, plan, computed, existing, memoryIndex, timeZone) {
  const requested = item.action;

  if (!V4_isTasksServiceAvailable_()) {
    return { appliedAction: 'SKIP_TASKS_NOT_ENABLED' };
  }

  const targetTaskId =
    (item.target && item.target.taskId) ||
    (item.target && item.target.memoryId && memoryIndex.byMemoryId[item.target.memoryId] && memoryIndex.byMemoryId[item.target.memoryId].taskId) ||
    (existing && existing.taskId) ||
    '';

  const notes = V4_buildNotes_(email, item, plan, timeZone);
  const due = V4_normalizeTaskDueRfc3339_(computed.deadline || computed.start);

  if (requested === 'CANCEL') {
    if (!targetTaskId) return { appliedAction: 'SKIP_NO_TARGET' };
    Tasks.Tasks.remove('@default', targetTaskId);
    return { appliedAction: 'CANCEL', taskId: targetTaskId };
  }

  if (requested === 'UPDATE') {
    if (!targetTaskId) return { appliedAction: 'SKIP_NO_TARGET' };
    const task = Tasks.Tasks.get('@default', targetTaskId);
    if (!task) return { appliedAction: 'SKIP_NOT_FOUND' };
    task.title = item.title.substring(0, 200);
    task.notes = String(notes || '').substring(0, 8000);
    if (due) task.due = due;
    const updated = Tasks.Tasks.update(task, '@default', targetTaskId);
    return { appliedAction: 'UPDATE', taskId: updated && updated.id ? updated.id : targetTaskId };
  }

  if (requested === 'CREATE') {
    const task = {
      title: item.title.substring(0, 200),
      notes: String(notes || '').substring(0, 8000),
    };
    if (due) task.due = due;
    const inserted = Tasks.Tasks.insert(task, '@default');
    return { appliedAction: 'CREATE', taskId: inserted && inserted.id ? inserted.id : '' };
  }

  return { appliedAction: 'SKIP' };
}

function V4_isTasksServiceAvailable_() {
  return typeof Tasks !== 'undefined' && Tasks.Tasks && typeof Tasks.Tasks.insert === 'function';
}

function V4_normalizeTaskDueRfc3339_(dateObj) {
  if (!dateObj) return '';
  const d = dateObj instanceof Date ? dateObj : null;
  if (!d || Number.isNaN(d.getTime())) return '';
  return d.toISOString();
}

function V4_buildNotes_(email, item, plan, timeZone) {
  const lines = [];

  // 结构规整：头部给“用户可读”摘要
  if (item.userFacing && item.userFacing.headline) {
    lines.push(item.userFacing.headline);
    lines.push('');
  }

  if (item.userFacing && item.userFacing.bullets && item.userFacing.bullets.length) {
    lines.push('要点：');
    for (const b of item.userFacing.bullets) lines.push(`- ${b}`);
    lines.push('');
  }

  if (item.userFacing && item.userFacing.nextSteps && item.userFacing.nextSteps.length) {
    lines.push('下一步：');
    for (const s of item.userFacing.nextSteps) lines.push(`□ ${s}`);
    lines.push('');
  }

  if (item.keyNumbers && item.keyNumbers.length) {
    lines.push('关键编号/信息：');
    lines.push(item.keyNumbers.join(' | '));
    lines.push('');
  }

  if (item.keyLinks && item.keyLinks.length) {
    lines.push('关键链接：');
    for (const l of item.keyLinks) {
      if (!l || !l.url) continue;
      lines.push(`- ${l.purpose || 'link'}: ${l.url}`);
    }
    lines.push('');
  }

  if (item.description) {
    lines.push('补充说明：');
    lines.push(item.description);
    lines.push('');
  }

  // AI memo（面向未来去重/更新）
  if (item.itemMemo) {
    lines.push('AI 备忘录：');
    lines.push(item.itemMemo);
    lines.push('');
  }

  // 溯源
  lines.push('───────────');
  lines.push(`原邮件：${email.permalink}`);
  lines.push(`主题：${email.subject}`);
  lines.push(`发件人：${email.from}`);
  lines.push(`收到时间：${email.receivedAtStr}`);
  lines.push(`执行风格：${plan.selectedStyle}`);

  return lines.join('\n').substring(0, 8000);
}

function V4_applyLabelsFromPlan_(thread, plan) {
  const c = plan.classification || {};
  const cat = String(c.overallCategory || 'INFO');

  if (cat === 'PROMO') V4_addLabel_(thread, V4_CONFIG.GMAIL.CATEGORY.PROMO);
  else if (cat === 'ACTIONABLE') V4_addLabel_(thread, V4_CONFIG.GMAIL.CATEGORY.ACTIONABLE);
  else V4_addLabel_(thread, V4_CONFIG.GMAIL.CATEGORY.INFO);

  // 细分：只要有日历/待办 item，就再贴更具体标签
  const hasEvent = plan.items.some(x => x.itemType === 'CALENDAR_EVENT' && x.action !== 'SKIP');
  const hasTask = plan.items.some(x => x.itemType === 'TASK' && x.action !== 'SKIP');
  if (hasEvent) V4_addLabel_(thread, V4_CONFIG.GMAIL.CATEGORY.EVENT);
  if (hasTask) V4_addLabel_(thread, V4_CONFIG.GMAIL.CATEGORY.TASK);

  const needReview = plan.items.some(x => x.requiresReview);
  if (needReview) V4_addLabel_(thread, V4_CONFIG.GMAIL.STATUS.REVIEW);

  // suggestedLabels
  if (Array.isArray(plan.suggestedLabels)) {
    for (const s of plan.suggestedLabels.slice(0, 4)) {
      const seg = V4_sanitizeLabelSegment_(s);
      if (seg) V4_addLabel_(thread, `${V4_CONFIG.GMAIL.EXTRA_PREFIX}${seg}`);
    }
  }
}

// ============================================================
// Gmail 标签 / 线程选择 / 幂等
// ============================================================
function V4_ensureLabelsExist_() {
  V4_getOrCreateLabel_(V4_CONFIG.GMAIL.ROOT_LABEL);
  Object.values(V4_CONFIG.GMAIL.STATUS).forEach(V4_getOrCreateLabel_);
  Object.values(V4_CONFIG.GMAIL.CATEGORY).forEach(V4_getOrCreateLabel_);
  Object.values(V4_CONFIG.GMAIL.SYNC).forEach(V4_getOrCreateLabel_);
}

function V4_getOrCreateLabel_(name) {
  if (V4_CACHE.labels[name]) return V4_CACHE.labels[name];
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  V4_CACHE.labels[name] = label;
  return label;
}

function V4_addLabel_(thread, name) {
  const label = V4_getOrCreateLabel_(name);
  thread.addLabel(label);
}

function V4_removeLabel_(thread, name) {
  const label = GmailApp.getUserLabelByName(name);
  if (label) thread.removeLabel(label);
}

function V4_clearAiDerivedLabels_(thread) {
  const toRemove = new Set([
    ...Object.values(V4_CONFIG.GMAIL.STATUS).filter(n => n !== V4_CONFIG.GMAIL.STATUS.PROCESSING),
    ...Object.values(V4_CONFIG.GMAIL.CATEGORY),
    ...Object.values(V4_CONFIG.GMAIL.SYNC),
  ]);

  const labels = thread.getLabels();
  for (const l of labels) {
    const name = l.getName();
    if (toRemove.has(name)) {
      thread.removeLabel(l);
      continue;
    }
    if (name.startsWith(V4_CONFIG.GMAIL.EXTRA_PREFIX)) thread.removeLabel(l);
  }
}

function V4_sanitizeLabelSegment_(s) {
  let x = String(s || '').trim();
  if (!x) return '';
  x = x.replace(/[\/\\]/g, '／');
  x = x.replace(/[\[\]]/g, '');
  x = x.replace(/\s+/g, ' ');
  return x.substring(0, 30);
}

function V4_getNextThreadToProcess_() {
  const source = GmailApp.getUserLabelByName(V4_CONFIG.GMAIL.SOURCE_LABEL);
  if (!source) throw new Error(`找不到来源标签：${V4_CONFIG.GMAIL.SOURCE_LABEL}`);

  const threads = source.getThreads(0, V4_CONFIG.PIPELINE.MAX_THREADS_SCAN) || [];
  for (const thread of threads) {
    if (V4_threadHasLabel_(thread, V4_CONFIG.GMAIL.STATUS.PROCESSING)) continue;

    const messages = thread.getMessages();
    if (!messages || messages.length === 0) continue;

    const latestId = messages[messages.length - 1].getId();
    const last = V4_getLastProcessedMessageId_(thread.getId());
    if (last !== latestId) return thread;
  }
  return null;
}

function V4_threadHasLabel_(thread, labelName) {
  const labels = thread.getLabels();
  for (const l of labels) if (l.getName() === labelName) return true;
  return false;
}

function V4_getLastProcessedMessageId_(threadId) {
  return PropertiesService.getScriptProperties().getProperty(`V4_thread:${threadId}:lastMessageId`) || '';
}

function V4_setLastProcessedMessageId_(threadId, messageId) {
  PropertiesService.getScriptProperties().setProperty(`V4_thread:${threadId}:lastMessageId`, String(messageId));
}

// ============================================================
// 日志表（Log + Memory）
// ============================================================
function V4_ensureLogSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('V4_LOG_SHEET_ID');
  let ss;

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (_) { ssId = null; }
  }

  if (!ssId) {
    ss = SpreadsheetApp.create(V4_CONFIG.MEMORY.SPREADSHEET_NAME);
    ssId = ss.getId();
    props.setProperty('V4_LOG_SHEET_ID', ssId);
  }

  let sheet = ss.getSheetByName(V4_CONFIG.MEMORY.SHEET_TAB);
  if (!sheet) {
    const first = ss.getSheets()[0] || ss.insertSheet();
    first.setName(V4_CONFIG.MEMORY.SHEET_TAB);
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
      'permalink',
      'style',
      'overallCategory',
      'overallPriority',
      'itemIndex',
      'itemType',
      'requestedAction',
      'appliedAction',
      'memoryId',
      'fingerprint',
      'title',
      'startTime',
      'endTime',
      'dueTime',
      'allDay',
      'location',
      'calendarEventId',
      'taskId',
      'requiresReview',
      'confidence',
      'userFacingHeadline',
      'itemMemo',
      'assistantMemo',
      'rawJson',
      'error',
    ];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }

  V4_CACHE.log.ssId = ssId;
  V4_CACHE.log.sheet = sheet;
  return sheet;
}

function V4_appendLogRows_(email, plan, exec, timeZone) {
  const sheet = V4_ensureLogSheet_();
  const rows = [];
  const now = new Date();

  const perItem = exec && exec.perItem ? exec.perItem : [];
  const rawJson = JSON.stringify(plan.raw || plan).substring(0, 5000);

  for (let i = 0; i < perItem.length; i++) {
    const r = perItem[i];
    rows.push([
      now,
      email.threadId,
      email.messageId,
      email.subject,
      email.from,
      email.receivedAt,
      email.permalink,
      plan.selectedStyle,
      plan.classification.overallCategory,
      plan.classification.priority,
      r.itemIndex,
      r.itemType,
      r.requestedAction,
      r.appliedAction,
      r.memoryId,
      r.fingerprint,
      r.title,
      r.startTime,
      r.endTime,
      r.dueTime,
      r.allDay,
      r.location,
      r.calendarEventId,
      r.taskId,
      r.requiresReview,
      r.confidence,
      r.userFacingHeadline,
      r.itemMemo,
      r.assistantMemo,
      rawJson,
      r.error || '',
    ]);
  }

  if (rows.length === 0) {
    // 至少写一条“空动作”的记录，便于记忆连续性
    rows.push([
      now,
      email.threadId,
      email.messageId,
      email.subject,
      email.from,
      email.receivedAt,
      email.permalink,
      plan.selectedStyle,
      plan.classification.overallCategory,
      plan.classification.priority,
      0,
      'NOTE',
      'SKIP',
      'NOTE',
      '',
      '',
      '(no items)',
      '',
      '',
      '',
      false,
      '',
      '',
      '',
      false,
      0.6,
      plan.classification.reasoning || '',
      '',
      plan.assistantMemo || '',
      rawJson,
      '',
    ]);
  }

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function V4_getTodayLogItems_(now, timeZone) {
  const sheet = V4_ensureLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const startRow = Math.max(2, lastRow - V4_CONFIG.MEMORY.MAX_ROWS_READ + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  const todayStart = new Date(now);
  todayStart.setHours(0, 0, 0, 0);

  const items = [];
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const processedAt = row[0];
    if (!(processedAt instanceof Date)) continue;
    if (processedAt < todayStart) break;

    items.push({
      processedAt: Utilities.formatDate(processedAt, timeZone, 'HH:mm'),
      subject: row[3],
      from: row[4],
      overallCategory: row[8],
      overallPriority: row[9],
      itemType: row[11],
      action: row[13],
      title: row[16],
      startTime: row[17],
      dueTime: row[19],
      permalink: row[6],
      memo: row[27],
      headline: row[26],
      requiresReview: row[24],
    });

    if (items.length >= V4_CONFIG.DAILY_REPORT.MAX_ITEMS_IN_PROMPT) break;
  }

  return items.reverse();
}

function V4_callGeminiForDailyReport_(items, dateStr, timeZone) {
  const prompt = `
You are writing a daily report for an email automation assistant.
Output MUST be valid JSON matching the schema. No extra text.

Context:
- Date: ${dateStr}
- Timezone: ${timeZone}
- Total log items: ${items.length}

Available styles:
- TERSE: extremely concise
- BALANCED: structured, information-dense
- WARM: supportive, but still concise

Data (each item is one logged action / memo):
${JSON.stringify(items, null, 2)}

Write a Chinese report that balances personalization and redundancy:
1) 今日概览（1-2句话）
2) 重点事项（真正需要注意/行动的放前面；群发推广要合并压缩）
3) 日历/待办变更（新增/更新/取消分组）
4) AI 备忘录（你作为助手今天学到/观察到的模式，最多 5 条）
5) 待复核清单（requiresReview=true）

Return JSON:
{
  "plainText": "...",
  "htmlBody": "..."
}`;

  const genCfg = {
    responseMimeType: 'application/json',
    temperature: V4_CONFIG.DAILY_REPORT.TEMPERATURE,
    thinkingConfig: { thinkingLevel: V4_CONFIG.DAILY_REPORT.THINKING_LEVEL },
    responseSchema: {
      type: 'object',
      properties: {
        plainText: { type: 'string' },
        htmlBody: { type: 'string' },
      },
      required: ['plainText', 'htmlBody'],
    },
  };

  try {
    const out = V4_fetchGeminiJson_([{ text: prompt }], genCfg);
    if (out && typeof out.plainText === 'string' && typeof out.htmlBody === 'string') return out;
  } catch (e) {
    Logger.log(`V4 日报 AI 失败：${String(e)}`);
  }

  // Fallback
  const list = items.slice(0, 60).map(i => `- [${i.itemType}/${i.action}] ${i.title || i.subject}`).join('\n');
  return {
    plainText: `V4 智能日报 ${dateStr}\n\n共 ${items.length} 条记录。\n${list}`,
    htmlBody: `<h2>V4 智能日报 ${dateStr}</h2><p>共 <strong>${items.length}</strong> 条记录。</p><pre>${V4_escapeHtml_(list)}</pre>`,
  };
}

function V4_escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ============================================================
// 维护工具
// ============================================================
function V4_testProcessOne_() {
  V4_processEmails_();
}

function V4_resetProcessorState_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  for (const k of Object.keys(all)) {
    if (k.startsWith('V4_thread:') && k.endsWith(':lastMessageId')) props.deleteProperty(k);
  }
  Logger.log('✅ V4 已清空 thread:*:lastMessageId 状态');
}
