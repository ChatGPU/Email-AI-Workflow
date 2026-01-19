/**
 * SmartEmailProcessor v3
 * 一个懂你的邮件助手
 * 
 * 为 PolyU LMS PhD 定制
 */

const CONFIG = {
  VERSION: '3.0.0',

  // ============ 用户画像 ============
  USER: {
    IDENTITY: '一位PolyU LMS (Logistics and Maritime Studies) 系的 PhD 学生，正在忙碌的科研生活中努力前行',
    FOCUS_AREAS: [
      '学术研究相关（会议、论文、seminar、workshop）',
      '导师/实验室相关',
      'PhD 专属活动和 deadline',
      '奖学金、资助、报销',
      '签证、住宿等行政事务',
    ],
    DEPRIORITIZE: [
      '本科生/授课型硕士专属活动',
      '全校性大型推广（与科研无关）',
      '学生会、社团招新（除非与专业相关）',
    ],
  },

  // ============ Gemini ============
  GEMINI: {
    API_KEY: 'YOUR_API_KEY_HERE',
    MODEL_NAME: 'gemini-3-flash-preview',
    API_ROOT: 'https://generativelanguage.googleapis.com/v1beta/models',
    THINKING_LEVEL: 'HIGH',
    MEDIA_RESOLUTION: 'MEDIA_RESOLUTION_HIGH',
  },

  // ============ Gmail 标签（精简版）============
  GMAIL: {
    SOURCE_LABEL: 'PolyU',
    
    // 简洁的标签体系
    LABELS: {
      ROOT: 'AI',
      
      // 分类（只保留核心）
      EVENT: 'AI/日程',
      TASK: 'AI/待办',
      INFO: 'AI/已阅',
      
      // 状态
      REVIEW: 'AI/请检查',
      ERROR: 'AI/处理失败',
      
      // 同步状态
      SYNCED_CAL: 'AI/已同步日历',
      SYNCED_TASK: 'AI/已同步待办',
    },
  },

  // ============ 处理参数 ============
  PROCESSING: {
    MAX_THREADS_SCAN: 30,
    MAX_BODY_CHARS: 22000,
    MAX_HTML_SNIPPET_CHARS: 5000,
    MAX_LINKS: 40,
    MAX_MEDIA_ITEMS: 8,
    MAX_IMAGE_ITEMS: 5,
    MAX_PDF_ITEMS: 2,
    MAX_MEDIA_BYTES_EACH: 6 * 1024 * 1024,
    MAX_TOTAL_MEDIA_BYTES: 14 * 1024 * 1024,
    MIN_IMAGE_BYTES: 8 * 1024,
    
    DRY_RUN: false,
  },

  // ============ 日报 ============
  DAILY_REPORT: {
    ENABLED: true,
    RECIPIENT_EMAIL: 'heibaiyouji@gmail.com',
    HOUR: 22,
    LOG_SPREADSHEET_NAME: 'SmartEmailProcessor Log',
    LOG_SHEET_TAB: 'log',
  },
};

// ============ 缓存 ============
const CACHE = { labels: {}, log: { ssId: null, sheet: null } };

/**
 * ========================================
 * 初始化与设置
 * ========================================
 */
function setupSmartEmailProcessor() {
  ensureLabelsExist_();
  ensureLogSheet_();
  setupTriggers();
  Logger.log('✅ 初始化完成');
}

function setGeminiApiKey(apiKey) {
  const key = String(apiKey || '').trim();
  if (key.length < 20) throw new Error('API Key 太短了');
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('✅ API Key 已保存');
}

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
 * ========================================
 * 主流程
 * ========================================
 */
function processEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('⏭️ 上一次还在跑');
    return;
  }

  let thread = null;

  try {
    ensureLabelsExist_();
    ensureLogSheet_();

    thread = getNextThreadToProcess_();
    if (!thread) {
      Logger.log('✓ 没有新邮件');
      return;
    }

    const messages = thread.getMessages();
    const latest = messages[messages.length - 1];
    const threadId = thread.getId();
    const latestMessageId = latest.getId();
    const tz = Session.getScriptTimeZone();
    const now = new Date();

    // 幂等检查
    const lastProcessed = getLastProcessedMessageId_(threadId);
    if (lastProcessed === latestMessageId) {
      Logger.log(`⏭️ 已处理过：${latest.getSubject()}`);
      return;
    }

    addLabel_(thread, CONFIG.GMAIL.LABELS.ROOT);
    clearPreviousLabels_(thread);

    const email = extractEmailData_(thread, messages, tz);
    const aiResult = callGeminiForEmail_(email, tz, now);
    const result = normalizeResult_(aiResult, email);

    const exec = applyActions_(thread, email, result, tz);

    appendLogRow_(email, result, exec);
    setLastProcessedMessageId_(threadId, latestMessageId);

    Logger.log(`✅ ${email.subject}`);
  } catch (e) {
    Logger.log(`❌ ${e.stack || e}`);
    if (thread) {
      try { addLabel_(thread, CONFIG.GMAIL.LABELS.ERROR); } catch (_) {}
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * ========================================
 * 数据提取
 * ========================================
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
    mediaParts: media.parts,
    mediaManifest: media.manifest,
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
  ];

  let cutIndex = null;
  for (const re of cutMarkers) {
    re.lastIndex = 0;
    const m = re.exec(t);
    if (m && m.index > 300) cutIndex = cutIndex === null ? m.index : Math.min(cutIndex, m.index);
  }
  if (cutIndex !== null) t = t.substring(0, cutIndex);

  const sigMarkers = [/\n--\s*\n/, /\nSent from my (iPhone|iPad|Android).*/i];
  for (const re of sigMarkers) {
    const m = re.exec(t);
    if (m && m.index > 200) t = t.substring(0, m.index);
  }

  return t.trim();
}

function extractLinks_(html, plain, maxLinks) {
  const results = [];
  const seen = new Set();

  function add(url, text, source) {
    const cleaned = normalizeUrl_(url);
    if (!cleaned || !isValidUrl_(cleaned)) return;
    if (seen.has(cleaned)) return;
    seen.add(cleaned);

    const anchorText = (text || '').trim();
    const domain = extractDomain_(cleaned);
    const type = classifyLinkType_(cleaned, anchorText);
    const score = scoreLink_(cleaned, anchorText, type);

    results.push({ url: cleaned, text: anchorText || domain, domain, type, score, source });
  }

  const anchorRe = /<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = anchorRe.exec(html || '')) !== null && results.length < maxLinks * 3) {
    const rawText = stripHtmlTags_(m[2] || '');
    add(m[1], decodeHtmlEntities_(rawText).trim(), 'HTML');
  }

  const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi;
  const combined = `${plain || ''}\n${stripHtmlTags_(html || '')}`;
  while ((m = urlRe.exec(combined)) !== null && results.length < maxLinks * 4) {
    add(m[0], '', 'BARE');
  }

  results.sort((a, b) => b.score - a.score);
  return results.slice(0, maxLinks);
}

function normalizeUrl_(url) {
  if (!url) return '';
  let u = String(url).trim().replace(/^<|>$/g, '').replace(/[)\].,;:!?"']+$/g, '');
  return u;
}

function isValidUrl_(url) {
  if (!url) return false;
  const u = String(url).trim();
  if (!/^https?:\/\//i.test(u)) return false;
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

  if (u.includes('unsubscribe') || t.includes('取消订阅')) return 'UNSUBSCRIBE';
  if (u.includes('zoom.') || u.includes('teams.microsoft') || u.includes('meet.google')) return 'MEETING';
  if (u.includes('calendar') || u.endsWith('.ics')) return 'CALENDAR';
  if (u.includes('pay') || u.includes('invoice')) return 'PAYMENT';
  if (u.includes('docs.google') || u.includes('drive.google') || u.includes('dropbox')) return 'DOCUMENT';
  if (u.includes('register') || u.includes('signup') || t.includes('报名')) return 'REGISTRATION';
  if (u.includes('facebook.com') || u.includes('twitter.com') || u.includes('linkedin.com')) return 'SOCIAL';
  return 'GENERAL';
}

function scoreLink_(url, text, type) {
  const baseByType = {
    MEETING: 100, PAYMENT: 95, DOCUMENT: 85, CALENDAR: 80, REGISTRATION: 70,
    GENERAL: 30, SOCIAL: 5, UNSUBSCRIBE: -50,
  };
  let score = baseByType[type] != null ? baseByType[type] : 20;

  const goodKw = ['join', 'register', 'ticket', 'verify', 'confirm', 'download', '报名', '注册'];
  const u = String(url).toLowerCase();
  const t = String(text || '').toLowerCase();
  for (const kw of goodKw) if (t.includes(kw) || u.includes(kw)) { score += 6; break; }

  if (u.includes('utm_') || u.includes('tracking')) score -= 20;
  if (text && text.trim().length >= 8) score += 5;
  return score;
}

function stripHtmlTags_(html) {
  return String(html || '').replace(/<[^>]*>/g, ' ');
}

function decodeHtmlEntities_(text) {
  return String(text || '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'");
}

function extractKeyPatterns_(text) {
  const t = String(text || '');
  const patterns = { referenceNumbers: [], codes: [], dates: [], times: [], amounts: [], emails: [], phones: [] };

  const pushUniq = (arr, v) => {
    const val = String(v || '').trim();
    if (val && arr.indexOf(val) === -1) arr.push(val);
  };

  let m;
  const emailRe = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
  while ((m = emailRe.exec(t)) !== null) pushUniq(patterns.emails, m[0]);

  const amountRe = /(?:HK\$|\$|¥|£|€)\s*[\d,]+(?:\.\d+)?/g;
  while ((m = amountRe.exec(t)) !== null) pushUniq(patterns.amounts, m[0]);

  const refRe = /(?:(?:order|confirmation|reference|booking|ticket|ref|no|#|订单|确认|编号)\s*[:#：]?\s*)([A-Z0-9][A-Z0-9\-]{5,30})/gi;
  while ((m = refRe.exec(t)) !== null) pushUniq(patterns.referenceNumbers, m[1]);

  const dateRe = /\b(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/g;
  while ((m = dateRe.exec(t)) !== null) pushUniq(patterns.dates, m[1]);

  const timeRe = /\b(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\b/g;
  while ((m = timeRe.exec(t)) !== null) pushUniq(patterns.times, m[1]);

  return patterns;
}

function extractConversationContext_(previousMessages, timeZone) {
  if (!previousMessages || previousMessages.length === 0) return null;
  const last = previousMessages.slice(-2);
  return last.map(msg => ({
    from: msg.getFrom() || '',
    date: Utilities.formatDate(msg.getDate(), timeZone, 'yyyy-MM-dd HH:mm'),
    snippet: cleanPlainBody_(msg.getPlainBody() || '').substring(0, 600),
  }));
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
      imageCandidates.push({ ...base, score: scoreImage_(name, size), kind: 'image' });
    } else if (mimeType === 'application/pdf' || /\.pdf$/i.test(name)) {
      pdfCandidates.push({ ...base, score: scorePdf_(name, size), kind: 'pdf' });
    } else {
      otherAttachments.push({ filename: name, mimeType, sizeKB: Math.round(size / 1024) });
    }
  }

  imageCandidates.sort((a, b) => b.score - a.score);
  pdfCandidates.sort((a, b) => b.score - a.score);

  const selected = [];
  for (const c of imageCandidates) {
    if (selected.filter(x => x.kind === 'image').length >= CONFIG.PROCESSING.MAX_IMAGE_ITEMS) break;
    if (c.size < CONFIG.PROCESSING.MIN_IMAGE_BYTES || c.size > CONFIG.PROCESSING.MAX_MEDIA_BYTES_EACH) continue;
    selected.push(c);
    if (selected.length >= CONFIG.PROCESSING.MAX_MEDIA_ITEMS) break;
  }
  for (const c of pdfCandidates) {
    if (selected.filter(x => x.kind === 'pdf').length >= CONFIG.PROCESSING.MAX_PDF_ITEMS) break;
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

    manifest.push({ index, kind: item.kind, filename: item.name, sizeKB: Math.round(item.size / 1024) });
    parts.push({ text: `[附件 #${index}: ${item.kind}, ${item.name}]` });
    parts.push({ inlineData: { mimeType: item.mimeType, data: Utilities.base64Encode(item.att.getBytes()) } });
  }

  return { parts, manifest, otherAttachments };
}

function scoreImage_(name, size) {
  const n = String(name || '').toLowerCase();
  let score = 30;
  const strong = ['qr', 'ticket', 'receipt', 'invoice', '二维码', '票', '凭证'];
  for (const kw of strong) if (n.includes(kw)) score += 80;
  const noise = ['logo', 'icon', 'banner', 'facebook', 'twitter'];
  for (const kw of noise) if (n.includes(kw)) score -= 60;
  if (size >= 200 * 1024 && size <= 1500 * 1024) score += 20;
  if (size < 20 * 1024) score -= 40;
  return score;
}

function scorePdf_(name, size) {
  const n = String(name || '').toLowerCase();
  let score = 50;
  const strong = ['invoice', 'receipt', 'ticket', 'statement', '发票', '收据'];
  for (const kw of strong) if (n.includes(kw)) score += 80;
  if (size >= 200 * 1024 && size <= 5 * 1024 * 1024) score += 20;
  return score;
}

/**
 * ========================================
 * AI 调用
 * ========================================
 */
function callGeminiForEmail_(email, timeZone, now) {
  const prompt = buildPrompt_(email, timeZone, now);
  const parts = [{ text: prompt }];

  if (email.mediaParts && email.mediaParts.length > 0) {
    parts.push(...email.mediaParts);
  }

  const genCfg = {
    responseMimeType: 'application/json',
    thinkingConfig: { thinkingLevel: CONFIG.GEMINI.THINKING_LEVEL },
    responseSchema: getResponseSchema_(),
  };

  if (CONFIG.GEMINI.MEDIA_RESOLUTION) {
    genCfg.mediaResolution = CONFIG.GEMINI.MEDIA_RESOLUTION;
  }

  return fetchGeminiJson_(parts, genCfg);
}

function buildPrompt_(email, timeZone, now) {
  const linksText = (email.extractedLinks || [])
    .slice(0, 15)
    .map(l => `- [${l.type}] "${l.text}" → ${l.url}`)
    .join('\n');

  const patterns = email.extractedPatterns || {};

  return `
你是一位温暖、聪明的私人助理。你的主人是 ${CONFIG.USER.IDENTITY}。

你的任务是帮 ta 处理这封邮件，但不只是机械地提取信息——你要像一个真正关心 ta 的朋友，用人话告诉 ta 这封邮件说了什么、需不需要行动、有什么建议。

═══════════════════════════════════
关于你的主人
═══════════════════════════════════
身份：${CONFIG.USER.IDENTITY}

ta 关心的事：
${CONFIG.USER.FOCUS_AREAS.map(x => '• ' + x).join('\n')}

可以弱化的内容（但别直接忽略，如果真的重要还是要提）：
${CONFIG.USER.DEPRIORITIZE.map(x => '• ' + x).join('\n')}

═══════════════════════════════════
当前时间
═══════════════════════════════════
${Utilities.formatDate(now, timeZone, 'yyyy年M月d日 EEEE HH:mm')}
时区：${timeZone}

═══════════════════════════════════
这封邮件
═══════════════════════════════════
主题：${email.subject}
发件人：${email.from}
收到时间：${email.receivedAtStr}
线程消息数：${email.threadMessageCount}

【正文】
${email.body}

【HTML 片段（可能有格式线索）】
${email.htmlSnippet}

【提取的链接】
${linksText || '(无)'}

【识别到的模式】
${JSON.stringify(patterns, null, 2)}

【附件】
${email.mediaManifest.length ? JSON.stringify(email.mediaManifest) : '(无图片/PDF)'}
${email.otherAttachments.length ? '其他附件: ' + JSON.stringify(email.otherAttachments) : ''}

${email.conversationContext ? '【对话上下文】\n' + JSON.stringify(email.conversationContext, null, 2) : ''}

═══════════════════════════════════
你的任务
═══════════════════════════════════

1. **理解这封邮件**：用一两句话，像朋友聊天一样告诉 ta 这封邮件是关于什么的。

2. **分类**：
   - EVENT：有明确的日期+时间，需要 ta 出席或参加（会议、培训、答辩、seminar）
   - TASK：需要 ta 做点什么，但没有固定时间段（提交材料、回复邮件、报销）
   - INFO：知道就好，不需要特别行动（通知、newsletter、推广）

3. **评估相关性**：
   - 这封邮件和 ta 作为 LMS PhD 的身份相关吗？
   - 是全校性的还是针对研究生的？
   - 是否需要特别关注？

4. **提取关键信息**：
   - 如果是 EVENT：什么时候、在哪里、要准备什么
   - 如果是 TASK：deadline 是什么、要做什么
   - 有没有重要的编号、链接、二维码

5. **给出建议**：
   - 用温暖但简洁的语气
   - 如果是重要的事，可以稍微叮嘱一下
   - 如果是好消息，可以替 ta 开心
   - 如果看起来可以忽略，直接说"这个可以先放着"

═══════════════════════════════════
输出格式（严格 JSON）
═══════════════════════════════════

{
  "category": "EVENT | TASK | INFO",
  "relevance": "HIGH | MEDIUM | LOW",  // 对 PhD 的相关性
  "needsReview": true/false,  // 如果你不太确定，设为 true
  
  "summary": "用一两句话像朋友聊天一样说明这封邮件是关于什么的",
  
  "title": "简短的标题（用于日历或待办）",
  
  "when": {
    "display": "用自然语言描述时间，比如'1月28日下午2:45'或'下周五之前'",
    "startTime": "ISO 8601 格式，如果有的话",
    "endTime": "ISO 8601 格式，如果有的话",
    "deadline": "ISO 8601 格式，如果是 deadline 的话",
    "allDay": true/false
  },
  
  "where": "地点或会议链接，如果有的话",
  
  "keyInfo": {
    "numbers": ["重要的编号、confirmation code 等"],
    "links": [
      {"label": "链接的用途", "url": "实际链接"}
    ],
    "fromImages": "从图片/PDF中发现的重要信息，比如二维码内容"
  },
  
  "advice": "给 ta 的建议和下一步行动，用温暖的口吻，可以是一段话或几个要点",
  
  "note": "任何你觉得值得一提的额外观察"
}

记住：你是在帮一个忙碌的 PhD 学生，ta 需要的是清晰、温暖、有用的信息，不是机械的数据罗列。
`;
}

function getResponseSchema_() {
  return {
    type: 'object',
    properties: {
      category: { type: 'string', enum: ['EVENT', 'TASK', 'INFO'] },
      relevance: { type: 'string', enum: ['HIGH', 'MEDIUM', 'LOW'] },
      needsReview: { type: 'boolean' },
      summary: { type: 'string' },
      title: { type: 'string' },
      when: {
        type: 'object',
        properties: {
          display: { type: 'string' },
          startTime: { type: 'string', nullable: true },
          endTime: { type: 'string', nullable: true },
          deadline: { type: 'string', nullable: true },
          allDay: { type: 'boolean' },
        },
      },
      where: { type: 'string', nullable: true },
      keyInfo: {
        type: 'object',
        properties: {
          numbers: { type: 'array', items: { type: 'string' } },
          links: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                label: { type: 'string' },
                url: { type: 'string' },
              },
            },
          },
          fromImages: { type: 'string', nullable: true },
        },
      },
      advice: { type: 'string' },
      note: { type: 'string', nullable: true },
    },
    required: ['category', 'relevance', 'needsReview', 'summary', 'title', 'advice'],
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
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  let json;
  try { json = JSON.parse(text); } catch (_) {
    throw new Error(`Gemini 返回非 JSON（HTTP ${code}）`);
  }

  if (json.error) throw new Error(json.error.message || JSON.stringify(json.error));

  const outText = json?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!outText) throw new Error('Gemini 返回格式异常');

  try { return JSON.parse(outText); } catch (e) {
    throw new Error(`输出无法解析为 JSON: ${outText.slice(0, 500)}`);
  }
}

function getApiKey_() {
  const fromConfig = String(CONFIG.GEMINI.API_KEY || '').trim();
  if (fromConfig && fromConfig !== 'YOUR_API_KEY_HERE') return fromConfig;
  const fromProps = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (fromProps && fromProps.trim()) return fromProps.trim();
  throw new Error('请先设置 Gemini API Key');
}

/**
 * ========================================
 * 结果处理
 * ========================================
 */
function normalizeResult_(raw, email) {
  const r = raw && typeof raw === 'object' ? raw : {};
  
  const category = ['EVENT', 'TASK', 'INFO'].includes(r.category) ? r.category : 'TASK';
  const relevance = ['HIGH', 'MEDIUM', 'LOW'].includes(r.relevance) ? r.relevance : 'MEDIUM';
  const needsReview = typeof r.needsReview === 'boolean' ? r.needsReview : true;
  
  const when = r.when && typeof r.when === 'object' ? r.when : {};
  const keyInfo = r.keyInfo && typeof r.keyInfo === 'object' ? r.keyInfo : {};

  return {
    category,
    relevance,
    needsReview,
    summary: String(r.summary || '').trim().substring(0, 500),
    title: String(r.title || email.subject || '邮件事项').trim().substring(0, 80),
    when: {
      display: String(when.display || '').trim(),
      startTime: when.startTime ? String(when.startTime).trim() : null,
      endTime: when.endTime ? String(when.endTime).trim() : null,
      deadline: when.deadline ? String(when.deadline).trim() : null,
      allDay: typeof when.allDay === 'boolean' ? when.allDay : false,
    },
    where: r.where ? String(r.where).trim().substring(0, 200) : null,
    keyInfo: {
      numbers: Array.isArray(keyInfo.numbers) ? keyInfo.numbers.filter(Boolean).slice(0, 10) : [],
      links: Array.isArray(keyInfo.links) ? keyInfo.links.slice(0, 5) : [],
      fromImages: keyInfo.fromImages ? String(keyInfo.fromImages).trim() : null,
    },
    advice: String(r.advice || '').trim().substring(0, 2000),
    note: r.note ? String(r.note).trim().substring(0, 500) : null,
  };
}

function applyActions_(thread, email, result, timeZone) {
  // 打标签
  const catLabel = CONFIG.GMAIL.LABELS[result.category];
  if (catLabel) addLabel_(thread, catLabel);
  if (result.needsReview) addLabel_(thread, CONFIG.GMAIL.LABELS.REVIEW);

  const notes = buildNotes_(result, email);
  const exec = { action: 'NONE', calendarEventId: '', taskId: '' };

  if (CONFIG.PROCESSING.DRY_RUN) return exec;

  // EVENT → Calendar
  if (result.category === 'EVENT') {
    const start = parseDateTimeSafe_(result.when.startTime);
    if (start) {
      try {
        const end = parseDateTimeSafe_(result.when.endTime) || new Date(start.getTime() + 60 * 60 * 1000);
        const eventId = createCalendarEvent_(result, start, end, notes);
        exec.action = 'CALENDAR';
        exec.calendarEventId = eventId;
        addLabel_(thread, CONFIG.GMAIL.LABELS.SYNCED_CAL);
        return exec;
      } catch (e) {
        Logger.log(`日历创建失败: ${e}`);
      }
    }
  }

  // TASK 或没有时间的 EVENT → Tasks
  if (result.category === 'TASK' || (result.category === 'EVENT' && !result.when.startTime)) {
    try {
      const taskId = createGoogleTask_(result, notes);
      exec.action = 'TASKS';
      exec.taskId = taskId;
      addLabel_(thread, CONFIG.GMAIL.LABELS.SYNCED_TASK);
    } catch (e) {
      Logger.log(`Tasks 创建失败: ${e}`);
    }
  }

  return exec;
}

function buildNotes_(result, email) {
  const lines = [];

  // 开头：温暖的总结
  lines.push(result.summary);
  lines.push('');

  // 时间（如果有）
  if (result.when.display) {
    lines.push(`🕐 ${result.when.display}`);
  }

  // 地点（如果有）
  if (result.where) {
    lines.push(`📍 ${result.where}`);
  }

  // 重要编号
  if (result.keyInfo.numbers && result.keyInfo.numbers.length) {
    lines.push('');
    lines.push(`📋 重要编号：${result.keyInfo.numbers.join('、')}`);
  }

  // 图片信息
  if (result.keyInfo.fromImages) {
    lines.push('');
    lines.push(`🖼️ ${result.keyInfo.fromImages}`);
  }

  // 重要链接
  if (result.keyInfo.links && result.keyInfo.links.length) {
    lines.push('');
    lines.push('🔗 相关链接：');
    for (const l of result.keyInfo.links) {
      if (l && l.url) lines.push(`   • ${l.label || '链接'}：${l.url}`);
    }
  }

  // 建议
  lines.push('');
  lines.push('💡 ' + result.advice);

  // 额外备注
  if (result.note) {
    lines.push('');
    lines.push(`📝 ${result.note}`);
  }

  // 来源
  lines.push('');
  lines.push('───────────');
  lines.push(`来自：${email.from}`);
  lines.push(`原邮件：${email.permalink}`);

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
  const title = result.title;

  if (result.when.allDay) {
    const startDate = new Date(start);
    startDate.setHours(0, 0, 0, 0);
    const endDate = new Date(end);
    endDate.setHours(0, 0, 0, 0);
    if (endDate.getTime() <= startDate.getTime()) endDate.setDate(endDate.getDate() + 1);

    const ev = cal.createAllDayEvent(title, startDate, endDate, {
      location: result.where || '',
      description: notes,
    });
    ev.addPopupReminder(60);
    return ev.getId();
  }

  const ev = cal.createEvent(title, start, end, {
    location: result.where || '',
    description: notes,
  });
  ev.addPopupReminder(30);
  ev.addPopupReminder(1440); // 提前一天
  return ev.getId();
}

function createGoogleTask_(result, notes) {
  if (typeof Tasks === 'undefined' || !Tasks.Tasks) {
    throw new Error('请启用 Tasks API');
  }

  const task = {
    title: result.title,
    notes: String(notes || '').substring(0, 8000),
  };

  const due = result.when.deadline || result.when.startTime;
  if (due) {
    const d = new Date(due);
    if (!Number.isNaN(d.getTime())) task.due = d.toISOString();
  }

  const inserted = Tasks.Tasks.insert(task, '@default');
  return inserted && inserted.id ? inserted.id : '';
}

/**
 * ========================================
 * 标签管理
 * ========================================
 */
function ensureLabelsExist_() {
  for (const name of Object.values(CONFIG.GMAIL.LABELS)) {
    getOrCreateLabel_(name);
  }
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

function clearPreviousLabels_(thread) {
  const toClear = [
    CONFIG.GMAIL.LABELS.EVENT,
    CONFIG.GMAIL.LABELS.TASK,
    CONFIG.GMAIL.LABELS.INFO,
    CONFIG.GMAIL.LABELS.REVIEW,
    CONFIG.GMAIL.LABELS.SYNCED_CAL,
    CONFIG.GMAIL.LABELS.SYNCED_TASK,
  ];
  const labels = thread.getLabels();
  for (const l of labels) {
    if (toClear.includes(l.getName())) thread.removeLabel(l);
  }
}

/**
 * ========================================
 * 线程选择与幂等
 * ========================================
 */
function getNextThreadToProcess_() {
  const source = GmailApp.getUserLabelByName(CONFIG.GMAIL.SOURCE_LABEL);
  if (!source) throw new Error(`找不到来源标签：${CONFIG.GMAIL.SOURCE_LABEL}`);

  const threads = source.getThreads(0, CONFIG.PROCESSING.MAX_THREADS_SCAN) || [];
  for (const thread of threads) {
    const messages = thread.getMessages();
    if (!messages || messages.length === 0) continue;

    const latestId = messages[messages.length - 1].getId();
    const last = getLastProcessedMessageId_(thread.getId());
    if (last !== latestId) return thread;
  }
  return null;
}

function getLastProcessedMessageId_(threadId) {
  return PropertiesService.getScriptProperties().getProperty(`t:${threadId}`) || '';
}

function setLastProcessedMessageId_(threadId, messageId) {
  PropertiesService.getScriptProperties().setProperty(`t:${threadId}`, String(messageId));
}

/**
 * ========================================
 * 日志
 * ========================================
 */
function ensureLogSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('LOG_SHEET_ID');
  let ss;

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (_) { ssId = null; }
  }

  if (!ssId) {
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
    sheet.getRange(1, 1, 1, 12).setValues([[
      'processedAt', 'subject', 'from', 'category', 'relevance',
      'needsReview', 'action', 'title', 'summary', 'advice', 'permalink', 'rawJson'
    ]]);
    sheet.setFrozenRows(1);
  }

  CACHE.log.ssId = ssId;
  CACHE.log.sheet = sheet;
  return sheet;
}

function appendLogRow_(email, result, exec) {
  const sheet = ensureLogSheet_();
  const row = [
    new Date(),
    email.subject,
    email.from,
    result.category,
    result.relevance,
    result.needsReview,
    exec.action || 'NONE',
    result.title,
    result.summary,
    result.advice,
    email.permalink,
    JSON.stringify(result).substring(0, 5000),
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

/**
 * ========================================
 * 每日日报
 * ========================================
 */
function sendDailyReport() {
  if (!CONFIG.DAILY_REPORT.ENABLED) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const tz = Session.getScriptTimeZone();
    const today = new Date();
    const dateStr = Utilities.formatDate(today, tz, 'yyyy年M月d日 EEEE');

    const items = getTodayLogEntries_();
    if (items.length === 0) {
      Logger.log('今日无邮件，跳过日报');
      return;
    }

    const summary = callGeminiForDailySummary_(items, dateStr, tz);

    const subject = `📬 今日邮件小结 · ${dateStr}`;
    GmailApp.sendEmail(CONFIG.DAILY_REPORT.RECIPIENT_EMAIL, subject, summary.plainText, {
      htmlBody: summary.htmlBody,
    });

    Logger.log('✅ 日报已发送');
  } catch (e) {
    Logger.log(`❌ 日报失败: ${e}`);
  } finally {
    lock.releaseLock();
  }
}

function getTodayLogEntries_() {
  const sheet = ensureLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const startRow = Math.max(2, lastRow - 200);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 12).getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const items = [];
  for (const row of values) {
    const processedAt = row[0];
    if (!(processedAt instanceof Date) || processedAt < today) continue;

    items.push({
      subject: row[1],
      from: row[2],
      category: row[3],
      relevance: row[4],
      action: row[6],
      title: row[7],
      summary: row[8],
      advice: row[9],
    });
  }
  return items;
}

function callGeminiForDailySummary_(items, dateStr, timeZone) {
  const prompt = `
你是一位温暖的私人助理，要为你的主人（一位 ${CONFIG.USER.IDENTITY}）写一份今日邮件小结。

今天是 ${dateStr}，共处理了 ${items.length} 封邮件。

以下是今天处理的邮件：
${JSON.stringify(items, null, 2)}

请写一份简短、温暖的日报，包含：

1. **今日一览**：一两句话总结今天的邮件情况
2. **重要事项**：如果有需要特别注意的（relevance=HIGH 或 category=EVENT/TASK），列出来并给出简短提醒
3. **明天的你**：如果有即将到来的 deadline 或活动，友善地提醒一下
4. **一句话**：可以是鼓励、提醒劳逸结合，或者只是一句温暖的话

口吻要像朋友聊天，不要太正式。如果今天都是些无关紧要的邮件，可以轻松地说"今天没什么大事"。

输出 JSON：
{
  "plainText": "纯文本版本",
  "htmlBody": "HTML 版本（简洁美观，避免用 emoji）"
}
`;

  const genCfg = {
    responseMimeType: 'application/json',
    thinkingConfig: { thinkingLevel: 'HIGH' },
  };

  try {
    const out = fetchGeminiJson_([{ text: prompt }], genCfg);
    if (out && out.plainText && out.htmlBody) return out;
  } catch (e) {
    Logger.log(`日报 AI 失败: ${e}`);
  }

  // Fallback
  const list = items.map(i => `• [${i.category}] ${i.title}`).join('\n');
  return {
    plainText: `${dateStr} 邮件小结\n\n今天处理了 ${items.length} 封邮件：\n${list}`,
    htmlBody: `<h2>${dateStr} 邮件小结</h2><p>今天处理了 <strong>${items.length}</strong> 封邮件。</p><ul>${items.map(i => `<li>[${i.category}] ${i.title}</li>`).join('')}</ul>`,
  };
}

/**
 * ========================================
 * 工具函数
 * ========================================
 */
function testProcessOne() {
  processEmails();
}

function resetAllState() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  for (const k of Object.keys(all)) {
    if (k.startsWith('t:')) props.deleteProperty(k);
  }
  Logger.log('✅ 已重置所有处理状态');
}

function deleteAllAiLabels() {
  const prefixes = ['AI/', '[AI]/', '[AI]'];
  const allLabels = GmailApp.getUserLabels();
  let count = 0;
  
  for (const label of allLabels) {
    const name = label.getName();
    for (const p of prefixes) {
      if (name === p.replace('/', '') || name.startsWith(p)) {
        try {
          label.deleteLabel();
          count++;
          Logger.log(`删除: ${name}`);
        } catch (e) {
          Logger.log(`无法删除 ${name}: ${e}`);
        }
        break;
      }
    }
  }
  
  Logger.log(`✅ 共删除 ${count} 个标签`);
}
