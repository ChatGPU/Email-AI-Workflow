/**
 * 配置区域
 */
const CONFIG = {

  MODEL_NAME: 'gemini-3-flash-preview', 
  API_KEY: '', // 替换你的 Key
  
  SOURCE_LABEL: 'PolyU',
  PROCESSED_LABEL: 'PolyU/Processed', // 确保 Gmail 里真的有这个子标签，或者只用 '[Processed]'
  
  // 免费版限制严格，一次处理 1-3 封即可，反正每 5 分钟会运行一次
  BATCH_SIZE: 1 
};

function processEmails() {
  const label = GmailApp.getUserLabelByName(CONFIG.SOURCE_LABEL);
  // 如果找不到子标签，代码会报错，建议先检查标签是否存在
  let processedLabel = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  
  // 如果处理标签不存在，自动创建（容错）
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(CONFIG.PROCESSED_LABEL);
  }

  // 只搜索未读邮件或特定标签，避免死循环
  const threads = label.getThreads(0, CONFIG.BATCH_SIZE);
  
  // 过滤掉已经处理过的 (双重保险，虽然 getThreads 本身只是取列表)
  const unprocessedThreads = threads.filter(t => !hasLabel(t, processedLabel));

  Logger.log(`本次运行发现 ${unprocessedThreads.length} 个未处理线程`);

  if (unprocessedThreads.length === 0) return;

  // 获取当前的准确时区和时间
  const timeZone = Session.getScriptTimeZone(); // 这里的时区取决于你 Script 项目的设置
  const now = new Date();
  
  for (const thread of unprocessedThreads) {
    // 速率限制保护：每次循环强制暂停 5 秒 (避免触发 Gemini Free Tier 的 15 RPM)
    Utilities.sleep(5000); 

    const msg = thread.getMessages()[thread.getMessageCount() - 1]; // 取最新一封
    const subject = msg.getSubject();
    const body = msg.getPlainBody().substring(0, 8000); // 截断过长邮件
    const date = msg.getDate();

    Logger.log(`>>> 正在处理: ${subject}`);

    try {
      const result = analyzeWithGemini(subject, body, date, now, timeZone);
      
      if (result && result.isEvent) {
        createCalendarEvent(result);
        Logger.log(`[成功] 日程已创建: ${result.title}`);
      } else {
        Logger.log(`[跳过] AI 判定这不是一个需要同步的日程`);
      }

      // 只有成功执行完才打标签
      thread.addLabel(processedLabel);

    } catch (e) {
      Logger.log(`[错误] 处理邮件失败: ${e.toString()}`);
      // 可以在这里发一封邮件给自己报错，或者跳过
    }
  }
}

// 辅助函数：检查是否有标签
function hasLabel(thread, labelObj) {
  const labels = thread.getLabels();
  for (let i = 0; i < labels.length; i++) {
    if (labels[i].getName() === labelObj.getName()) return true;
  }
  return false;
}

function analyzeWithGemini(subject, body, msgDate, now, timeZone) {
  const prompt = `
    Context:
    - Current User Time: ${Utilities.formatDate(now, timeZone, "yyyy-MM-dd HH:mm:ss")}
    - User Timezone: ${timeZone}
    - Email Date: ${msgDate}
    
    Task:
    Analyze the email content below. Identify if there is a specific calendar event, meeting, or deadline.
    
    Email Subject: ${subject}
    Email Body:
    ${body}

    Requirements:
    1. If NO event is found, set "isEvent": false.
    2. If an event IS found:
       - "isEvent": true
       - "title": A concise summary.
       - "startTime" / "endTime": ISO 8601 format INCLUDING TIMEZONE OFFSET (e.g. 2024-05-20T14:00:00+08:00). Crucial!
       - "description": Original email link or summary.
       - "location": Physical location or Zoom link.
    
    Output JSON only.
  `;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.MODEL_NAME}:generateContent?key=${CONFIG.API_KEY}`;
  
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "responseMimeType": "application/json" }
  };

  const response = UrlFetchApp.fetch(url, {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  });

  const json = JSON.parse(response.getContentText());
  
  if (json.error) throw new Error(json.error.message);
  
  const content = json.candidates[0].content.parts[0].text;
  
  // 清洗 JSON (防止 Gemini 返回 ```json ... ```)
  const cleanJson = content.replace(/```json/g, "").replace(/```/g, "").trim();
  
  return JSON.parse(cleanJson);
}

function createCalendarEvent(data) {
  const cal = CalendarApp.getDefaultCalendar();
  const desc = `${data.description}\n\n[AI Sync]`;
  
  // 直接使用 ISO 字符串，new Date() 会自动处理 +08:00 这种偏移量
  const start = new Date(data.startTime);
  const end = new Date(data.endTime);

  if (data.allDay) {
    cal.createAllDayEvent(`[Outlook] ${data.title}`, start, {description: desc, location: data.location});
  } else {
    // 校验时间有效性
    if (isNaN(start.getTime())) throw new Error("无效的开始时间格式");
    
    cal.createEvent(`[Outlook] ${data.title}`, start, end, {description: desc, location: data.location});
  }

}
