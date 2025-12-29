function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Job Tracker')
    .addItem('📱 Get Mobile App Link', 'showDeployInstructions')
    .addToUi();
}

function showDeployInstructions() {
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; padding: 10px;">
      <h3>How to use on your Phone/Watch</h3>
      <p>To use the Job Tracker App, you need to create your own private link:</p>
      <ol>
        <li>Go to <b>Extensions > Apps Script</b> in the menu above.</li>
        <li>In the new tab, click the blue <b>Deploy</b> button (top right).</li>
        <li>Select <b>New Deployment</b>.</li>
        <li>Click the "Select type" gear icon ⚙️ and choose <b>Web App</b>.</li>
        <li><b>IMPORTANT:</b> Set "Execute as" to <b>User accessing the web app</b>.</li>
        <li>Set "Who has access" to <b>Anyone with Google Account</b> (or Only Myself).</li>
        <li>Click <b>Deploy</b>.</li>
        <li>Copy the <b>Web App URL</b> and send it to your phone!</li>
      </ol>
    </div>
  `).setWidth(400).setHeight(400);
  ui.showModalDialog(htmlOutput, '🚀 Launch Instructions');
}

const SHEET_NAME = 'JAT'; 
const GEMINI_MODEL = 'gemini-2.0-flash-exp'; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Job Tracker Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Saves the User's API Key securely.
 * Includes validation to ensure it looks like a real Google Cloud key.
 */
function saveUserApiKey(key) {
  const cleanKey = key ? key.trim() : "";
  if (!cleanKey || !cleanKey.startsWith("AIza") || cleanKey.length < 30) {
    throw new Error("Invalid API Key format. Google API keys usually start with 'AIza'.");
  }
  PropertiesService.getUserProperties().setProperty('USER_GEMINI_KEY', cleanKey);
  return { success: true, message: "API Key saved securely!" };
}

function checkHasApiKey() {
  const key = PropertiesService.getUserProperties().getProperty('USER_GEMINI_KEY');
  return { hasKey: !!key };
}

function getApiKeyOrDie() {
  const key = PropertiesService.getUserProperties().getProperty('USER_GEMINI_KEY');
  if (!key) throw new Error("⚠️ Setup Required: Please click the 'Settings' button and enter your Gemini API Key.");
  return key;
}

/**
 * SECURITY: Prevents CSV Injection.
 * If a string starts with =, +, -, or @, Google Sheets might execute it as a formula.
 * We prepend a single quote to force it to be treated as text.
 */
function sanitizeForSheets(value) {
  if (typeof value === 'string' && /^[=@+-]/.test(value)) {
    return "'" + value;
  }
  return value;
}


function addJobToSheet(formData) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found. Please create a tab named '${SHEET_NAME}'.`);

    // Sanitize inputs before using them
    const safeData = {
      jobTitle: sanitizeForSheets(formData.jobTitle || 'Unknown Role'),
      company: sanitizeForSheets(formData.company || 'Unknown Company'),
      status: sanitizeForSheets(formData.status || 'Applied'),
      source: sanitizeForSheets(formData.source || 'Manual Entry'),
      stage: sanitizeForSheets(formData.stage || 'Application'),
      description: sanitizeForSheets(formData.description || ''),
      email: sanitizeForSheets(formData.email || ''),
      nextSteps: sanitizeForSheets(formData.nextSteps || '')
    };

    const row = [
      new Date(),                 
      safeData.jobTitle,          
      safeData.description, 
      safeData.company,           
      safeData.status,            
      safeData.source,            
      safeData.email,       
      'Pending',                  
      safeData.stage,             
      '',                         
      '',                         
      safeData.nextSteps,   
      ''                          
    ];

    sheet.appendRow(row);
    
    if (formData.nextStepDate) {
      // Calendar events don't need CSV sanitization, but good to be clean
      createCalendarEvent(formData.company, formData.jobTitle, formData.nextStepDate, formData.reminderMinutes);
    }

    return { success: true, message: 'Job added successfully!' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}


function scanEmailsForJobs() {
  try {
    const apiKey = getApiKeyOrDie(); 

    const query = 'subject:("application" OR "interview" OR "assessment" OR "coding" OR "hirevue" OR "hackerrank" OR "candidate") newer_than:14d';
    const threads = GmailApp.search(query);
    
    if (threads.length === 0) {
      return { success: false, message: "Debug: Found 0 emails matching keywords in the last 14 days." };
    }

    let newJobsCount = 0;
    let updatedJobsCount = 0;
    let logBuffer = []; 

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: `Error: Sheet '${SHEET_NAME}' not found.` };

    const existingData = sheet.getDataRange().getValues();
    const companyRowMap = {};
    existingData.forEach((row, index) => {
      const companyName = row[3] ? row[3].toString().toLowerCase().trim() : "";
      if (companyName) companyRowMap[companyName] = index + 1; 
    });

    const threadsToProcess = threads.slice(0, 3); 

    for (const thread of threadsToProcess) {
      const messages = thread.getMessages();
      const subject = thread.getFirstMessageSubject();
      const emailDateObj = messages[messages.length - 1].getDate();
      const emailDateStr = emailDateObj.toDateString(); 

      const fullText = messages.map(m => 
        `From: ${m.getFrom()}\nDate: ${m.getDate()}\nSubject: ${m.getSubject()}\nBody: ${m.getPlainBody().substring(0, 2500)}`
      ).join("\n---\n");

      // CALL GEMINI
      const extraction = callGeminiAgent(fullText, apiKey, emailDateStr);
      
      if (typeof extraction === 'string') {
        if (!extraction.includes("No content")) {
           const cleanError = extraction.replace(/["{}]/g, "").substring(0, 150); 
           logBuffer.push(`❌ Failed '${subject}': ${cleanError}...`);
        }
        Utilities.sleep(10000); 
        continue;
      }

      if (extraction && extraction.applications && extraction.applications.length > 0) {
        extraction.applications.forEach(app => {
          if (!app.company || app.company === "N/A" || app.company.toLowerCase().includes("newsletter")) return;
          if (app.jobTitle && app.jobTitle.toLowerCase().includes("weekly problem")) return;

          const companyKey = app.company.toLowerCase().trim();
          
          // Sanitize Gemini Output
          const safeCompany = sanitizeForSheets(app.company);
          const safeTitle = sanitizeForSheets(app.jobTitle || "Unknown Title");
          const safeDesc = sanitizeForSheets(app.description || "Auto-extracted via Gemini");
          
          let nextSteps = "";
          let newStage = "Application";
          
          if (app.assessments && app.assessments.length > 0) {
            nextSteps = app.assessments.map(a => `${a.type} (Due: ${a.deadline})`).join("; ");
            nextSteps = sanitizeForSheets(nextSteps);
            newStage = "Assessment";
          }

          if (companyRowMap[companyKey]) {
            if (newStage === "Assessment") {
              const rowIndex = companyRowMap[companyKey];
              sheet.getRange(rowIndex, 9).setValue(newStage);   
              sheet.getRange(rowIndex, 12).setValue(nextSteps); 
              updatedJobsCount++;
              logBuffer.push(`🔄 Updated '${safeCompany}'`);
            } else {
              logBuffer.push(`ℹ️ Skipped '${safeCompany}': No new data.`);
            }
          } else {
            sheet.appendRow([
              new Date(), 
              safeTitle, 
              safeDesc, 
              safeCompany, 
              "Applied", 
              "Gmail Scan", 
              messages[0].getFrom(), 
              "Received", 
              newStage, 
              "", "", nextSteps, ""
            ]);
            companyRowMap[companyKey] = sheet.getLastRow();
            newJobsCount++;
            logBuffer.push(`✅ Added '${safeCompany}'`);
          }
        });
      }
      
      Utilities.sleep(10000);
    }
    
    const summary = `Scanned ${threadsToProcess.length} emails. (+${newJobsCount} New, 🔄${updatedJobsCount} Updated).`;
    const details = logBuffer.join("\n");
    return { success: true, message: summary + "\nDetails:\n" + details };

  } catch (e) {
    return { success: false, message: "Critical Script Error: " + e.toString() };
  }
}

function callGeminiAgent(emailText, apiKey, emailDateStr) {
const prompt = `
  Role: Job Application Tracker Assistant.
  Current Date Context: ${new Date().toDateString()}
  Email Reference Date: ${emailDateStr}
  Task: Analyze the email to identify legitimate job applications.

  Instructions:
   1. **Company Name Extraction**: 
      - Detect Platforms: If the email is from "HireVue", "HackerRank", "Workday", "Greenhouse", or "Ashby", read the email body to find the *Client* Company Name.
   2. **Job Title**: Extract the specific role.
   3. **Description (RAG)**: 
      - Trigger the Google Search tool: "Job description for [Job Title] at [Company]". 
      - Summarize in 1 sentence (Location + Start Date).
   4. **Assessments & Deadlines**: 
      - Look for "test", "coding challenge", "deadline", "complete by".
      - **CALCULATION**: If text says "7 days from receiving this email", ADD 7 days to Email Reference Date. Return 'YYYY-MM-DD HH:MM'. If time is missing, default to 23:59.
   5. **Strict Filtering**: 
      - IGNORE: Marketing, "Weekly Coding Challenges", Newsletters, Rejections.

  Output strictly valid JSON.
  
  Input Email Text:
  <email_content>
  ${emailText}
  </email_content>
  `;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  
  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "tools": [{ "google_search": {} }], 
    "generationConfig": { "temperature": 0.2 }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) return `API Error ${responseCode}: ${responseText}`;

    const json = JSON.parse(responseText);
    
    if (json.candidates && json.candidates[0].content) {
      const rawContent = json.candidates[0].content.parts[0].text;
      const startIndex = rawContent.indexOf('{');
      const endIndex = rawContent.lastIndexOf('}');
      if (startIndex !== -1 && endIndex !== -1) {
        return JSON.parse(rawContent.substring(startIndex, endIndex + 1));
      }
    }
    return "Error: No valid JSON returned.";
  } catch (e) {
    return "Exception: " + e.toString();
  }
}

function createCalendarEvent(company, title, dateString, reminderMinutes) {
  try {
    const startDate = new Date(dateString);
    if (isNaN(startDate.getTime())) return; 
    
    const eventTitle = `Deadline: ${company} - ${title}`;
    const calendar = CalendarApp.getDefaultCalendar();
    const endDate = new Date(startDate.getTime() + (60 * 60 * 1000)); 
    const event = calendar.createEvent(eventTitle, startDate, endDate);
    
    if (reminderMinutes) {
      event.removeAllReminders();
      event.addPopupReminder(parseInt(reminderMinutes));
    }
  } catch (e) {
    console.log("Calendar Error: " + e);
  }
}

function applyActiveFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet.getFilter()) sheet.getFilter().remove();
  const filter = sheet.getDataRange().createFilter();
  const criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['Rejected', 'Closed', 'Not Selected']).build();
  filter.setColumnFilterCriteria(5, criteria); 
  return { success: true, message: 'Filter applied.' };
}
