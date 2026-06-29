// ─────────────────────────────────────────────────────────────────
// SPRINGBOARD TALENT — ASSESSMENT RESPONSE HANDLER
// Google Apps Script Web App
// Deploy as: Execute as Me | Anyone can access
// ─────────────────────────────────────────────────────────────────

// CONFIGURATION — update this folder ID
const RESPONSES_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';
// Get this from the URL of your "01 — Raw Responses" folder in Drive
// e.g. https://drive.google.com/drive/folders/1ABC123... → use 1ABC123...

function doPost(e) {
  try {
    const raw = e.postData.contents;
    const data = JSON.parse(raw);
    saveResponseToDoc(data);
    return ContentService
      .createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error',message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput('Springboard Talent Assessment Handler — OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function saveResponseToDoc(data) {
  const candidate = data.candidate || {};
  const name = candidate.name || 'Unknown';
  const domain = candidate.domain || '';
  const level = candidate.level || '';
  const email = candidate.email || '';
  const location = candidate.location || '';
  const timestamp = data.timestamp || new Date().toISOString();
  const dateStr = timestamp.substring(0,10);
  const timeSpent = Math.round((data.timeSpent || 0) / 60);

  // Create the doc
  const docTitle = `${name} — ${domain} × ${level} — ${dateStr}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // ── HEADER ──
  const header = body.appendParagraph('SPRINGBOARD TALENT — CANDIDATE ASSESSMENT RESPONSES');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAttributes({[DocumentApp.Attribute.FOREGROUND_COLOR]: '#0D2B1F'});

  body.appendParagraph(`Name: ${name}`).setHeading(DocumentApp.ParagraphHeading.NORMAL);
  body.appendParagraph(`Email: ${email}`);
  body.appendParagraph(`Domain: ${domain}  |  Level: ${level}`);
  body.appendParagraph(`Location: ${location}`);
  body.appendParagraph(`Submitted: ${timestamp}`);
  body.appendParagraph(`Time spent: ${timeSpent} minutes`);
  body.appendHorizontalRule();

  // ── SECTION A + DOMAIN QUESTIONS ──
  body.appendParagraph('SECTION A — MOMENT RETRIEVAL').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const answers = data.answers || {};
  Object.keys(answers).forEach(key => {
    const ans = answers[key];
    if(!ans || key === 'FC') return;
    if(ans.code && ans.code.startsWith('A')) {
      body.appendParagraph(ans.eyebrow || ans.code)
        .setAttributes({[DocumentApp.Attribute.BOLD]: true});
      body.appendParagraph(`Question: ${ans.question}`);
      body.appendParagraph(`Answer: ${ans.answer || '(no response)'}`);
      if(ans.followup) {
        body.appendParagraph(`Follow-up: ${ans.followup}`)
          .setAttributes({[DocumentApp.Attribute.ITALIC]: true});
        body.appendParagraph(`Follow-up answer: ${ans.followupAnswer || '(no response)'}`);
      }
      body.appendParagraph('');
    }
  });

  // ── SECTION B ──
  body.appendHorizontalRule();
  body.appendParagraph('SECTION B — FORCED PRIORITISATION').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  Object.keys(answers).forEach(key => {
    const ans = answers[key];
    if(!ans || key === 'FC') return;
    if(ans.code && ans.code.startsWith('B')) {
      body.appendParagraph(ans.eyebrow || ans.code)
        .setAttributes({[DocumentApp.Attribute.BOLD]: true});
      body.appendParagraph(`Question: ${ans.question}`);
      if(ans.ranking) {
        const rankLines = Object.entries(ans.ranking)
          .map(([opt, rank]) => `  Option ${opt}: Ranked ${rank || 'not set'}`)
          .join('\n');
        body.appendParagraph(`Ranking:\n${rankLines}`);
      }
      body.appendParagraph(`Rationale: ${ans.rationale || '(no response)'}`);
      body.appendParagraph('');
    }
  });

  // ── SECTION C — FORCED CHOICE ──
  body.appendHorizontalRule();
  body.appendParagraph('SECTION C — FORCED CHOICE PAIRS').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const fc = data.fcAnswers || answers['FC'] || {};
  const fcLabels = {
    FC1:{q:'When solving a problem I am more naturally drawn to…',A:'Finding the best solution for this situation',B:'Finding a solution that works across similar situations'},
    FC2:{q:'I get more energy from…',A:'Closing something',B:'Opening something'},
    FC3:{q:'When a process isn\'t working, I first want to…',A:'Fix the immediate bottleneck',B:'Understand why the process exists'},
    FC4:{q:'In a planning discussion I am more likely to…',A:'Push to make a decision',B:'Raise the question no one has asked'},
    FC5:{q:'I do my best work when…',A:'Given a clear mandate',B:'Given a problem and freedom to define the solution'},
    FC6:{q:'Work I look back on most proudly is where I…',A:'Built something that didn\'t exist',B:'Made something existing work better'},
    FC7:{q:'When joining a new organisation I…',A:'Understand culture before changing anything',B:'Start identifying what needs to change from day one'},
    FC8:{q:'When something is broken my instinct is to…',A:'Fix it and move on',B:'Understand why it broke'},
    FC9:{q:'Under deadline pressure I am more likely to…',A:'Cut scope to protect quality',B:'Push the team harder to protect scope'},
    FC10:{q:'When I receive feedback I disagree with…',A:'Explain my reasoning and defend my position',B:'Ask questions before responding'},
    FC11:{q:'When a project fails I first think about…',A:'What I could have done differently',B:'What circumstances made it difficult'},
    FC12:{q:'When a wrong decision is about to be made I…',A:'Say clearly I think it\'s wrong',B:'Raise concern once then support the group'}
  };

  Object.keys(fcLabels).forEach(key => {
    const chosen = fc[key];
    const lbl = fcLabels[key];
    if(!lbl) return;
    const chosenText = chosen ? `${chosen}: ${lbl[chosen]}` : '(not answered)';
    body.appendParagraph(`${key} — ${lbl.q}`)
      .setAttributes({[DocumentApp.Attribute.BOLD]: true});
    body.appendParagraph(`Selected: ${chosenText}`);
  });

  // ── SECTION D — REFLECTIVE ──
  body.appendHorizontalRule();
  body.appendParagraph('SECTION D — REFLECTIVE DILEMMA').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  Object.keys(answers).forEach(key => {
    const ans = answers[key];
    if(!ans || key === 'FC') return;
    if(ans.code && ans.code.startsWith('D')) {
      body.appendParagraph(ans.eyebrow || ans.code)
        .setAttributes({[DocumentApp.Attribute.BOLD]: true});
      body.appendParagraph(`Question: ${ans.question}`);
      body.appendParagraph(`Answer: ${ans.answer || '(no response)'}`);
      body.appendParagraph('');
    }
  });

  // ── ANALYSIS TRIGGER ──
  body.appendHorizontalRule();
  body.appendParagraph('FOR CLAUDE ANALYSIS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('To generate the Signal Brief, copy all content above and paste into the Springboard Talent — Signal Analysis project in Claude with the message:');
  body.appendParagraph('"Analyse this assessment and generate the Signal Brief"')
    .setAttributes({[DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.ITALIC]: true});

  // Save doc
  doc.saveAndClose();

  // Move to the responses folder
  try {
    const folder = DriveApp.getFolderById(RESPONSES_FOLDER_ID);
    const file = DriveApp.getFileById(doc.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch(err) {
    // If folder move fails, doc stays in root — still saved
    Logger.log('Folder move failed: ' + err.toString());
  }

  // Send email notification
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if(userEmail) {
      const docUrl = 'https://docs.google.com/document/d/' + doc.getId();
      MailApp.sendEmail({
        to: userEmail,
        subject: `New assessment: ${name} — ${domain} × ${level}`,
        body: `A new Springboard Talent assessment has been submitted.\n\nCandidate: ${name}\nDomain: ${domain}\nLevel: ${level}\nEmail: ${email}\nTime spent: ${timeSpent} minutes\n\nOpen the response doc:\n${docUrl}\n\nTo generate the Signal Brief, open the Springboard Talent — Signal Analysis project in Claude, paste the doc contents, and type:\n"Analyse this assessment and generate the Signal Brief"`
      });
    }
  } catch(err) {
    Logger.log('Email failed: ' + err.toString());
  }
}
