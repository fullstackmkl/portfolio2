let mailtoLinks = [];

function processFile() {
  const fileInput = document.getElementById('fileInput');
  const file = fileInput.files[0];
  const errorMessage = document.getElementById('error-message');
  const openAllEmailsButton = document.getElementById('openAllEmailsButton');
  const notificationType = document.getElementById('notificationType').value;

  if (!file) {
    errorMessage.textContent = 'Please upload a file first.';
    errorMessage.style.display = 'block';
    return;
  }

  if (!file.name.match(/\.(xlsx|xls)$/)) {
    errorMessage.textContent = 'Invalid file type. Please upload an Excel file.';
    errorMessage.style.display = 'block';
    return;
  }

  errorMessage.style.display = 'none';

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      if (jsonData.length === 0) {
        errorMessage.textContent = 'The file is empty.';
        errorMessage.style.display = 'block';
        return;
      }

      mailtoLinks = [];
      groupAndGenerateEmails(jsonData, notificationType);

      if (mailtoLinks.length > 0) {
        openAllEmailsButton.style.display = 'block';
      } else {
        openAllEmailsButton.style.display = 'none';
      }
    } catch (error) {
      errorMessage.textContent = 'Error processing file. Please ensure the file is in the correct format.';
      errorMessage.style.display = 'block';
    }
  };
  reader.onerror = function () {
    errorMessage.textContent = 'Error reading file. Please try again.';
    errorMessage.style.display = 'block';
  };
  reader.readAsArrayBuffer(file);
}

function groupAndGenerateEmails(data, notificationType) {
  // Sort data by Email
  data.sort((a, b) => a.Email.localeCompare(b.Email));

  const emailDraftsContainer = document.getElementById('emailDrafts');
  emailDraftsContainer.innerHTML = '';
  const errorMessage = document.getElementById('error-message');
  errorMessage.style.display = 'none';

  let currentEmail = '';
  let currentItems = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const email = row.Email;
    const manager = row.Owner || 'Manager';
    const timesheetId = row['Timesheet ID'] || 'N/A';

    if (!email) {
      continue; // Skip rows without an email
    }

    if (email !== currentEmail) {
      if (currentItems.length > 0) {
        // Generate email draft for the previous manager
        const emailBody = generateEmailBody(currentItems, notificationType);
        const mailtoLink = `mailto:${currentEmail}?subject=${getSubject(notificationType)}&body=${encodeURIComponent(emailBody)}`;
        mailtoLinks.push(mailtoLink);
        appendEmailDraft(emailDraftsContainer, currentEmail, emailBody, mailtoLink);
      }

      // Reset for the new manager
      currentEmail = email;
      currentItems = [];
    }

    currentItems.push({ manager, timesheetId });
  }

  // Generate email draft for the last manager in the list
  if (currentItems.length > 0) {
    const emailBody = generateEmailBody(currentItems, notificationType);
    const mailtoLink = `mailto:${currentEmail}?subject=${getSubject(notificationType)}&body=${encodeURIComponent(emailBody)}`;
    mailtoLinks.push(mailtoLink);
    appendEmailDraft(emailDraftsContainer, currentEmail, emailBody, mailtoLink);
  }

  if (emailDraftsContainer.innerHTML === '') {
    errorMessage.textContent = 'No valid data found in the file.';
    errorMessage.style.display = 'block';
  }
}

function generateEmailBody(items, notificationType) {
  let body = `Dear ${items[0].manager},\n\n`;

  switch (notificationType) {
    case 'timesheetReminder':
      body += 'This is a reminder to submit your timesheets. Please review and approve the following timesheets:\n\n';
      break;
    case 'draftTimesheetSubmittal':
      body += 'Please find the draft timesheet submissions below. Kindly review and submit your timesheets:\n\n';
      break;
    case 'auditReminder':
      body += 'This is an audit reminder. Please review the following timesheets for compliance:\n\n';
      break;
    default:
      body += 'Please review and approve the following timesheets:\n\n';
  }

  items.forEach(item => {
    body += `- Timesheet ID: ${item.timesheetId}\n`;
  });

  body += `\nThank you,\nYour Automated System`;
  return body;
}

function getSubject(notificationType) {
  switch (notificationType) {
    case 'timesheetReminder':
      return 'Timesheet Reminder';
    case 'draftTimesheetSubmittal':
      return 'Draft Timesheet Submittal';
    case 'auditReminder':
      return 'Audit Reminder';
    default:
      return 'Timesheet Approval Needed';
  }
}

function appendEmailDraft(container, email, body, mailtoLink) {
  const emailDraft = document.createElement('div');
  emailDraft.classList.add('email-draft');
  emailDraft.innerHTML = `
    <p><strong>To:</strong> ${email}</p>
    <p><strong>Subject:</strong> ${getSubject(document.getElementById('notificationType').value)}</p>
    <p><strong>Body:</strong></p>
    <pre>${body}</pre>
    <a href="${mailtoLink}" target="_blank">Open in Outlook</a>
  `;
  container.appendChild(emailDraft);
}

function openAllEmailsSequentially() {
  let i = 0;

  function openNextEmail() {
    if (i < mailtoLinks.length) {
      window.open(mailtoLinks[i], '_blank');
      i++;
      setTimeout(openNextEmail, 1000); // Adjust the delay as necessary
    }
  }

  openNextEmail();
}