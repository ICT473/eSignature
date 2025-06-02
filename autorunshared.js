// Helper to fetch user profile info
async function getUserData() {
  const profile = Office.context.mailbox.userProfile;
  return {
    name: profile.displayName,
    email: profile.emailAddress,
    jobtitle: profile.title || "",
    department: profile.department || ""
  };
}

async function getTemplate(type) {
  const resp = await fetch(`https://ict473.github.io/eSignature/${type}.html`);
  return await resp.text();
}

async function renderSignature(type) {
  let template = await getTemplate(type);
  const user = await getUserData();
  template = template.replace(/%%name%%/g, user.name);
  template = template.replace(/%%email%%/g, user.email);
  template = template.replace(/%%jobtitle%%/g, user.jobtitle);
  template = template.replace(/%%department%%/g, user.department);
  return template;
}

function setSignatureHTML(html, callback) {
  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    callback
  );
}

// Detect the compose context for event-based activation
function getSignatureType() {
  // Outlook event-based activation does not pass event type, so infer from item
  if (Office.context.mailbox.item) {
    const subject = Office.context.mailbox.item.subject || "";
    // Try to detect reply/forward by subject prefix (not always reliable)
    if (/^re:/i.test(subject)) return "reply";
    if (/^fw:/i.test(subject)) return "forward";
    if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Appointment) return "new";
    return "new";
  }
  return "new";
}

// This function is called by all event-based activations
function checkSignature(event) {
  const type = getSignatureType();
  renderSignature(type).then(html => {
    setSignatureHTML(html, function() {
      if (event && typeof event.completed === "function") event.completed();
    });
  }).catch(() => {
    if (event && typeof event.completed === "function") event.completed();
  });
}

// For manual taskpane insertion
window.insertNewSignature = async function () {
  const html = await renderSignature("new");
  Office.context.mailbox.item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html });
};
window.insertReplySignature = async function () {
  const html = await renderSignature("reply");
  Office.context.mailbox.item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html });
};
window.insertForwardSignature = async function () {
  const html = await renderSignature("forward");
  Office.context.mailbox.item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html });
};

// For event-based activation (Office.actions)
if (typeof Office !== "undefined") {
  Office.actions = Office.actions || {};
  Office.actions.associate("checkSignature", checkSignature);
}