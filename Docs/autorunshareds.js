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

// Improved: Use Office.js API to detect compose type accurately
function getSignatureTypeAsync() {
  return new Promise((resolve) => {
    if (Office.context.mailbox.item && Office.context.mailbox.item.getComposeTypeAsync) {
      Office.context.mailbox.item.getComposeTypeAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          switch(result.value) {
            case Office.MailboxEnums.ComposeType.NewMail:
              resolve("new");
              break;
            case Office.MailboxEnums.ComposeType.Reply:
              resolve("reply");
              break;
            case Office.MailboxEnums.ComposeType.Forward:
              resolve("forward");
              break;
            default:
              resolve("new");
          }
        } else {
          resolve("new");
        }
      });
    } else {
      resolve("new");
    }
  });
}

// Updated: Now uses the improved compose type detection
function checkSignature(event) {
  getSignatureTypeAsync().then(type => {
    return renderSignature(type);
  }).then(html => {
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
