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

function setSignatureHTML(html) {
  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      // Optionally handle errors here
    }
  );
}

// Event-based activation handlers
function onMessageCompose(event) {
  renderSignature("new").then(setSignatureHTML).finally(() => event.completed());
}
function onMessageReply(event) {
  renderSignature("reply").then(setSignatureHTML).finally(() => event.completed());
}
function onMessageForward(event) {
  renderSignature("forward").then(setSignatureHTML).finally(() => event.completed());
}
function onAppointmentCompose(event) {
  renderSignature("new").then(setSignatureHTML).finally(() => event.completed());
}

// For taskpane/manual use:
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