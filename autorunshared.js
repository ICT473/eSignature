async function getUserData() {
  const profile = Office.context.mailbox.userProfile;
  const displayName = profile.displayName || "";
  const [firstname, ...lastnameArr] = displayName.split(" ");
  const lastname = lastnameArr.join(" ");
  return {
    name: displayName,
    firstname,
    lastname,
    title: profile.title || "",
    jobtitle: profile.title || "",
    department: profile.department || "",
    email: profile.emailAddress || ""
  };
}

async function getTemplate(type) {
  const resp = await fetch(`https://ict473.github.io/eSignature/${type}.html`);
  return await resp.text();
}

async function renderSignature(type) {
  let template = await getTemplate(type);
  const user = await getUserData();
  // Replace all supported placeholders
  template = template
    .replace(/%%name%%/gi, user.name)
    .replace(/%%Firstname%%/gi, user.firstname)
    .replace(/%%Lastname%%/gi, user.lastname)
    .replace(/%%Title%%/gi, user.title)
    .replace(/%%jobtitle%%/gi, user.jobtitle)
    .replace(/%%Department%%/gi, user.department)
    .replace(/%%Email%%/gi, user.email)
    // For placeholders not in Office.js, blank them out
    .replace(/%%Phone%%/gi, "")
    .replace(/%%Fax%%/gi, "")
    .replace(/%%StreetAddress%%/gi, "")
    .replace(/%%City%%/gi, "");
  return template;
}

function setSignatureHTML(html, callback) {
  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    callback
  );
}

function getSignatureType() {
  if (Office.context.mailbox.item) {
    const subject = Office.context.mailbox.item.subject || "";
    if (/^re:/i.test(subject)) return "reply";
    if (/^fw:/i.test(subject)) return "forward";
    if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Appointment) return "new";
    return "new";
  }
  return "new";
}

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

if (typeof Office !== "undefined") {
  Office.actions = Office.actions || {};
  Office.actions.associate("checkSignature", checkSignature);
}