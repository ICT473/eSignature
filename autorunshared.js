async function getUserData() {
  const profile = Office.context.mailbox.userProfile;
  const displayName = profile.displayName || "";
  const [firstname, ...lastnameArr] = displayName.split(" ");
  const lastname = lastnameArr.join(" ");
  return {
    firstname,
    lastname,
    title: profile.title || "",
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
  template = template
    .replace(/%%Firstname%%/gi, user.firstname)
    .replace(/%%Lastname%%/gi, user.lastname)
    .replace(/%%Title%%/gi, user.title)
    .replace(/%%Department%%/gi, user.department)
    .replace(/%%Email%%/gi, user.email);
  // Fields not available in Office.js
  template = template
    .replace(/%%Phone%%/gi, "")
    .replace(/%%Fax%%/gi, "")
    .replace(/%%StreetAddress%%/gi, "")
    .replace(/%%City%%/gi, "");
  return template;
}

// ... rest of your code (setSignatureHTML, checkSignature, etc.)