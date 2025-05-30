// Fetches the template from your GitHub Pages hosting
async function getTemplate(templateName) {
  const url = `https://ict473.github.io/eSignature/${templateName}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error("Failed to fetch template: " + url);
  return await response.text();
}

function checkSignature(event) {
  let item = Office.context.mailbox.item;
  let templateFile = "new.html"; // default

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    switch (item.messageType) {
      case Office.MailboxEnums.MessageType.Reply:
      case Office.MailboxEnums.MessageType.ReplyAll:
        templateFile = "reply.html";
        break;
      case Office.MailboxEnums.MessageType.Forward:
        templateFile = "forward.html";
        break;
      default:
        templateFile = "new.html";
    }
  }

  getTemplate(templateFile)
    .then(templateHtml => {
      item.body.getAsync(Office.CoercionType.Html, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          let body = result.value || "";
          if (body.trim() === "") {
            item.body.setAsync(
              templateHtml,
              { coercionType: Office.CoercionType.Html },
              function() { event.completed(); }
            );
          } else {
            event.completed();
          }
        } else {
          event.completed();
        }
      });
    })
    .catch(() => { event.completed(); });
}

// Office.actions.associate for autorun events (modern requirement)
if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate("checkSignature", checkSignature);
}

// For environments where global window binding is required for autorun
window.checkSignature = checkSignature;