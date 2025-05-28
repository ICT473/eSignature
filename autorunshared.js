async function getTemplate(templateName) {
  const url = `https://gidcgd.sharepoint.com/sites/signatures/SiteAssets/${templateName}`;
  const response = await fetch(url, { credentials: "include" });
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

if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate("checkSignature", checkSignature);
}