Office.onReady(() => {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
});

async function onNewMessageComposeHandler(event) {
  try {
    let type = "new";
    const item = Office.context.mailbox.item;
    if (item.getComposeTypeAsync) {
      await new Promise(resolve => {
        item.getComposeTypeAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            switch (result.value) {
              case Office.MailboxEnums.ComposeType.NewMail: type = "new"; break;
              case Office.MailboxEnums.ComposeType.Reply: type = "reply"; break;
              case Office.MailboxEnums.ComposeType.Forward: type = "forward"; break;
            }
          }
          resolve();
        });
      });
    } else {
      const body = await getBodyText();
      if (body.includes("From:") || body.includes("Sent:")) {
        type = "forward";
      } else if (item.conversationId && item.internetMessageId) {
        type = "reply";
      }
    }

    const signatureUrl = `https://ict473.github.io/eSignature/${type}.html`;
    const response = await fetch(signatureUrl);
    const signatureHtml = await response.text();

    Office.context.mailbox.item.body.setAsync(
      signatureHtml,
      { coercionType: Office.CoercionType.Html },
      () => event.completed && event.completed()
    );
  } catch (err) {
    console.error("Signature insertion failed:", err);
    if (event && typeof event.completed === "function") event.completed();
  }
}

function getBodyText() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      resolve(result.value || "");
    });
  });
}
