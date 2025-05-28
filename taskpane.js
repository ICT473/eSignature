const msalConfig = {
    auth: {
        clientId: 'd355dab3-5dcf-47d9-81ad-05012e7b885b',
        authority: "https://login.microsoftonline.com/247e416e-5287-401b-b997-da1aeca56d12",
        redirectUri: "https://gidcgd.sharepoint.com/sites/signatures/SiteAssets/taskpane.html"
    }
};

const SHAREPOINT_SITE_ID = "gidcgd.sharepoint.com,be9c3247-27e6-4f17-95fa-3ed8caec009c,11f0f19d-fb70-446a-a091-939199b6c93d";
const SHAREPOINT_LIST = "signaturesadmin";

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function getToken() {
    let account = msalInstance.getAllAccounts()[0];
    if (!account) {
        await msalInstance.loginPopup({ scopes: ["User.Read", "Sites.Read.All"] });
        account = msalInstance.getAllAccounts()[0];
    }
    const token = await msalInstance.acquireTokenSilent({ scopes: ["User.Read", "Sites.Read.All"] });
    return token.accessToken;
}

async function getUserProfile(token) {
    const resp = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` }
    });
    return await resp.json();
}

// Helper: determine if all recipients are internal
function isInternalMail(recipients) {
    const orgDomain = "gidcgd.sharepoint.com".replace(/^.*\/\/|\/.*$/g, "").replace(/^www\./, "");
    if (!recipients || recipients.length === 0) return false;
    return recipients.every(r => r.emailAddress && r.emailAddress.address.endsWith(`@${orgDomain}`));
}

async function getSignatureTemplate(token, scenario, recipientType) {
    const url = `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_ID}/lists/${SHAREPOINT_LIST}/items?$expand=fields`;
    const resp = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` }
    });
    const data = await resp.json();
    let match = data.value.find(
        i => i.fields.Scenario === scenario && (i.fields.RecipientType === recipientType || i.fields.RecipientType === "both")
    );
    if (!match && scenario !== "new") {
        match = data.value.find(i => i.fields.Scenario === "new" && (i.fields.RecipientType === recipientType || i.fields.RecipientType === "both"));
    }
    return match?.fields?.TemplateHtml || "<p>No signature template found.</p>";
}

function fillTemplate(html, profile) {
    return html
        .replace(/%%Firstname%%/g, profile.givenName || "")
        .replace(/%%Lastname%%/g, profile.surname || "")
        .replace(/%%Title%%/g, profile.jobTitle || "")
        .replace(/%%Department%%/g, profile.department || "")
        .replace(/%%Phone%%/g, profile.mobilePhone || (profile.businessPhones ? profile.businessPhones[0] : ""))
        .replace(/%%Fax%%/g, profile.fax || "")
        .replace(/%%Email%%/g, profile.mail || profile.userPrincipalName || "")
        .replace(/%%StreetAddress%%/g, profile.streetAddress || "")
        .replace(/%%City%%/g, profile.city || "");
}

async function showSignature() {
    const token = await getToken();
    const profile = await getUserProfile(token);

    await Office.onReady();
    let scenario = "new";
    let recipientType = "both";
    try {
        if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
            const item = Office.context.mailbox.item;
            if (item.itemClass && item.itemClass.indexOf("IPM.Note") === 0) {
                if (item.internetMessageId) {
                    scenario = "reply";
                } else {
                    scenario = "new";
                }
            }
            const allRecipients = [].concat(
                item.to || [],
                item.cc || [],
                item.bcc || []
            );
            if (isInternalMail(allRecipients)) {
                recipientType = "internal";
                scenario = "internal";
            }
        }
    } catch (e) {
        // fallback to new/both
    }

    const templateHtml = await getSignatureTemplate(token, scenario, recipientType);
    const filled = fillTemplate(templateHtml, profile);
    document.getElementById("signature-preview").innerHTML = filled;
    document.getElementById("insert-btn").onclick = function() {
        Office.context.mailbox.item.body.setSelectedDataAsync(filled, { coercionType: Office.CoercionType.Html });
    }
}

Office.onReady(() => {
    showSignature();
});