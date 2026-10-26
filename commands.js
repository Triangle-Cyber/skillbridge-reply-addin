Office.onReady(() => {});

const WAITLIST_URL = "https://job-boards.greenhouse.io/trianglecyber/jobs/5384277008";

function buildBody() {
return `
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;">
<p>At the moment, all available SkillBridge slots are filled. We will reopen applications as soon as additional opportunities become available.</p>
<p>Join our waitlist to be notified when applications reopen.</p>
<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="margin:16px 0;">
<tr><td align="center" bgcolor="#0078D4" style="border-radius:4px;">
<a href="${WAITLIST_URL}" target="_blank" style="display:inline-block;padding:12px 24px;font-family:Calibri,Arial,sans-serif;font-size:11pt;font-weight:bold;color:#ffffff;text-decoration:none;border-radius:4px;">Add to Waitlist</a>
</td></tr>
</table>
</div>`;
}

function showError(message) {
const item = Office.context.mailbox.item;
if (item && item.notificationMessages) {
item.notificationMessages.replaceAsync("waitlistError", {
type: "errorMessage",
message: message
});
}
}

function createWaitlistReply(event) {
showError("DIAG: Waitlist command ran");
try {
Office.context.mailbox.item.displayReplyForm({ htmlBody: buildBody() });
} catch (err) {
console.error(err);
showError("Waitlist Reply failed: " + err.message);
}
event.completed();
}

Office.actions.associate("createWaitlistReply", createWaitlistReply);
