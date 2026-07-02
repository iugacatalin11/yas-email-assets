// YAS Beauty — Outlook signature add-in (event-based activation)
// Ruleaza la "New message compose" si insereaza semnatura cu datele utilizatorului.

var ASSETS = "https://iugacatalin11.github.io/yas-email-assets";
var USERS_URL = ASSETS + "/addin/users.json";

// Date comune (la fel pentru toti)
var SHARED = {
  phoneDisplay: "0750 641 690",
  phoneTel: "+40750641690",
  web: "yasbeauty.com",
  webUrl: "https://yasbeauty.com",
  instagram: "https://www.instagram.com/yasbeautylab/",
  tiktok: "https://www.tiktok.com/@yas.beauty.lab",
  facebook: "https://www.facebook.com/cursurimanicure"
};

function esc(s) {
  return (s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function buildSignature(name, title, phoneDisplay, phoneTel, email) {
  return '' +
'<table cellpadding="0" cellspacing="0" border="0" align="center" width="440" style="border-collapse:collapse;font-family:Arial,Helvetica,sans-serif;width:440px;">' +
  '<tr><td align="center" style="padding-bottom:12px;">' +
    '<img src="' + ASSETS + '/logo.png" alt="YAS Beauty" width="160" style="display:block;width:160px;height:auto;border:0;margin:0 auto;" /></td></tr>' +
  '<tr><td align="center" style="border-top:1px solid #e7e2dc;padding-top:13px;text-align:center;">' +
    '<div style="font-size:17px;font-weight:bold;color:#1a1a1a;line-height:1.2;letter-spacing:0.3px;">' + esc(name) + '</div>' +
    '<div style="font-size:11px;color:#b88a6a;font-weight:bold;letter-spacing:2px;text-transform:uppercase;padding-top:3px;">' + esc(title) + '</div>' +
    '<div style="font-size:13px;color:#555555;line-height:1.9;padding-top:11px;">' +
      '<a href="tel:' + phoneTel + '" style="color:#555555;text-decoration:none;"><img src="' + ASSETS + '/phone.png" width="14" height="14" style="display:inline-block;width:14px;height:14px;border:0;vertical-align:middle;" />&nbsp;&nbsp;' + esc(phoneDisplay) + '</a>&nbsp;&nbsp;&nbsp;&nbsp;' +
      '<a href="mailto:' + esc(email) + '" style="color:#555555;text-decoration:none;"><img src="' + ASSETS + '/mail.png" width="14" height="14" style="display:inline-block;width:14px;height:14px;border:0;vertical-align:middle;" />&nbsp;&nbsp;' + esc(email) + '</a><br/>' +
      '<a href="' + SHARED.webUrl + '" style="color:#555555;text-decoration:none;"><img src="' + ASSETS + '/web.png" width="14" height="14" style="display:inline-block;width:14px;height:14px;border:0;vertical-align:middle;" />&nbsp;&nbsp;' + SHARED.web + '</a>' +
    '</div>' +
    '<div style="padding-top:12px;line-height:0;">' +
      '<a href="' + SHARED.instagram + '" style="text-decoration:none;"><img src="' + ASSETS + '/instagram.png" width="20" height="20" style="display:inline-block;width:20px;height:20px;border:0;" /></a>' +
      '<a href="' + SHARED.tiktok + '" style="text-decoration:none;padding-left:14px;"><img src="' + ASSETS + '/tiktok.png" width="20" height="20" style="display:inline-block;width:20px;height:20px;border:0;" /></a>' +
      '<a href="' + SHARED.facebook + '" style="text-decoration:none;padding-left:14px;"><img src="' + ASSETS + '/facebook.png" width="20" height="20" style="display:inline-block;width:20px;height:20px;border:0;" /></a>' +
    '</div>' +
    '<div style="border-top:1px solid #eeeeee;margin-top:16px;padding-top:11px;font-size:10px;color:#aaaaaa;line-height:1.5;letter-spacing:0.2px;">' +
      'Acest mesaj poate con&#539;ine informa&#539;ii confiden&#539;iale destinate exclusiv destinatarului.' +
    '</div>' +
  '</td></tr></table>';
}

function insertSignature(event) {
  var email = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
  var displayName = Office.context.mailbox.userProfile.displayName || "";

  var apply = function (user) {
    var name = (user && user.name) || displayName;
    var title = (user && user.title) || "";
    var pd = (user && user.phoneDisplay) || SHARED.phoneDisplay;
    var pt = (user && user.phoneTel) || SHARED.phoneTel;
    var html = buildSignature(name, title, pd, pt, email);
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function () { event.completed(); }
    );
  };

  // Ia datele utilizatorului din users.json (nume + functie per email).
  fetch(USERS_URL, { cache: "no-store" })
    .then(function (r) { return r.json(); })
    .then(function (users) { apply(users[email]); })
    .catch(function () { apply(null); }); // fallback: doar nume din profil
}

Office.onReady(function () {});
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("insertSignature", insertSignature);
}
