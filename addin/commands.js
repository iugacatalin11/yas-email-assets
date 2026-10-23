// YAS Beauty — Outlook signature add-in (event-based activation)  v1.1
// Ruleaza la:
//   - OnNewMessageCompose   -> insereaza semnatura cu datele utilizatorului
//   - OnMessageFromChanged  -> re-insereaza semnatura cand se schimba expeditorul
//                              (ex. From = Office@ / Contact@). Fara asta, Outlook
//                              (mai ales pe Mac) inlocuieste semnatura cu cea
//                              implicita a contului partajat ("Get Outlook for Mac").

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

// users.json cu timeout: daca nu raspunde in 4s, mergem pe fallback (profilul Office)
function fetchUsers(timeoutMs) {
  return new Promise(function (resolve) {
    var done = false;
    var finish = function (val) { if (!done) { done = true; resolve(val); } };
    setTimeout(function () { finish(null); }, timeoutMs);
    try {
      fetch(USERS_URL, { cache: "no-store" })
        .then(function (r) { return r.json(); })
        .then(function (users) { finish(users); })
        .catch(function () { finish(null); });
    } catch (e) {
      finish(null);
    }
  });
}

// Adresa curenta din campul From (compose). Daca API-ul nu e disponibil
// sau nu raspunde in 3s, intoarcem null si folosim adresa din profil.
function getFromAddress(timeoutMs) {
  return new Promise(function (resolve) {
    var done = false;
    var finish = function (val) { if (!done) { done = true; resolve(val); } };
    setTimeout(function () { finish(null); }, timeoutMs);
    try {
      var item = Office.context.mailbox.item;
      if (item && item.from && item.from.getAsync) {
        item.from.getAsync(function (res) {
          if (res && res.status === Office.AsyncResultStatus.Succeeded && res.value) {
            finish({
              email: (res.value.emailAddress || "").toLowerCase(),
              displayName: res.value.displayName || ""
            });
          } else {
            finish(null);
          }
        });
      } else {
        finish(null);
      }
    } catch (e) {
      finish(null);
    }
  });
}

// Logica comuna: construieste si seteaza semnatura.
//   - numele + functia vin din users.json dupa adresa PERSOANEI (profilul Office),
//     ca sa apara cine scrie efectiv, chiar daca trimite ca Office@/Contact@;
//   - adresa afisata in semnatura e cea din From (cutia partajata, daca e cazul).
//   Exceptie: daca in users.json exista o intrare pentru adresa partajata
//   (ex. "contact@yasbeauty.com": { "name": "Echipa YAS Beauty", "title": "..." }),
//   atunci se foloseste acel nume in locul numelui persoanei.
function applySignature(event, useFromField) {
  var completed = false;
  var finish = function () {
    if (completed) { return; }
    completed = true;
    try { event.completed(); } catch (e) { /* nimic */ }
  };
  try {
    var profile = (Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) || {};
    var selfEmail = (profile.emailAddress || "").toLowerCase();
    var displayName = profile.displayName || "";

    var fromPromise = useFromField ? getFromAddress(3000) : Promise.resolve(null);

    Promise.all([fetchUsers(4000), fromPromise]).then(function (results) {
      try {
        var users = results[0];
        var from = results[1];
        var fromEmail = (from && from.email) || selfEmail;

        var self = users && users[selfEmail];
        var shared = (fromEmail !== selfEmail) && users && users[fromEmail];

        var name = (shared && shared.name) || (self && self.name) || displayName;
        var title = (shared && shared.name) ? (shared.title || "") : ((self && self.title) || "");
        var pd = (shared && shared.phoneDisplay) || (self && self.phoneDisplay) || SHARED.phoneDisplay;
        var pt = (shared && shared.phoneTel) || (self && self.phoneTel) || SHARED.phoneTel;

        var html = buildSignature(name, title, pd, pt, fromEmail);
        Office.context.mailbox.item.body.setSignatureAsync(
          html,
          { coercionType: Office.CoercionType.Html },
          function () { finish(); }
        );
        // plasa de siguranta: daca setSignatureAsync nu mai revine, inchidem evenimentul
        setTimeout(finish, 8000);
      } catch (e) {
        finish();
      }
    });
  } catch (e) {
    finish();
  }
}

// OnNewMessageCompose — mesaj nou (From = contul principal, de regula)
function insertSignature(event) {
  applySignature(event, true);
}

// OnMessageFromChanged — utilizatorul a schimbat expeditorul (ex. pe Office@/Contact@).
// Outlook tocmai a inlocuit semnatura cu cea a contului nou; o punem pe a noastra la loc.
function onFromChanged(event) {
  applySignature(event, true);
}

// Inregistrarea handler-elor si handshake-ul obligatoriu cu office.js.
// Office.initialize / Office.onReady TREBUIE apelate sincron la incarcarea
// scriptului, altfel office.js arunca "Office.js has not fully loaded" si
// handler-ul nu mai e apelat niciodata.
function registerHandler() {
  try {
    if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
      Office.actions.associate("insertSignature", insertSignature);
      Office.actions.associate("onFromChanged", onFromChanged);
      return true;
    }
  } catch (e) { /* nimic */ }
  return false;
}

if (typeof Office !== "undefined") {
  Office.initialize = function () {};
  registerHandler();
  if (Office.onReady) {
    Office.onReady(function () { registerHandler(); });
  }
}
