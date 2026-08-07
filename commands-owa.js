/*
 * BM+G Filing Auto-CC - OWA diagnostic build
 * Registers BOTH OnNewMessageCompose and OnMessageRecipientsChanged.
 * Idempotent: whichever fires first adds filing@; later firings no-op.
 * ES2016-and-earlier syntax only.
 */

var FILING_ADDRESS = "filing@bmplusg.com.au";
var FILING_DISPLAY = "BM+G Filing";

console.log("[FilingOWA] commands-owa.js loaded");

function addFilingCc(event, source) {
  console.log("[FilingOWA] handler fired via: " + source);
  var item = Office.context.mailbox.item;

  item.cc.getAsync(function (result) {
    console.log("[FilingOWA] cc.getAsync status: " + result.status);
    var alreadyPresent = false;
    var i;
    var addr;

    if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
      for (i = 0; i < result.value.length; i++) {
        addr = (result.value[i].emailAddress || "").toLowerCase();
        if (addr === FILING_ADDRESS) {
          alreadyPresent = true;
          break;
        }
      }
    }

    if (alreadyPresent) {
      console.log("[FilingOWA] already present, no-op");
      event.completed();
      return;
    }

    item.cc.addAsync(
      [{ displayName: FILING_DISPLAY, emailAddress: FILING_ADDRESS }],
      function (addResult) {
        console.log("[FilingOWA] cc.addAsync status: " + addResult.status + " (via " + source + ")");
        event.completed();
      }
    );
  });
}

function onNewMessageComposeHandler(event) {
  addFilingCc(event, "OnNewMessageCompose");
}

function onRecipientsChangedHandler(event) {
  addFilingCc(event, "OnMessageRecipientsChanged");
}

try {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
  console.log("[FilingOWA] both handlers associated");
} catch (e) {
  console.log("[FilingOWA] associate failed: " + e.message);
}
