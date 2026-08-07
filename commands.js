/*
 * BM+G Filing Auto-CC - production
 * iBNS - adds filing@bmplusg.com.au to CC on every new compose,
 * reply, reply-all and forward. Skips if already present.
 * ES2016-and-earlier syntax only (event add-in runtime constraint).
 */

var FILING_ADDRESS = "filing@bmplusg.com.au";
var FILING_DISPLAY = "BM+G Filing";

function notify(text) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync("filingcc", {
      type: "informationalMessage",
      message: text,
      icon: "Icon.16x16",
      persistent: false
    });
  } catch (e) {
    // Notification is best-effort only.
  }
}

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;

  item.cc.getAsync(function (result) {
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
      event.completed();
      return;
    }

    item.cc.addAsync(
      [{ displayName: FILING_DISPLAY, emailAddress: FILING_ADDRESS }],
      function () {
        notify("filing@bmplusg.com.au added to CC. Remove it if this email should not be filed.");
        event.completed();
      }
    );
  });
}

try {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
} catch (e) {
  // Older runtimes resolve the handler by global name.
}
