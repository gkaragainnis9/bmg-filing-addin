/*
 * BM+G Filing Auto-CC
 * iBNS - adds filing@bmplusg.com.au to CC on every new compose,
 * reply, reply-all and forward. Skips if already present.
 */

var FILING_ADDRESS = "filing@bmplusg.com.au";
var FILING_DISPLAY = "BM+G Filing";

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;

  item.cc.getAsync(function (result) {
    var alreadyPresent = false;

    if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
      for (var i = 0; i < result.value.length; i++) {
        var addr = (result.value[i].emailAddress || "").toLowerCase();
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
        event.completed();
      }
    );
  });
}

// Required for classic Outlook on Windows (JavaScript-only runtime)
// and for OWA / new Outlook (browser runtime via commands.html).
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
