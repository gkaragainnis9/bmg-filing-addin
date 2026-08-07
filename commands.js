/*
 * BM+G Filing Auto-CC v3
 * iBNS - adds filing@bmplusg.com.au to CC on every new compose,
 * reply, reply-all and forward. Skips if already present.
 * ES2016-and-earlier syntax only (event add-in runtime constraint).
 * v3: unconditional Office.actions.associate (matches Microsoft's
 * autolaunch walkthrough sample) + console breadcrumbs.
 */

var FILING_ADDRESS = "filing@bmplusg.com.au";
var FILING_DISPLAY = "BM+G Filing";

console.log("[FilingCC] commands.js loaded");

function onNewMessageComposeHandler(event) {
  console.log("[FilingCC] handler invoked");
  var item = Office.context.mailbox.item;

  item.cc.getAsync(function (result) {
    console.log("[FilingCC] cc.getAsync status: " + result.status);
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
      console.log("[FilingCC] filing address already in CC, skipping");
      event.completed();
      return;
    }

    item.cc.addAsync(
      [{ displayName: FILING_DISPLAY, emailAddress: FILING_ADDRESS }],
      function (addResult) {
        console.log("[FilingCC] cc.addAsync status: " + addResult.status);
        event.completed();
      }
    );
  });
}

try {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  console.log("[FilingCC] handler associated");
} catch (e) {
  console.log("[FilingCC] associate failed: " + e.message);
}
