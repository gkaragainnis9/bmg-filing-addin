/*
 * BM+G Filing Auto-CC v2
 * iBNS - adds filing@bmplusg.com.au to CC on every new compose,
 * reply, reply-all and forward. Skips if already present.
 * Code kept to ES2016-and-earlier per Microsoft guidance for
 * event-based add-ins (no async/await, no ternary operators).
 */

var FILING_ADDRESS = "filing@bmplusg.com.au";
var FILING_DISPLAY = "BM+G Filing";

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
        event.completed();
      }
    );
  });
}

/*
 * Per Microsoft troubleshooting guidance for event-based add-ins:
 * classic Outlook on Windows requires Office.actions.associate to map
 * the manifest FunctionName to the handler. Gate by platform.
 * On the web/new Outlook the handler is resolved by its global name.
 */
if (Office.context && (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null)) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
