/* Register the send handler named by manifest.xml. */
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

function completeEventWithDiagnostics(event, options, outcome) {
  console.log("[OSA-DIAG] immediately before event.completed", outcome, options);
  event.completed(options);
  console.log("[OSA-DIAG] event.completed returned; host acceptance is not confirmed");
}

function onMessageSendHandler(event) {
  console.log("[OSA-DIAG] onMessageSendHandler entered");

  // In compose mode, read the current subject asynchronously.
  try {
    const messageItem = Office.context.mailbox.item;
    console.log("[OSA-DIAG] immediately before subject.getAsync");
    messageItem.subject.getAsync((result) => {
      console.log("[OSA-DIAG] subject.getAsync callback entered");
      console.log("[OSA-DIAG] AsyncResult status:", result.status);
      console.log("[OSA-DIAG] returned Subject value:", result.value);

      if (result.status !== Office.AsyncResultStatus.Succeeded || typeof result.value !== "string") {
        console.error("Unable to retrieve the message subject.", result.error);
        completeEventWithDiagnostics(event, { allowEvent: true }, "subject retrieval failed; allow send");
        return;
      }

      // Test the subject as-is with the specified, case-sensitive expression.
      const subjectPattern = /^(?:(RE|FW|FWD):\s*)?\[[^\[\]]+\]/;
      if (subjectPattern.test(result.value)) {
        // A valid subject allows the message to continue sending.
        completeEventWithDiagnostics(event, { allowEvent: true }, "valid subject; allow send");
        return;
      }

      // A failed validation invokes Smart Alerts; PromptUser permits Send Anyway or Don't Send.
      completeEventWithDiagnostics(event, {
        allowEvent: false,
        errorMessage: "Subject format reminder\nPlease use: [Matter Code] Subject\nExample: [LO26A001] Legal Opinion"
      }, "invalid subject; invoke PromptUser Smart Alert");
    });
  } catch (error) {
    console.error("Unable to start subject retrieval.", error);
    completeEventWithDiagnostics(event, { allowEvent: true }, "subject retrieval threw; allow send");
  }
}