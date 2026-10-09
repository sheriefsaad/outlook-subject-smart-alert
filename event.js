/* global Office, console */

const subjectPattern = new RegExp("^(?:(RE|FW|FWD):\\s*)?\\[[^\\[\\]]+\\]");
const invalidSubjectMessage =
  "Subject format reminder\nPlease use: [Matter Code] Subject\nExample: [LO26A001] Legal Opinion";

function onMessageSendHandler(event) {
  let completed = false;

  const complete = (result) => {
    if (completed) {
      return;
    }

    completed = true;
    event.completed(result);
  };

  try {
    Office.context.mailbox.item.subject.getAsync((asyncResult) => {
      try {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Unable to retrieve the message subject.", asyncResult.error);
          complete({ allowEvent: true });
          return;
        }

        if (subjectPattern.test(asyncResult.value)) {
          complete({ allowEvent: true });
          return;
        }

        complete({
          allowEvent: false,
          errorMessage: invalidSubjectMessage,
        });
      } catch (error) {
        console.error("Error while validating the message subject.", error);
        complete({ allowEvent: true });
      }
    });
  } catch (error) {
    console.error("Unable to start message subject retrieval.", error);
    complete({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
