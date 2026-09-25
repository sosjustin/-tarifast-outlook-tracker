function onMessageSendHandler(event) {
  event.completed({
    allowEvent: true
  });
}

Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);
