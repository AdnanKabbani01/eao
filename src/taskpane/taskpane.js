document.getElementById("sendBtn").addEventListener("click", async function () {
  var userCommand = document.getElementById("messageInput").value;
  document.getElementById("messageInput").value = "";

  var chatBox = document.getElementById("chatBox");
  var newMessage = document.createElement("p");
  newMessage.textContent = "You: " + userCommand;
  chatBox.appendChild(newMessage);

  // Send the user command to the Python backend
  const response = await fetch("http://127.0.0.1:5000/run-command", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ command: userCommand }),
  });

  const result = await response.json();

  var botMessage = document.createElement("p");
  botMessage.textContent = "Bot: " + result.message;
  chatBox.appendChild(botMessage);
  chatBox.scrollTop = chatBox.scrollHeight;

  if (result.vba_code) {
    await Excel.run(async (context) => {
      // The VBA code is already applied in the backend, no further action needed here
    }).catch(function (error) {
      console.log(error);
      var errorMessage = document.createElement("p");
      errorMessage.textContent = "Error: " + error.message;
      chatBox.appendChild(errorMessage);
      chatBox.scrollTop = chatBox.scrollHeight;
    });
  }
});
