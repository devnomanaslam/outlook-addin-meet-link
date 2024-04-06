import React from "react";

export default function App() {
  const handleInsertMeetingLink = () => {
    const test_url = "https://google.com/";
    const test_body = "Here is the example text.";
    Office.context.mailbox.item.location.setAsync(test_url, function (asyncResultLocation) {
      if (asyncResultLocation.status === Office.AsyncResultStatus.Failed) {
        console.error("Error setting location:", asyncResultLocation.error.message);
      } else {
        console.log("Location set successfully.");
      }
    });

    Office.context.mailbox.item.body.setAsync(
      test_body,
      { coercionType: Office.CoercionType.Html },
      function (setResult) {
        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Inserted...");
        } else {
          console.error("Error setting body: " + setResult.error.message);
        }
      }
    );
  };

  const buttonStyle = {
    backgroundColor: "#4CAF50" /* Green */,
    border: "none",
    color: "white",
    padding: "15px 0",
    textAlign: "center",
    textDecoration: "none",
    display: "block",
    width: "80%",
    margin: "0 auto",
    cursor: "pointer",
    borderRadius: "10px",
    boxShadow: "0 4px 6px rgba(0, 0, 0, 0.1)",
    position: "absolute",
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%)",
  };

  return (
    <div style={{ position: "relative", height: "100vh" }}>
      <button style={buttonStyle} onClick={handleInsertMeetingLink}>
        Insert meeting link
      </button>
    </div>
  );
}
