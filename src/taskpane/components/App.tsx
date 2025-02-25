import * as React from "react";
import { Button, Input, Text, Card } from "@fluentui/react-components";
import { useStyles } from "./App.styles";

interface AppProps {}

interface ChatMessage {
  id: number;
  text: string;
  isUser: boolean;
}

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<ChatMessage[]>([
    { id: 1, text: "Hello! How can I help you today?", isUser: false },
  ]);
  const [inputText, setInputText] = React.useState("");
  const chatContainerRef = React.useRef(null);

  const handleSendMessage = () => {
    if (inputText.trim() === "") return;

    // Add user message
    const newUserMessage: ChatMessage = {
      id: messages.length + 1,
      text: inputText,
      isUser: true,
    };

    setMessages([...messages, newUserMessage]);
    setInputText("");

    // Check for the specific message and respond accordingly
    let botResponseText = "";
    const normalizedInput = inputText.trim().replace(/\.$/, ""); // remove the dot
    if (normalizedInput === "Sort the table by sales in descending order") {
      botResponseText = "ok";
    } else if (normalizedInput === "Create a scatter plot of sales and costs") {
      botResponseText = "ok";
    } else if (normalizedInput === "Insert a column of profits") {
      botResponseText = "ok";
    } else {
      botResponseText = "目前暂不支持，请重新输入";
    }

    // Add bot response
    const botResponse: ChatMessage = {
      id: messages.length + 2,
      text: botResponseText,
      isUser: false,
    };
    setMessages((prev) => [...prev, botResponse]);
  };

  // Auto-scroll to bottom when messages change
  React.useEffect(() => {
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages]);

  return (
    <div className={styles.root}>
      <Text weight="semibold" size={500} as="h1">
        Chat Interface
      </Text>

      <div className={styles.chatContainer} ref={chatContainerRef}>
        {messages.map((message) => (
          <div key={message.id} className={styles.messageRow}>
            <Card className={message.isUser ? styles.userMessage : styles.botMessage}>
              <Text>{message.text}</Text>
            </Card>
          </div>
        ))}
      </div>

      <div className={styles.inputContainer}>
        <Input
          className={styles.inputField}
          value={inputText}
          onChange={(_e, data) => setInputText(data.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              handleSendMessage();
            }
          }}
          placeholder="Type your message here..."
        />
        <Button appearance="primary" onClick={handleSendMessage}>
          Send
        </Button>
      </div>
    </div>
  );
};

export default App;
