import * as React from "react";
import { Button, Input, Text, Card, Image } from "@fluentui/react-components";
import { useStyles } from "./App.styles";
/* global Excel */
interface AppProps {}

interface ChatMessage {
  id: number;
  type: "text" | "table" | "image";
  content: string | string[][] | string; // text content, table data, or image path
  isUser: boolean;
}

async function sortTableBySales(context: Excel.RequestContext, table: Excel.Table): Promise<void> {
  const salesColumn = table.columns.getItemOrNullObject("sales");
  salesColumn.load("index");
  await context.sync();

  if (salesColumn.isNullObject) {
    throw new Error("No 'sales' column found in the table");
  }

  const sortFields = [
    {
      key: salesColumn.index,
      ascending: false,
    },
  ];
  table.sort.apply(sortFields);
  await context.sync();
}

// Assume the first table is the source table
async function getFirstTable(context: Excel.RequestContext): Promise<Excel.Table> {
  const tables = context.workbook.worksheets.getActiveWorksheet().tables;
  tables.load("items");
  await context.sync();

  if (tables.items.length === 0) {
    throw new Error("No tables found in the worksheet");
  }
  return tables.items[0];
}

async function getSortedTableData(): Promise<string[][]> {
  let tableData: string[][] = [];

  await Excel.run(async (context) => {
    const sourceTable = await getFirstTable(context);
    sourceTable.load(["name", "headerRowRange"]);
    await context.sync();

    // Create a temporary worksheet for our copy
    const tempSheetName = `TempSheet_${Date.now()}`;
    const tempSheet = context.workbook.worksheets.add(tempSheetName);

    // Copy the source table
    const sourceRange = sourceTable.getRange();
    sourceRange.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    // Create a new table with the same data
    const tempTable = tempSheet.tables.add(
      tempSheet.getRange("A1").getResizedRange(sourceRange.rowCount - 1, sourceRange.columnCount - 1),
      true
    );
    tempTable.getRange().values = sourceRange.values;
    await context.sync();

    // Sort the temporary table
    await sortTableBySales(context, tempTable);
    const tempRange = tempTable.getRange();
    tempRange.load("values");
    await context.sync();

    // Convert to string array for our message
    tableData = tempRange.values.map((row) => row.map((cell) => (cell !== null ? String(cell) : "")));

    // Delete the temporary worksheet
    tempSheet.delete();
    await context.sync();
  });

  return tableData;
}

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<ChatMessage[]>([
    { id: 1, type: "text", content: "Hello! How can I help you today?", isUser: false },
    {
      id: 2,
      type: "table",
      content: [
        ["Product", "Sales", "Cost"],
        ["Item A", "$1200", "$800"],
        ["Item B", "$950", "$600"],
      ],
      isUser: false,
    },
    { id: 3, type: "image", content: "assets/logo-filled.png", isUser: false },
  ]);
  const [inputText, setInputText] = React.useState("");
  const chatContainerRef = React.useRef(null);

  const handleSendMessage = async () => {
    if (inputText.trim() === "") return;

    // Add user message
    const newUserMessage: ChatMessage = {
      id: messages.length + 1,
      type: "text",
      content: inputText,
      isUser: true,
    };

    setMessages([...messages, newUserMessage]);
    setInputText("");

    // Check for the specific message and respond accordingly
    let botResponse: ChatMessage;
    const normalizedInput = inputText.trim().replace(/\.$/, ""); // remove the dot
    if (normalizedInput === "Sort the table by sales in descending order") {
      try {
        const tableData = await getSortedTableData();

        botResponse = {
          id: messages.length + 2,
          type: "table",
          content: tableData,
          isUser: false,
        };
      } catch (error) {
        botResponse = {
          id: messages.length + 2,
          type: "text",
          content: error.toString(),
          isUser: false,
        };
      }
    } else if (normalizedInput === "Create a scatter plot of sales and costs") {
      botResponse = {
        id: messages.length + 2,
        type: "image",
        content: "assets/chart-example.png", // Example image path
        isUser: false,
      };
    } else if (normalizedInput === "Insert a column of profits") {
      botResponse = {
        id: messages.length + 2,
        type: "table",
        content: [
          ["Product", "Sales", "Cost", "Profit"],
          ["Item A", "$1200", "$800", "$400"],
          ["Item B", "$950", "$600", "$350"],
        ],
        isUser: false,
      };
    } else {
      botResponse = {
        id: messages.length + 2,
        type: "text",
        content: "目前暂不支持，请重新输入",
        isUser: false,
      };
    }

    setMessages((prev) => [...prev, botResponse]);
  };

  const handleApplySortToWorksheet = async () => {
    try {
      await Excel.run(async (context) => {
        const table = await getFirstTable(context);
        await sortTableBySales(context, table);
      });

      const successMessage: ChatMessage = {
        id: messages.length + 1,
        type: "text",
        content: "Table sorted successfully!",
        isUser: false,
      };
      setMessages((prev) => [...prev, successMessage]);
    } catch (error) {
      const errorMessage: ChatMessage = {
        id: messages.length + 1,
        type: "text",
        content: `Error: ${error.toString()}`,
        isUser: false,
      };
      setMessages((prev) => [...prev, errorMessage]);
    }
  };

  const renderMessageContent = (message: ChatMessage) => {
    switch (message.type) {
      case "text":
        return <Text>{message.content as string}</Text>;

      case "image":
        return <Image src={message.content as string} alt="Chat image" className={styles.messageImage} />;

      case "table": {
        const tableData = message.content as string[][];
        return (
          <div>
            <table className={styles.messageTable}>
              <thead>
                <tr>
                  {tableData[0].map((header, index) => (
                    <th key={index}>{header}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {tableData.slice(1).map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {row.map((cell, cellIndex) => (
                      <td key={cellIndex}>{cell}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
            <div className={styles.tableActions}>
              <Button appearance="primary" size="small" onClick={handleApplySortToWorksheet}>
                Apply to Worksheet
              </Button>
            </div>
          </div>
        );
      }

      default:
        return <Text>Unsupported message type</Text>;
    }
  };

  // Auto-scroll to bottom when messages change
  React.useEffect(() => {
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages]);

  // Function to handle clicking on a recommended prompt button
  const handlePromptClick = (promptText: string) => {
    setInputText(promptText);
  };

  return (
    <div className={styles.root}>
      <Text weight="semibold" size={500} as="h1">
        Chat Interface
      </Text>

      <div className={styles.chatContainer} ref={chatContainerRef}>
        {messages.map((message) => (
          <div key={message.id} className={styles.messageRow}>
            <Card className={message.isUser ? styles.userMessage : styles.botMessage}>
              {renderMessageContent(message)}
            </Card>
          </div>
        ))}
      </div>

      <div className={styles.recommendedPrompts}>
        <Button
          appearance="outline"
          size="small"
          onClick={() => handlePromptClick("Sort the table by sales in descending order")}
        >
          Sort table by sales
        </Button>
        <Button
          appearance="outline"
          size="small"
          onClick={() => handlePromptClick("Create a scatter plot of sales and costs")}
        >
          Create scatter plot
        </Button>
        <Button appearance="outline" size="small" onClick={() => handlePromptClick("Insert a column of profits")}>
          Insert profits column
        </Button>
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
