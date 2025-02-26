import * as React from "react";
import { Button, Input, Text, Card, Image } from "@fluentui/react-components";
import { useStyles } from "./App.styles";
/* global Excel */
/* global clearInterval, setInterval, setTimeout, NodeJS */

// Processing time in milliseconds (0 = no delay)
const PROCESSING_DELAY_MS = 3000;

interface AppProps {}

interface ChatMessage {
  id: number;
  type: "text" | "table" | "image";
  content: string | string[][] | string; // text content, table data, or image path
  isUser: boolean;
  isLoading?: boolean;
}

async function sortTableBySales(context: Excel.RequestContext, table: Excel.Table): Promise<void> {
  const salesColumn = table.columns.getItemOrNullObject("sales");
  salesColumn.load("index");
  await context.sync();

  if (salesColumn.isNullObject) {
    throw new Error("No 'sales' column found in the table");
  }
  table.sort.apply([{ key: salesColumn.index, ascending: false }]);
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

    const sourceRange = sourceTable.getRange();
    sourceRange.load(["values", "rowCount", "columnCount"]);
    await context.sync();
    const tempTable = tempSheet.tables.add(
      tempSheet.getRange("A1").getResizedRange(sourceRange.rowCount - 1, sourceRange.columnCount - 1),
      true
    );
    tempTable.getRange().values = sourceRange.values;
    await context.sync();
    await sortTableBySales(context, tempTable);
    const tempRange = tempTable.getRange();
    tempRange.load("values");
    await context.sync();

    // Convert to string array for our message
    tableData = tempRange.values.map((row) => row.map((cell) => (cell !== null ? String(cell) : "")));

    tempSheet.delete();
    await context.sync();
  });

  return tableData;
}

async function prepareChartData(context: Excel.RequestContext): Promise<Excel.Range> {
  const table = await getFirstTable(context);
  table.load(["columns", "name"]);
  await context.sync();
  const salesColumn = table.columns.getItemOrNullObject("Sales");
  const costsColumn = table.columns.getItemOrNullObject("Costs");
  salesColumn.load("index");
  costsColumn.load("index");
  await context.sync();

  if (salesColumn.isNullObject || costsColumn.isNullObject) {
    throw new Error("Could not find 'Sales' or 'Costs' columns in the table");
  }
  const salesRange = salesColumn.getDataBodyRange();
  const costsRange = costsColumn.getDataBodyRange();
  salesRange.load("address");
  costsRange.load("address");
  await context.sync();
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const chartRange = sheet.getRange(`${salesRange.address}:${costsRange.address}`);
  return chartRange;
}

function formatScatterChart(chart: Excel.Chart) {
  chart.title.text = "'Costs' by 'Sales'";
  chart.legend.visible = false;

  chart.axes.valueAxis.title.text = "Costs";
  chart.axes.valueAxis.title.visible = true;
  chart.axes.categoryAxis.title.text = "Sales";
  chart.axes.categoryAxis.title.visible = true;

  chart.axes.valueAxis.majorGridlines.visible = true;
  chart.axes.categoryAxis.majorGridlines.visible = true;

  // show values in thousands
  chart.axes.valueAxis.numberFormat = "#,##0,K";
  chart.axes.categoryAxis.numberFormat = "#,##0,K";
}

async function createScatterChartInSheet() {
  await Excel.run(async (context) => {
    const chartDataRange = await prepareChartData(context);
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.add(Excel.ChartType.xyscatter, chartDataRange, Excel.ChartSeriesBy.columns);
    formatScatterChart(chart);
    chart.setPosition("B35", "I50");
    await context.sync();
  });
}

async function createSalesCostsScatterChart(): Promise<string> {
  let imageBase64 = "";

  await Excel.run(async (context) => {
    const chartDataRange = await prepareChartData(context);

    // Create a temporary sheet for the chart
    const tempSheetName = `TempChartSheet_${Date.now()}`;
    const tempSheet = context.workbook.worksheets.add(tempSheetName);

    // Copy the data to the temp sheet
    const dataValues = chartDataRange.load("values");
    await context.sync();

    const tempRange = tempSheet
      .getRange("A1")
      .getResizedRange(dataValues.values.length - 1, dataValues.values[0].length - 1);
    tempRange.values = dataValues.values;
    await context.sync();

    // Create chart on the temp sheet
    const chart = tempSheet.charts.add(Excel.ChartType.xyscatter, tempRange, Excel.ChartSeriesBy.columns);
    formatScatterChart(chart);
    await context.sync();

    const chartImage = chart.getImage();
    await context.sync();
    imageBase64 = "data:image/png;base64," + chartImage.value;

    tempSheet.delete();
    await context.sync();
  });

  return imageBase64;
}

async function addProfitColumn(): Promise<void> {
  await Excel.run(async (context) => {
    const table = await getFirstTable(context);
    const dataRange = table.getDataBodyRange();
    dataRange.load("rowCount");
    await context.sync();

    const profitFormula = "=[@Sales]-[@Costs]";
    const columnData = [["Profits"]];
    for (let i = 0; i < dataRange.rowCount; i++) {
      columnData.push([profitFormula]);
    }
    const newColumn = table.columns.add(null, columnData);

    // Format the column as integers (no decimals)
    const profitRange = newColumn.getDataBodyRange();
    profitRange.numberFormat = [["0"]];

    // Auto-fit the columns for better visibility
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getUsedRange().format.autofitColumns();

    await context.sync();
  });
}

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<ChatMessage[]>([
    { id: 1, type: "text", content: "Hello! How can I help you today?", isUser: false },
  ]);
  const [inputText, setInputText] = React.useState("");
  const [isProcessing, setIsProcessing] = React.useState(false);
  const chatContainerRef = React.useRef(null);
  const loadingDotsIntervalRef = React.useRef<NodeJS.Timeout | null>(null);

  const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

  const addMessage = (message: ChatMessage) => {
    setMessages((prev) => [...prev, message]);
  };

  const handleSendMessage = async () => {
    if (inputText.trim() === "" || isProcessing) return;

    // Add user message
    addMessage({
      id: messages.length + 1,
      type: "text",
      content: inputText,
      isUser: true,
    });

    // Add a loading message
    addMessage({
      id: messages.length + 2,
      type: "text",
      content: "Processing",
      isUser: false,
      isLoading: true,
    });

    setInputText("");
    setIsProcessing(true);

    let dots = 0;
    const updateLoadingMessage = () => {
      dots = (dots + 1) % 4;
      setMessages((prev) =>
        prev.map((msg) => (msg.isLoading ? { ...msg, content: `Processing${".".repeat(dots)}` } : msg))
      );
    };

    if (PROCESSING_DELAY_MS > 0) {
      loadingDotsIntervalRef.current = setInterval(updateLoadingMessage, 500);
      await sleep(PROCESSING_DELAY_MS);
    }

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
      try {
        const chartImageBase64 = await createSalesCostsScatterChart();

        botResponse = {
          id: messages.length + 2,
          type: "image",
          content: chartImageBase64,
          isUser: false,
        };
      } catch (error) {
        botResponse = {
          id: messages.length + 2,
          type: "text",
          content: `Error creating chart: ${error.toString()}`,
          isUser: false,
        };
      }
    } else if (normalizedInput === "Insert a column of profits") {
      botResponse = {
        id: messages.length + 2,
        type: "text",
        content:
          "I'll add a Profit column using the formula: Profit = Sales - Cost. This will calculate the profit for each product.",
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

    if (loadingDotsIntervalRef.current) {
      clearInterval(loadingDotsIntervalRef.current);
      loadingDotsIntervalRef.current = null;
    }
    setMessages((prev) => prev.filter((msg) => !msg.isLoading).concat(botResponse));

    setIsProcessing(false);
  };

  React.useEffect(() => {
    return () => {
      if (loadingDotsIntervalRef.current) {
        clearInterval(loadingDotsIntervalRef.current);
      }
    };
  }, []);

  const handleApplySortToWorksheet = async () => {
    try {
      await Excel.run(async (context) => {
        const table = await getFirstTable(context);
        await sortTableBySales(context, table);
      });

      addMessage({
        id: messages.length + 1,
        type: "text",
        content: "Table sorted successfully!",
        isUser: false,
      });
    } catch (error) {
      addMessage({
        id: messages.length + 1,
        type: "text",
        content: `Error: ${error.toString()}`,
        isUser: false,
      });
    }
  };

  const handleAddProfitColumn = async () => {
    try {
      await addProfitColumn();

      addMessage({
        id: messages.length + 1,
        type: "text",
        content: "Profit column added successfully!",
        isUser: false,
      });
    } catch (error) {
      addMessage({
        id: messages.length + 1,
        type: "text",
        content: `Error: ${error.toString()}`,
        isUser: false,
      });
    }
  };

  const renderMessageContent = (message: ChatMessage) => {
    if (message.isLoading) {
      return <Text>{message.content}</Text>;
    }

    switch (message.type) {
      case "text": {
        const isProfitMessage = (message.content as string).includes("Profit = Sales - Cost");
        return (
          <div>
            <Text>{message.content as string}</Text>
            {isProfitMessage && (
              <div className={styles.tableActions} style={{ marginTop: "8px" }}>
                <Button appearance="primary" size="small" onClick={handleAddProfitColumn}>
                  插入公式
                </Button>
              </div>
            )}
          </div>
        );
      }

      case "image": {
        const isChartImage = (message.content as string).startsWith("data:image");
        return (
          <div>
            <Image src={message.content as string} alt="Chat image" className={styles.messageImage} />
            {isChartImage && (
              <div className={styles.tableActions}>
                <Button appearance="primary" size="small" onClick={handleInsertChart}>
                  插入
                </Button>
              </div>
            )}
          </div>
        );
      }

      case "table": {
        const tableData = message.content as string[][];
        return (
          <div>
            <div className={styles.tableContainer}>
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
            </div>
            <div className={styles.tableActions}>
              <Button appearance="primary" size="small" onClick={handleApplySortToWorksheet}>
                应用操作
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

  const handlePromptClick = (promptText: string) => {
    setInputText(promptText);
  };

  const handleInsertChart = async () => {
    try {
      await createScatterChartInSheet();

      addMessage({
        id: messages.length + 1,
        type: "text",
        content: "The chart has been inserted into the worksheet.",
        isUser: false,
      });
    } catch (error) {
      addMessage({
        id: messages.length + 1,
        type: "text",
        content: `错误: ${error.toString()}`,
        isUser: false,
      });
    }
  };

  const promptOptions = [
    { text: "Sort table by sales", prompt: "Sort the table by sales in descending order" },
    { text: "Create scatter plot", prompt: "Create a scatter plot of sales and costs" },
    { text: "Insert profits column", prompt: "Insert a column of profits" },
  ];

  return (
    <div className={styles.root}>
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
        {promptOptions.map((option, index) => (
          <Button
            key={index}
            appearance="outline"
            size="small"
            onClick={() => handlePromptClick(option.prompt)}
            disabled={isProcessing}
          >
            {option.text}
          </Button>
        ))}
      </div>

      <div className={styles.inputContainer}>
        <Input
          className={styles.inputField}
          value={inputText}
          onChange={(_e, data) => setInputText(data.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !isProcessing) {
              handleSendMessage();
            }
          }}
          placeholder="Type your message here..."
          disabled={isProcessing}
        />
        <Button appearance="primary" onClick={handleSendMessage} disabled={isProcessing}>
          Send
        </Button>
      </div>
    </div>
  );
};

export default App;
