import * as React from "react";
import { Button, Input, Text, Card, Image } from "@fluentui/react-components";
import { useStyles } from "./App.styles";
/* global Excel */
/* global clearInterval, setInterval, setTimeout, NodeJS */

// Processing time in milliseconds (0 = no delay)
// const PROCESSING_DELAY_MS = 3000;
const PROCESSING_DELAY_MS = 0;

interface AppProps {}

interface ChatMessage {
  id: number;
  type: "text" | "table" | "image";
  content: string | string[][] | string; // text content, table data, or image path
  isUser: boolean;
  isLoading?: boolean; // 新增加载状态标志
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

async function prepareChartData(context: Excel.RequestContext): Promise<Excel.Range> {
  const table = await getFirstTable(context);
  table.load(["columns", "name"]);
  await context.sync();

  // Find the Sales and Costs columns
  const salesColumn = table.columns.getItemOrNullObject("Sales");
  const costsColumn = table.columns.getItemOrNullObject("Costs");
  salesColumn.load("index");
  costsColumn.load("index");
  await context.sync();

  if (salesColumn.isNullObject || costsColumn.isNullObject) {
    throw new Error("Could not find 'Sales' or 'Costs' columns in the table");
  }

  // Get the data range for both columns
  const salesRange = salesColumn.getDataBodyRange();
  const costsRange = costsColumn.getDataBodyRange();
  // Load the address property before using it
  salesRange.load("address");
  costsRange.load("address");
  await context.sync();

  // Create a chart source range that includes both columns
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

  // Format axis labels to show values in thousands
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
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.add(Excel.ChartType.xyscatter, chartDataRange, Excel.ChartSeriesBy.columns);
    formatScatterChart(chart);
    await context.sync();

    const chartImage = chart.getImage();
    await context.sync();
    imageBase64 = "data:image/png;base64," + chartImage.value;

    // Delete the chart after getting the image
    chart.delete();
    await context.sync();
  });

  return imageBase64;
}

async function addProfitColumn(): Promise<void> {
  await Excel.run(async (context) => {
    // Get the first table
    const table = await getFirstTable(context);

    // Get the table data to determine row count
    const dataRange = table.getDataBodyRange();
    dataRange.load("rowCount");
    await context.sync();

    // Create the formula to calculate profit
    const profitFormula = "=[@Sales]-[@Costs]";

    // Create an array with the header and formulas for each row
    const columnData = [["Profits"]];

    // Add formula for each row in the table
    for (let i = 0; i < dataRange.rowCount; i++) {
      columnData.push([profitFormula]);
    }

    // Add the column with the formula
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
  const [isProcessing, setIsProcessing] = React.useState(false); // 新增处理状态
  const chatContainerRef = React.useRef(null);
  const loadingDotsIntervalRef = React.useRef<NodeJS.Timeout | null>(null); // 用于存储定时器引用

  // 添加一个模拟延迟的函数
  const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

  const handleSendMessage = async () => {
    if (inputText.trim() === "" || isProcessing) return;

    // Add user message
    const newUserMessage: ChatMessage = {
      id: messages.length + 1,
      type: "text",
      content: inputText,
      isUser: true,
    };

    // 添加一个加载中的消息
    const loadingMessage: ChatMessage = {
      id: messages.length + 2,
      type: "text",
      content: "Processing",
      isUser: false,
      isLoading: true,
    };

    setMessages([...messages, newUserMessage, loadingMessage]);
    setInputText("");
    setIsProcessing(true);

    // 创建一个更新加载消息的函数
    let dots = 0;
    const updateLoadingMessage = () => {
      dots = (dots + 1) % 4;
      const dotsText = ".".repeat(dots);

      setMessages((prevMessages) =>
        prevMessages.map((msg) => (msg.isLoading ? { ...msg, content: `Processing${dotsText}` } : msg))
      );
    };

    // Only start the loading animation if there's a delay
    if (PROCESSING_DELAY_MS > 0) {
      // 启动定时器，每500毫秒更新一次加载消息
      loadingDotsIntervalRef.current = setInterval(updateLoadingMessage, 500);

      // 等待设定的延迟时间
      await sleep(PROCESSING_DELAY_MS);
    }

    // 检查for the specific message and respond accordingly
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

    // 清除定时器
    if (loadingDotsIntervalRef.current) {
      clearInterval(loadingDotsIntervalRef.current);
      loadingDotsIntervalRef.current = null;
    }

    // 移除加载消息并添加实际响应
    setMessages((prevMessages) => prevMessages.filter((msg) => !msg.isLoading).concat(botResponse));

    setIsProcessing(false);
  };

  // 清理定时器
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

  const handleAddProfitColumn = async () => {
    try {
      await addProfitColumn();

      const successMessage: ChatMessage = {
        id: messages.length + 1,
        type: "text",
        content: "Profit column added successfully!",
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
    // 如果是加载中的消息，显示加载动画
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

  // Function to handle clicking on a recommended prompt button
  const handlePromptClick = (promptText: string) => {
    setInputText(promptText);
  };

  // Add a handler for the insert chart button
  const handleInsertChart = async () => {
    try {
      await createScatterChartInSheet();

      const successMessage: ChatMessage = {
        id: messages.length + 1,
        type: "text",
        content: "散点图已成功插入到工作表中！",
        isUser: false,
      };
      setMessages((prev) => [...prev, successMessage]);
    } catch (error) {
      const errorMessage: ChatMessage = {
        id: messages.length + 1,
        type: "text",
        content: `错误: ${error.toString()}`,
        isUser: false,
      };
      setMessages((prev) => [...prev, errorMessage]);
    }
  };

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
        <Button
          appearance="outline"
          size="small"
          onClick={() => handlePromptClick("Sort the table by sales in descending order")}
          disabled={isProcessing}
        >
          Sort table by sales
        </Button>
        <Button
          appearance="outline"
          size="small"
          onClick={() => handlePromptClick("Create a scatter plot of sales and costs")}
          disabled={isProcessing}
        >
          Create scatter plot
        </Button>
        <Button
          appearance="outline"
          size="small"
          onClick={() => handlePromptClick("Insert a column of profits")}
          disabled={isProcessing}
        >
          Insert profits column
        </Button>
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
