/* global Excel console */

export async function insertText(text: string) {
  // Write text to the top left cell.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function sortTable(tableName: string, key: number, ascending: boolean) {
  await Excel.run(async (context) => {
    // const sheet = context.workbook.worksheets.getItem("Sample");
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem(tableName);

    const sortFields = [
      {
        key: key,
        ascending: ascending,
      },
    ];
    table.sort.apply(sortFields);

    await context.sync();
  });
}
