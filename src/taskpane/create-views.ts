/* global Excel */

export async function createViews() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Copy the contents from the original sheet to the new sheets
      const originalRange = sheet.getUsedRange();
      originalRange.load("values");
      await context.sync();

      // Sheet names
      const priceSheetName = "по цене";
      const addressSheetName = "по адресу";

      // Remove sheets if they exist
      await removeSheetIfExists(priceSheetName, context);
      await removeSheetIfExists(addressSheetName, context);

      // Create and sort the sheets
      await createAndSortSheet(sheet, priceSheetName, 3, context); // 4th column (Цена за единицу)
      await createAndSortSheet(sheet, addressSheetName, 7, context); // 8th column (address)

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// Function to remove a sheet if it exists
async function removeSheetIfExists(sheetName: string, context: Excel.RequestContext) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  await context.sync();

  if (!sheet.isNullObject) {
    sheet.delete();
    await context.sync();
  }
}

// Helper function to create and sort a sheet
async function createAndSortSheet(
  originalSheet: Excel.Worksheet,
  newSheetName: string,
  sortColumn: number,
  context: Excel.RequestContext
) {
  const newSheet = originalSheet.copy(Excel.WorksheetPositionType.after, originalSheet);
  newSheet.name = newSheetName;

  const newRange = newSheet.getUsedRange();
  const dataRange = newRange.getResizedRange(-1, 0).getOffsetRange(1, 0);
  dataRange.sort.apply([
    {
      key: sortColumn, // Column to sort by
      ascending: true,
    },
  ]);

  await context.sync();
}
