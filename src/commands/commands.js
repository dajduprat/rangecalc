
/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

async function getSelectedRanges() {
  return Excel.run(async (context) => {
    try {
      const rangeAreas = context.workbook.getSelectedRanges();
      rangeAreas.worksheet.load({ id: true, name: true });
      const rangeCollection = rangeAreas.areas;
      rangeCollection.load({ address: true });

      try {
        await context.sync();
      } catch (error) {
        // Handle Excel "Wait until..." errors
        if (
          error instanceof Error &&
          (error.message.includes("Wait until") || error.message.includes("is currently busy"))
        ) {
          console.warn(`Excel is busy. Could not get selected ranges. ${error.message}`);
          return [];
        }

        throw error;
      }

      return rangeAreas.areas.items.map((range) => ({
        address: range.address,
        worksheet: { id: rangeAreas.worksheet.id, name: rangeAreas.worksheet.name },
      }));
    } catch (error) {
      console.error(`Error in getSelectedRanges: ${error}`);
      return [];
    }
  });
}

async function getRangeOrIntersection(rangeAddress, worksheet, context) {
  let range;

  if (rangeAddress.indexOf(",") > -1) {
    // Multiple ranges case
    range = worksheet.getRanges(rangeAddress);
  } else {
    // Single range  - get intersection with used range
    const initialRange = worksheet.getRange(rangeAddress);
    const usedRange = worksheet.getUsedRange();
    const intersection = initialRange.getIntersectionOrNullObject(usedRange);
    await context.sync();
    // If there's no intersection, fall back to the original range
    range = !intersection.isNullObject ? intersection : initialRange;
  }

  return range;
}

async function performCalculation(context, range, rangeAddress, sheet) {
  try {
    range.calculate();
    await context.sync();
  } catch (error) {
    // Handle ItemNotFound error when range reference becomes stale
    if (error?.code === 'ItemNotFound') {
      console.log('>> Range reference became stale, getting fresh reference for calculation <<');
      // Get a fresh reference to the range and try again
      const freshRange = await getRangeOrIntersection(
        rangeAddress,
        sheet,
        context
      );
      freshRange.calculate();
      await context.sync();
    } else {
      throw error;
    }
  }
}

async function executeCalculation(isSafe = false) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load(["id", "name"]);
    await context.sync();

    const ranges = await getSelectedRanges();
    const rangeAddress = ranges.map((range) => range.address).join(",");

    const loadedRanges = [];
    for (const range of ranges) {
      const excelRange = sheet.getRange(range.address);
      excelRange.load(["formulas", "address", "columnIndex", "rowIndex"]);
      loadedRanges.push(excelRange);
    }

    await context.sync();

    // First get the specified range(s)
    const range = await getRangeOrIntersection(rangeAddress, sheet, context);

    if (isSafe) {
      await performCalculation(context, range, rangeAddress, sheet);
    } else {
      range.calculate();
      await context.sync();
    }

    console.log("Calculation completed for selected range");
  });
}

function onCalculateRibbonClick(event) {
  executeCalculation(false);
  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

function onSafeCalculateRibbonClick(event) {
  executeCalculation(true);
  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}


// Register the function with Office.
Office.actions.associate("onCalculateRibbonClick", onCalculateRibbonClick);
Office.actions.associate("onSafeCalculateRibbonClick", onSafeCalculateRibbonClick);
