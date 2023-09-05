
/**
 * Fetch json from a URL and show data in excel
 * @customfunction
 *  
 */
async function fetchMarketDataResults() {
  debugger
  try {
    const url = "https://gist.githubusercontent.com/AppFiction/750ae98fa400bb836515eee162b39057/raw/348b1216a59fef357945710cc20270b4ede089b1/market-data.json";
    const response = await fetch(url);
    const marketData = await response.json();

    // Convert market data to an array of arrays for Excel
    const dataMatrix = [];
    for (const symbol in marketData.results) {
      const item = marketData.results[symbol];
      const rowData = marketData.inputs.map(input => item[input]);
      dataMatrix.push(rowData);
    }

    // Write data to Excel worksheet starting from the current cell
    Excel.run(function (context) {
      // const range = sheet.getRange(currentCellAddress);
      const range = context.workbook.getSelectedRange();
      const resizedRange = range.getResizedRange(dataMatrix.length - 1, dataMatrix[0].length - 1);

      resizedRange.values = dataMatrix;

      return context.sync();
    }).catch(function (error) {
      console.error("Error writing data to worksheet: ", error);
    });

    return "Results Data fetched and added to the worksheet successfully!";
  } catch (error) {
    console.error("Error fetching data: ", error);
    return "Error fetching data. Check console for details.";
  }
}

// Define the custom function using @customfunction decorator to fetch market data inputs
/** @customfunction */
async function fetchMarketDataInputs() {
  try {
    const url = "https://gist.githubusercontent.com/AppFiction/750ae98fa400bb836515eee162b39057/raw/348b1216a59fef357945710cc20270b4ede089b1/market-data.json";
    // Mock market data - Replace this with actual data from your API
    const response = await fetch(url);
    const marketData = await response.json();

    // Convert inputs array to a 2D array for Excel
    const inputsMatrix = [marketData.inputs];

    // Write inputs data to Excel worksheet starting from the current cell
    Excel.run(function (context) {
      const range = context.workbook.getSelectedRange();
      const resizedRange = range.getResizedRange(inputsMatrix.length - 1, inputsMatrix[0].length - 1);

      resizedRange.values = inputsMatrix;

      return context.sync();
    }).catch(function (error) {
      console.error("Error writing inputs data to worksheet: ", error);
    });

    return "Inputs data fetched and added to the worksheet successfully!";
  } catch (error) {
    console.error("Error fetching inputs data: ", error);
    return "Error fetching inputs data. Check console for details.";
  }
}


// Import the Excel namespace from the Office JavaScript API
/* global Excel */

// Define the custom function using @customfunction decorator
/**
 *  @customfunction 
 * */
async function fetchSupportedInputsByAnalytics() {
  try {
    // Fetch JSON data from the specified URL using a GET request
    const url = "https://gist.githubusercontent.com/AppFiction/29f83f73b8d54cc82f693b2d4f449c24/raw/c6a8f1a69509e67593255e2071216a0a2e30df02/gistfile1.json";
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`Failed to fetch data. Status: ${response.status}`);
    }

    const jsonData = await response.json();

    // Write data to Excel worksheet starting from the current cell
    Excel.run(function (context) {
      const range = context.workbook.getSelectedRange();

      // Convert JSON data to a 2D array for Excel
      const dataMatrix = jsonData.map(item => {
        return [
          item.securityld,
          item.priceDate,
          item.baseCurrency,
          item.idType2,
          item.liqCap,
          item.liqFloor,
          item.risk,
          item.riskHorizon,
          item.riskLookbackPeriod,
          item.riskReturnHorizon,
          item.useBestPracticeRealm,
          item.scenarioStartDate,
          item.scenarioEndDate,
          item.columnOrder.join(', '), // Join the columnOrder array as a string
          item.sortByColumns.join(', ') // Join the sortByColumns array as a string
        ];
      });
      const titles = Object.keys(jsonData[0]);
      const titleMatrix = [titles.map(str => {
        // Ensure that each element in the array is a string
        if (typeof str !== 'string') {
          throw new Error("Array should only contain strings.");
        }
  
        // Capitalize the first letter and make the rest of the characters lowercase
        return str.charAt(0).toUpperCase() + str.slice(1);
      })];
      const titlesRange = range.getOffsetRange(1, 0).getResizedRange(titleMatrix.length - 1, titleMatrix[0].length - 1);
      const resizedRange = range.getOffsetRange(2, 0).getResizedRange(dataMatrix.length - 1, dataMatrix[0].length - 1);


      titlesRange.values = titleMatrix;
      resizedRange.values = dataMatrix;

      return context.sync();
    }).catch(function (error) {
      console.error("Error writing data to worksheet: ", error);
    });

    return "SupportedInputsByAnalytics";
  } catch (error) {
    console.error("Error adding data: ", error);
    return "Error adding data. Check console for details.";
  }
}



