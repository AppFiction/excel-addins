/* global clearInterval, console, setInterval */

/**
 * Fetch json from a URL and show data in excel
 * @customfunction
 *  
 */
async function fetchMarketData() {
  debugger
  const gistRawURL = "https://gist.githubusercontent.com/AppFiction/750ae98fa400bb836515eee162b39057/raw/348b1216a59fef357945710cc20270b4ede089b1/market-data.json";
  try {

    // Mock market data - Replace this with actual data from your API
    const response = await fetch(gistRawURL);
    const marketData = await response.json();

    console.log("RES", response);

    // Get the current cell address from the invocation object
    // const currentCellAddress = Excel.Script.CustomFunctions.invocation.address;

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

    return "Data fetched and added to the worksheet successfully!";
  } catch (error) {
    console.error("Error fetching data: ", error);
    return "Error fetching data. Check console for details.";
  }
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
export function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

// Function to introduce a delay (mocking an asynchronous operation)
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
export function currentTime() {
  return new Date().toLocaleTimeString();
}

