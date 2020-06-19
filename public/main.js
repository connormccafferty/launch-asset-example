fin.desktop.main(async () => {
  let excelInstance;

  async function initializeExcelEvents() {
    return new Promise((resolve, reject) => {
      try {
        fin.desktop.ExcelService.addEventListener(
          "excelConnected",
          onExcelConnected
        );
        fin.desktop.ExcelService.addEventListener(
          "excelDisconnected",
          onExcelDisconnected
        );
        resolve();
      } catch (err) {
        reject(err);
      }
    });
  }

  function checkConnectionStatus() {
    fin.desktop.Excel.getConnectionStatus((connected) => {
      if (connected) {
        console.log("Already connected to Excel, synthetically raising event.");
        onExcelConnected(fin.desktop.Excel);
      } else {
        console.log("Excel not connected");
      }
    });
  }

  function onExcelConnected(data) {
    console.log(data);
    if (excelInstance) {
      return;
    }

    console.log("Excel Connected: " + data.connectionUuid);

    // Grab a snapshot of the current instance, it can change!
    excelInstance = fin.desktop.Excel;

    // excelInstance.addEventListener("workbookAdded", console.log);
    excelInstance.addEventListener("workbookOpened", console.log);
    excelInstance.addEventListener("workbookClosed", console.log);
    // excelInstance.addEventListener("workbookSaved", console.log);

    // alternative way to access the asset workbook if LEP / run isn't called on Excel first
    // fin.desktop.Excel.getWorkbooks((workbooks) => {
    //   console.log(workbooks[0]);
    // });
  }

  function onExcelDisconnected(data) {
    console.log("Excel Disconnected: " + data.connectionUuid);

    if (data.connectionUuid !== excelInstance.connectionUuid) {
      return;
    }

    // excelInstance.removeEventListener("workbookAdded", console.log);
    excelInstance.removeEventListener("workbookOpened", console.log);
    excelInstance.removeEventListener("workbookClosed", console.log);
    // excelInstance.removeEventListener("workbookSaved", console.log);

    excelInstance = undefined;

    checkConnectionStatus();
  }

  await initializeExcelEvents();
  await fin.desktop.ExcelService.init();
  checkConnectionStatus();
});
