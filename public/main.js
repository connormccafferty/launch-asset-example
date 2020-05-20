fin.desktop.main(() => {
  let excelInstance;

  function initializeExcelEvents() {
    fin.desktop.ExcelService.addEventListener(
      "excelConnected",
      onExcelConnected
    );
    fin.desktop.ExcelService.addEventListener(
      "excelDisconnected",
      onExcelDisconnected
    );
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

  initializeExcelEvents();

  fin.desktop.ExcelService.init()
    .then(checkConnectionStatus)
    .catch((err) => console.error(err));
});
