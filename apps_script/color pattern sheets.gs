function colorPatternNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("inventory");
  const sheet = ss.getActiveSheet(); // run it on whichever pattern tab is active

  // Get inventory data: B = Floss#, G = HEX
  const masterValues = master.getRange("B2:G" + master.getLastRow()).getValues();
  const flossToHex = {};
  masterValues.forEach(row => {
    let flossNum = row[0];   // column B (Floss#)
    let hex = row[5];        // column G (HEX)
    if (flossNum && hex) {
      hex = String(hex).trim();
      if (!hex.startsWith("#")) hex = "#" + hex;
      flossToHex[flossNum] = hex;
    }
  });

  // Helper to pick black or white font for readability
  function readableFont(hex) {
    if (!hex) return "#000000";
    const c = hex.replace("#", "");
    const r = parseInt(c.substring(0, 2), 16);
    const g = parseInt(c.substring(2, 4), 16);
    const b = parseInt(c.substring(4, 6), 16);
    const brightness = (299 * r + 587 * g + 114 * b) / 1000;
    return brightness > 128 ? "#000000" : "#FFFFFF";
  }

  // Loop over active sheet (pattern tab)
  const lastRow = sheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) { // skip header
    const flossNum = sheet.getRange(i, 2).getValue(); // column B = Floss#
    const cell = sheet.getRange(i, 3); // column C = Name
    if (flossNum && flossToHex[flossNum]) {
      const hex = flossToHex[flossNum];
      cell.setBackground(hex);
      cell.setFontColor(readableFont(hex));
    } else {
      cell.setBackground(null); // clear if no match
      cell.setFontColor("#000000");
    }
  }
}
