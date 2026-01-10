function colorDescriptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

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

  for (let i = 2; i <= lastRow; i++) { // start at row 2 to skip header
    let hex = sheet.getRange(i, 7).getValue(); // column G = HEX
    if (hex) {
      hex = String(hex).trim();
      if (!hex.startsWith("#")) hex = "#" + hex;
      const cell = sheet.getRange(i, 3); // Description column C
      cell.setBackground(hex);
      cell.setFontColor(readableFont(hex));
    } else {
      const cell = sheet.getRange(i, 3);
      cell.setBackground(null);
      cell.setFontColor("#000000");
    }
  }
}

function onEdit(e) {
  const sheet = e.range.getSheet();

  // only run if the edit is in column G (7)
  if (e.range.getColumn() === 7) {
    const row = e.range.getRow();
    let hex = e.range.getValue();
    const cell = sheet.getRange(row, 3); // Description column C

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

    if (hex) {
      hex = String(hex).trim();
      if (!hex.startsWith("#")) hex = "#" + hex;
      cell.setBackground(hex);
      cell.setFontColor(readableFont(hex));
    } else {
      cell.setBackground(null);
      cell.setFontColor("#000000");
    }
  }
}
