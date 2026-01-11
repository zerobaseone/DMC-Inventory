function fillClosestColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // prevent running on the inventory tab
  if (sheet.getName().toLowerCase() === "inventory") {
    SpreadsheetApp.getUi().alert('Please run this on a pattern sheet, not on the "inventory" tab.');
    return;
  }

  const inv = ss.getSheetByName("inventory");
  if (!inv) throw new Error('Sheet "inventory" not found.');

  const startRow = 3;

  // ---------- helpers ----------
  const isRgbValid = (r, g, b) =>
    [r, g, b].every(x => Number.isFinite(x) && x >= 0 && x <= 255);

  function readableFont(hex) {
    if (!hex) return "#000000";
    try {
      const c = hex.replace("#", "");
      const r = parseInt(c.substring(0, 2), 16);
      const g = parseInt(c.substring(2, 4), 16);
      const b = parseInt(c.substring(4, 6), 16);
      const brightness = (299 * r + 587 * g + 114 * b) / 1000;
      return brightness > 128 ? "#000000" : "#FFFFFF";
    } catch (err) {
      return "#000000";
    }
  }

  // ---------- read inventory ----------
  const invLastRow = inv.getLastRow();
  if (invLastRow < 2) return;

  // inventory: A owned (checkbox boolean), B floss#, C name, D/E/F RGB, G hex
  const invValues = inv.getRange(2, 1, invLastRow - 1, 7).getValues();

  const flossToInfo = new Map();
  const ownedColors = [];

  invValues.forEach(row => {
    const owned = row[0] === true;        // A
    const flossNum = row[1];              // B
    const name = row[2];                  // C
    const r = Number(row[3]);             // D
    const g = Number(row[4]);             // E
    const b = Number(row[5]);             // F
    let hex = String(row[6] || "").trim(); // G (hex)

    if (hex) {
      if (!hex.startsWith("#")) hex = "#" + hex;
      if (!/^#([0-9A-F]{6})$/i.test(hex)) hex = "";
    }

    if (flossNum == null || flossNum === "") return;

    const key = String(flossNum).trim();
    flossToInfo.set(key, { owned, name, r, g, b, hex });

    // if owned + valid RGB (hex is optional for coloring)
    if (owned && isRgbValid(r, g, b)) {
      ownedColors.push({
        floss: key,
        name: name || "",
        r, g, b,
        hex, // from inventory column G
      });
    }
  });

  if (ownedColors.length === 0) {
    SpreadsheetApp.getUi().alert("No owned colors with usable RGB found in inventory.");
    return;
  }

  // ---------- read pattern sheet ----------
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const flossNums = sheet.getRange(startRow, 2, numRows, 1).getValues(); // col B

  const closestCol = [];
  const nameCol = [];
  const bgColors = [];
  const fontColors = [];

  flossNums.forEach(row => {
    const key = row[0] == null ? "" : String(row[0]).trim();

    // defaults
    let outClosest = "";
    let outName = "";
    let outBg = "#FFFFFF";
    let outFont = "#000000";

    if (key && flossToInfo.has(key)) {
      const target = flossToInfo.get(key);

      // leave blank if owned
      if (!target.owned && isRgbValid(target.r, target.g, target.b)) {
        const tr = Number(target.r);
        const tg = Number(target.g);
        const tb = Number(target.b);

        let minDist = Infinity;
        let closest = null;

        for (const c of ownedColors) {
          const dr = tr - c.r;
          const dg = tg - c.g;
          const db = tb - c.b;
          const d = dr * dr + dg * dg + db * db;
          if (d < minDist) {
            minDist = d;
            closest = c;
          }
        }

        if (closest) {
          outClosest = closest.floss;
          outName = closest.name || "";

          // only color if we have a valid hex in inventory
          if (closest.hex) {
            outBg = closest.hex;
            outFont = readableFont(outBg);
          }
        }
      }
    }

    closestCol.push([outClosest]);
    nameCol.push([outName]);
    bgColors.push([outBg]);
    fontColors.push([outFont]);
  });

  // ---------- write results ----------
  sheet.getRange(startRow, 4, numRows, 1).setValues(closestCol);     // col D
  sheet.getRange(startRow, 5, numRows, 1).setValues(nameCol);        // col E
  sheet.getRange(startRow, 5, numRows, 1).setBackgrounds(bgColors);  // col E
  sheet.getRange(startRow, 5, numRows, 1).setFontColors(fontColors); // col E
}
