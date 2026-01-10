function fillClosestColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName("inventory");
  if (!inv) {
    throw new Error('Sheet "inventory" not found.');
  }

  const sheet = ss.getActiveSheet(); // run it on whichever pattern tab is active
  const invValues = inv.getRange("A2:G" + inv.getLastRow()).getValues(); // A..G

  // build lookup and owned list
  const flossToInfo = {};
  const ownedColors = [];
  invValues.forEach(row => {
    const owned = String(row[0]).toLowerCase();   // A
    const flossNum = row[1];                      // B
    const name = row[2];                          // C
    const r = Number(row[3]);                     // D
    const g = Number(row[4]);                     // E
    const b = Number(row[5]);                     // F
    let hex = String(row[6] || "").trim();        // G

    if (hex) {
      if (!hex.startsWith("#")) hex = "#" + hex;
      // validate 6-digit hex
      if (!/^#([0-9A-F]{6})$/i.test(hex)) {
        hex = null;
      }
    } else {
      hex = null;
    }

    if (flossNum !== "" && flossNum !== null && typeof flossNum !== 'undefined') {
      flossToInfo[flossNum] = { owned, name, r, g, b, hex };
      if (owned && hex) {
        ownedColors.push({ floss: flossNum, r, g, b, name, hex });
      } else if (owned && (r !== null && !isNaN(r))) {
        // If owned but no hex, still add by RGB if available
        ownedColors.push({ floss: flossNum, r, g, b, name, hex });
      }
    }
  });

  // nothing to compare against?
  if (ownedColors.length === 0) {
    SpreadsheetApp.getUi().alert("No owned colors with usable color data found in inventory.");
    return;
  }

  // Get floss numbers from active pattern sheet (col B)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // nothing to do
  const flossNums = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // col B

  const closestCol = [];   // column D (closest floss #)
  const nameCol = [];      // column E (closest name)
  const bgColors = [];     // backgrounds for col E
  const fontColors = [];   // font colors for col E for readability

  // helper to pick readable font color
  function readableFont(hex) {
    if (!hex) return "#000000";
    try {
      const c = hex.replace("#", "");
      const r = parseInt(c.substring(0,2),16);
      const g = parseInt(c.substring(2,4),16);
      const b = parseInt(c.substring(4,6),16);
      const brightness = (299*r + 587*g + 114*b) / 1000;
      return brightness > 128 ? "#000000" : "#FFFFFF";
    } catch (err) {
      return "#000000";
    }
  }

  flossNums.forEach(r => {
    const flossNum = r[0];
    if (flossNum && flossToInfo[flossNum]) {
      const target = flossToInfo[flossNum];

      if (target.owned === "true") {
        // already owned → leave blank (omit entry)
        closestCol.push([""]);
        nameCol.push([""]);
        bgColors.push(["#FFFFFF"]);
        fontColors.push(["#000000"]);
      } else {
        // find closest owned by RGB distance
        let minDist = Infinity;
        let closest = null;
        ownedColors.forEach(c => {
          // handle possible missing RGB
          const tr = Number(target.r || 0);
          const tg = Number(target.g || 0);
          const tb = Number(target.b || 0);
          const d = Math.pow(tr - Number(c.r || 0), 2) +
                    Math.pow(tg - Number(c.g || 0), 2) +
                    Math.pow(tb - Number(c.b || 0), 2);
          if (d < minDist) {
            minDist = d;
            closest = c;
          }
        });
        if (closest) {
          closestCol.push([closest.floss]);
          nameCol.push([closest.name || ""]);
          const bg = closest.hex || "#FFFFFF";
          bgColors.push([bg]);
          fontColors.push([readableFont(bg)]);
        } else {
          closestCol.push([""]);
          nameCol.push([""]);
          bgColors.push(["#FFFFFF"]);
          fontColors.push(["#000000"]);
        }
      }
    } else {
      closestCol.push([""]);
      nameCol.push([""]);
      bgColors.push(["#FFFFFF"]);
      fontColors.push(["#000000"]);
    }
  });

  // Write columns in batch:
  sheet.getRange(3, 4, closestCol.length, 1).setValues(closestCol);   // col D = closest floss #
  sheet.getRange(3, 5, nameCol.length, 1).setValues(nameCol);         // col E = name
  sheet.getRange(3, 5, bgColors.length, 1).setBackgrounds(bgColors);  // color column E
  sheet.getRange(3, 5, fontColors.length, 1).setFontColors(fontColors); // set readable text color
}
