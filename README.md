# DMC Floss Inventory & Closest Color Tool

This project is a Google Sheets–based DMC floss inventory and color-matching tool.  

If you don’t have the exact floss a pattern calls for, the tool finds the closest matching DMC color you already own using RGB values and Euclidean distance.

<img width="1400" height="753" alt="example" src="https://github.com/user-attachments/assets/368d0286-eab9-4a3e-b540-53b1f2b85cd6" />


## Features

- Track owned DMC floss colors by floss number
- Visually display floss colors using RGB background fills
- Find the closest DMC color match for substitutions or planning
- Generates a copy-pastable list of unowned floss colors for convenient ordering via 123Stitch.com.

---

## How to Use

### 1. Make a copy of the Google Sheet
File -> Make a copy
https://docs.google.com/spreadsheets/d/1XTAc7M69CB0zcmJLx-77cY30G178C4qkiQFMqOsdkxI/edit?usp=sharing

### 2. Copy the Template
- Make a copy of the **Blank Pattern** sheet into your own Google Sheets file

### 3. Enter Floss Numbers
- Type or paste DMC floss numbers into the **Floss #** column
- The **Owned** and **Name** columns will auto-populate via lookup formulas

### 4. Color the Pattern Sheet
- Go to **Extensions → Apps Script**
- Run the script:
  - `color pattern sheets.gs`
- This colors the background of the floss name cells based on RGB values

### 5. (Optional) Find Closest Colors
- In **Apps Script**, run:
  - `closest color euclidean.gs`
- This calculates the nearest DMC color using RGB Euclidean distance

---

## Apps Scripts Included

- **color bg from hex.gs** – Applies RGB-based background colors to inventory
- **color pattern sheets.gs** – Applies RGB-based background colors to individual pattern sheet
- **closest color euclidean.gs** – Computes closest DMC color match

---
## Closest Color Calculation

Closest color matching is calculated using the Euclidean distance formula in RGB space:
```
d = √[(R₂ − R₁)² + (G₂ − G₁)² + (B₂ − B₁)²]
```

The owned DMC color with the smallest distance from the target RGB value is selected as the closest match.

---

## Notes & Limitations

- RGB Euclidean distance is a simple mathematical approximation and does not fully reflect human color perception. It is best suited for small projects or limited areas of color rather than large designs.
  DMC thread is relatively inexpensive, and dye lots can vary - always swatch before committing to a final color.

- This tool is intended for personal crafting, inventory management, and pattern planning.

---

## Attribution

RGB color data for DMC floss is derived from:

- **DMC Cotton Floss converted to RGB Values**  
  Source: https://github.com/adrianj/CrossStitchCreator  
  Author: adrianj  

Hex color values are generated programmatically from RGB values within this project.

