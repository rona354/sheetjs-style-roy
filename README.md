<h1 align='center'>sheetjs-style-roy</h1>
<p align='center'>
  <!-- <a href="https://travis-ci.com/Yoshino-UI/Yoshino/">
    <img src="https://travis-ci.com/Yoshino-UI/Yoshino.svg" alt="travis ci badge">
  </a> -->
  <img src='https://img.shields.io/npm/v/sheetjs-style.svg?style=flat-square' alt="version">
  <img src='https://img.shields.io/npm/l/sheetjs-style.svg' alt="license">
  <img src='https://img.shields.io/npm/dt/sheetjs-style.svg?style=flat-square' alt="downloads">
  <img src='https://img.shields.io/npm/dm/sheetjs-style.svg?style=flat-square' alt="downloads-month">
</p>
<p align='center'>support set and get cell style for sheetjs!</p>
<p align='center'>API is the same as sheetjs!</p>

## install
```
npm install sheetjs-style-roy
```

## How to Use?
Please read [SheetJs Documents](https://github.com/SheetJS/sheetjs/blob/3468395494c450ea8ba7e20afb1bd6127f516ccd/README.md)!

## Just like [sheetjs-style](https://github.com/ShanaMaid/sheetjs-style) but come with new features...
for example:
```js
  const selectedSheet = "Sheet1";
  const secondSheet = "Sheet2";
  const firstCol = "A";
  const firstRow = "1"; 

  const mainSheet = getSheets(Workbook, selectedSheet);
  const targetSheet = getSheets(Workbook, secondSheet);

  XLSX.utils.sheet_add_sheet(   // replace a sheet with another sheet
    choosenSheet,
    targetSheet,
    {
      header: 1,
      origin: firstCol + firstRow, // e.g. "A1" to skip the header
    }
  );

  /** Attempts to write wb to filename. In browser-based environments, it will attempt to force a client-side download. */
  XLSX.writeFile(Workbook, "test_mutated_data.xlsx", {
    cellStyles: true,
  });
```

# Cell Styles

Cell styles are specified by a style object that roughly parallels the OpenXML structure.  The style object has five
top-level attributes: `fill`, `font`, `numFmt`, `alignment`, and `border`.


| Style Attribute | Sub Attributes | Values |
| :-------------- | :------------- | :------------- |
| fill            | patternType    |  `"solid"` or `"none"`
|                 | fgColor        |  `COLOR_SPEC`
|                 | bgColor        |  `COLOR_SPEC`
| font            | name           |  `"Calibri"` // default
|                 | sz             |  `"11"` // font size in points
|                 | color          |  `COLOR_SPEC`
|                 | bold           |  `true` or `false`
|                 | underline      |  `true` or `false`
|                 | italic         |  `true` or `false`
|                 | strike         |  `true` or `false`
|                 | outline        |  `true` or `false`
|                 | shadow         |  `true` or `false`
|                 | vertAlign      |  `true` or `false`
| numFmt          |                |  `"0"`  // integer index to built in formats, see StyleBuilder.SSF property
|                 |                |  `"0.00%"` // string matching a built-in format, see StyleBuilder.SSF
|                 |                |  `"0.0%"`  // string specifying a custom format
|                 |                |  `"0.00%;\\(0.00%\\);\\-;@"` // string specifying a custom format, escaping special characters
|                 |                |  `"m/dd/yy"` // string a date format using Excel's format notation
| alignment       | vertical       | `"bottom"` or `"center"` or `"top"`
|                 | horizontal     | `"left"` or `"center"` or `"right"`
|                 | wrapText       |  `true ` or ` false`
|                 | readingOrder   |  `2` // for right-to-left
|                 | textRotation   | Number from `0` to `180` or `255` (default is `0`)
|                 |                |  `90` is rotated up 90 degrees
|                 |                |  `45` is rotated up 45 degrees
|                 |                | `135` is rotated down 45 degrees
|                 |                | `180` is rotated down 180 degrees
|                 |                | `255` is special,  aligned vertically
| border          | top            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`
|                 | bottom         | `{ style: BORDER_STYLE, color: COLOR_SPEC }`
|                 | left           | `{ style: BORDER_STYLE, color: COLOR_SPEC }`
|                 | right          | `{ style: BORDER_STYLE, color: COLOR_SPEC }`
|                 | diagonal       | `{ style: BORDER_STYLE, color: COLOR_SPEC }`
|                 | diagonalUp     | `true` or `false`
|                 | diagonalDown   | `true` or `false`

**COLOR_SPEC**: Colors for `fill`, `font`, and `border` are specified as objects, either:
* `{ auto: 1}` specifying automatic values
* `{ rgb: "FFFFAA00" }` specifying a hex ARGB value
* `{ theme: "1", tint: "-0.25"}` specifying an integer index to a theme color and a tint value (default 0)
* `{ indexed: 64}` default value for `fill.bgColor`

**BORDER_STYLE**: Border style is a string value which may take on one of the following values:
 * `thin`
 * `medium`
 * `thick`
 * `dotted`
 * `hair`
 * `dashed`
 * `mediumDashed`
 * `dashDot`
 * `mediumDashDot`
 * `dashDotDot`
 * `mediumDashDotDot`
 * `slantDashDot`


Borders for merged areas are specified for each cell within the merged area.  So to apply a box border to a merged area of 3x3 cells, border styles would need to be specified for eight different cells:
* left borders for the three cells on the left,
* right borders for the cells on the right
* top borders for the cells on the top
* bottom borders for the cells on the left
 
## Thanks
[sheetjs-style](https://github.com/ShanaMaid/sheetjs-style)