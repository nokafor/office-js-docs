# Calling for suggestion on method names

Below are the two of the many methods that is being planned for Excel1.4 release. We need your specific input on the naming suggestion. 

Please let us know your [feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=excel-1.4). 

### Proposed method-1: worksheet.getBoundingRange 

This method gets the smallest range object that encompasses the provided ranges. For example, the bounding range between `"B2:C5"` and `"D10:E15"` is `"B2:E16"`.

This function is a superset of the [getBoundingRect](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/range.md#getboundingrectanotherrange-range-or-string) method on the Range object. That is, `worksheet.getBoundingRange([range1, range2])` is a identical to for `range1.getBoundingRect(range2)`. However, this version of the function is also able to accept more than two Range objects.

```
/// </summary>
/// <param name="ranges">An array of Range objects or addresses, 
/// which will be fully encompassed by the resulting range.</param>
/// <returns> Range </returns>

Range GetBoundingRange(Array<Excel.Range|string> ranges);
```

Alternate names under consideration:  `worksheet.getRangeBetween`

### Proposed method-2: Worksheet.getRangeR1C1  

Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.

```
/// <summary>
/// Gets the range object beginning at a particular row index and column index, 
/// and spanning a certain number of rows and columns.
/// </summary>
/// <param name="startRow">Start row (zero-indexed).</param>
/// <param name="startColumn">Start column (zero-indexed).</param>
/// <param name="rowCount">Number of rows to include in the range.</param>
/// <param name="columnCount">Number of columns to include in the range.</param>
/// <returns> Range </returns>

    Range getRangeR1C1(startRow, startColumn, rowCount, columnCount);
```

Other names under consideration: `getCells()`, `getRangeByIndexes()`
