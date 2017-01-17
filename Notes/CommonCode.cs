//Year at first cell
int year = sheet.GetRow(idx).GetCell(sheet.GetRow(idx).FirstCellNum).NumericCellValue;
sheet.GetRow(idx).GetCell(sheet.GetRow(idx).FirstCellNum).NumericCellValue;
//Year on headers
double year = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx).NumericCellValue;
// Qty
petroleumSheet.GetRow(counter).GetCell(cellIdx).NumericCellValue;          

// Start/End dates
new DateTime(int.Parse(year.ToString()), 1, 1);
new DateTime(int.Parse(year.ToString()), 12, 31);
petroleumDatum.StartDate = new DateTime(int.Parse(year.ToString()), 1, 1);
petroleumDatum.EndDate = new DateTime(int.Parse(year.ToString()), 12, 31);
//Name always in same row.
int rowCellHeaders = 4;//pointer to row with headers
petroleumSheet.Name = petroleumSheet.GetRow(rowCellHeaders).GetCell(cellIdx)
.StringCellValue;

//Comparing CellType
petroleumSheet.GetRow(rowIdx).GetCell(petroleumSheet.GetRow(rowIdx).FirstCellNum).CellType == CellType.Numeric

//Access first cell
petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum)

//First Cell Value
petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).FirstCellNum).StringCellValue
//Last Cell Value
petroleumSheet.GetRow(counter).GetCell(petroleumSheet.GetRow(counter).LastCellNum).StringCellValue