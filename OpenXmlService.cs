using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Services
{
    public class OpenXmlService
    {
        private readonly string _excelFilePath;

        public OpenXmlService(string excelTemplateFilePath)
        {
            _excelFilePath = HttpContext.Current.Server.MapPath(excelTemplateFilePath);
        }

        public Stream Export(List<string[]> excelRows)
        {
            //  ***********************************
            //              USE OPENXML
            //  ***********************************
            //  Get Excel Template file (.xlsx)
            string filename = Path.GetFileName(_excelFilePath);
            string newfilepath = Path.Combine(Path.GetTempPath(), filename);
            //  Delete if previous file exists already
            try {
                if (File.Exists(newfilepath)) {
                    File.Delete(newfilepath);
                }
            } catch { }
            //  Copy file at user's TEMP folder
            File.Copy(_excelFilePath, newfilepath, true);
            //  Open copied file
            SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(newfilepath, true);

            //  Select one of the worksheets given its sheetId
            //Sheet sheet = spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>().ToList()[sheetId];
            //  Or simply get 1st worksheet of document
            Sheet sheet = spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>().First();
            Worksheet worksheet = ((WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheet.Id)).Worksheet;

            //  Find Cell that contains "1" and update multiple rows starting from that cell
            KeyValuePair<Cell, Row> RowToUpdate = FindCellContainsValue(worksheet, 0, 2000, 0, 0, "1");
            UpdateMultipleRows(worksheet, excelRows, RowToUpdate.Value.RowIndex.Value, 1);
            //  Close the modified copied Excel file
            spreadSheet.Close();
            //  Get the modified file
            var file = File.ReadAllBytes(newfilepath);
            //  Delete copied at TEMP file. Try..catch is used just for sure. It is not necessary to delete it after all.
            try {
                File.Delete(newfilepath);
            } catch { }
            //  Returns modified file as a MemoryStream
            return new MemoryStream(file);
        }

        #region Update xls cells

        /// <summary>
        /// Updates a cell given the worksheet, text to set as value and row / column ID
        /// </summary>
        /// <param name="worksheet">Worksheet object that has cell we need to update</param>
        /// <param name="text">Text to update cell value</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="columnIndex">Column Index</param>
        private static void UpdateCell(Worksheet worksheet, string text,
            uint rowIndex, uint columnIndex)
        {
            if (worksheet != null) {
                Cell cell = GetCell(worksheet, GetExcelColumnName(columnIndex), rowIndex);
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(text));
                cell.DataType = CellValues.String;
            }
        }

        /// <summary>
        /// Updates multiple columns for a row given the worksheet, table of strings and row / column ID of the first cell to update. It updates as many columns as the length of the table of strings
        /// </summary>
        /// <param name="worksheet">Worksheet object that has cells we need to update</param>
        /// <param name="text">Table of strings to update cell values</param>
        /// <param name="firstRowIndex">Row Index of the first cell</param>
        /// <param name="firstColumnIndex">Column Index of the first cell</param>
        private static void UpdateColumnsForRow(Worksheet worksheet, string[] text,
            uint firstRowIndex, uint firstColumnIndex)
        {
            if (worksheet != null) {
                for (uint i = 0; i < text.Length; ++i) {
                    UpdateCell(worksheet, text[i], firstRowIndex, firstColumnIndex + i);
                }
            }
        }
        /// <summary>
        /// Updates multiple rows given the first left upper cell index.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="excelRows">List of tables of strings to hold the data to update</param>
        /// <param name="firstRowIndex">Row Index of the left upper cell</param>
        /// <param name="firstColumnIndex">Column Index of the left upper cell</param>
        private static void UpdateMultipleRows(Worksheet worksheet, List<string[]> excelRows, uint firstRowIndex, uint firstColumnIndex)
        {
            if (worksheet != null) {
                InsertRows(firstRowIndex, worksheet.GetFirstChild<SheetData>(), excelRows.Count - 1, excelRows[0].Length);
                for (uint i = 0; i < excelRows.Count; ++i) {
                    UpdateColumnsForRow(worksheet, excelRows[(int)i], firstRowIndex + i, firstColumnIndex);
                }
            }
        }

        /// <summary>
        /// Gets the cell at the specified column and row given a worksheet, a column name, and a row index
        /// </summary>
        /// <param name="worksheet">Worksheet object that has cell we need to get</param>
        /// <param name="columnName">Column Index. It is a string as in spreadsheet (eg. "D")</param>
        /// <param name="rowIndex">Row Index</param>
        /// <returns></returns>
        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            try {
                //  Ελέγχει αν το Row περιέχει το κελί που θέλουμε να εισάγουμε τιμή. Αν δεν το περιέχει, γίνεται append
                if (!row.Elements<Cell>().Any(c => c.CellReference.Value == columnName + rowIndex)) {
                    row.Append(new Cell() {
                        CellFormula = new CellFormula(),
                        CellReference = columnName + rowIndex
                    });
                }
                return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).First();
            } catch (Exception e) {
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        /// <param name="textToSearch"></param>
        /// <returns></returns>
        private static KeyValuePair<Cell, Row> FindCellContainsValue(Worksheet worksheet, uint startRow, uint endRow, uint startColumn, uint endColumn, string textToSearch)
        {
            for (uint rowIndex = startRow; rowIndex <= endRow; ++rowIndex) {
                for (uint columnIndex = startColumn; columnIndex <= endColumn; ++columnIndex) {
                    Cell cell = GetCell(worksheet, GetExcelColumnName(columnIndex), rowIndex);
                    if (cell != null && cell.CellValue != null && cell.CellValue.InnerText == textToSearch) {
                        cell.CellValue = new CellValue(string.Empty);
                        return new KeyValuePair<Cell, Row>(cell, GetRow(worksheet, rowIndex));
                    }
                }
            }
            return new KeyValuePair<Cell, Row>();
        }
        /// <summary>
        /// Returns ALL the cells that contain string value. WARNING: It is super slow...
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        /// <param name="textToSearch"></param>
        /// <param name="sharedTable"></param>
        /// <returns></returns>
        private static List<KeyValuePair<Cell, Row>> FindAllCellsContainsValue(Worksheet worksheet, uint startRow, uint endRow, uint startColumn, uint endColumn, string textToSearch, SharedStringTable sharedTable)
        {
            List<KeyValuePair<Cell, Row>> CellsContainsValue = new List<KeyValuePair<Cell, Row>>();
            for (uint rowIndex = startRow; rowIndex <= endRow; ++rowIndex) {
                for (uint columnIndex = startColumn; columnIndex <= endColumn; ++columnIndex) {
                    Cell cell = GetCell(worksheet, GetExcelColumnName(columnIndex), rowIndex);
                    if (cell != null && cell.CellValue != null) {
                        int id = -1;
                        if (Int32.TryParse(cell.InnerText, out id)) {
                            SharedStringItem item = sharedTable.Elements<SharedStringItem>().ElementAt(id);
                            if ((item.Text != null && item.Text.Text.Contains(textToSearch))
                                || (item.InnerText != null && item.InnerText.Contains(textToSearch))
                                || (item.InnerXml != null && item.InnerXml.Contains(textToSearch))) {
                                CellsContainsValue.Add(new KeyValuePair<Cell, Row>(cell, GetRow(worksheet, rowIndex)));
                            }
                        }
                    }
                }
            }
            return CellsContainsValue;
        }
        /// <summary>
        /// Insert rows at index. It calls the necessary methods to move the existing rows and insert the new ones.
        /// </summary>
        /// <param name="rowIndex">Index of the row that will be the first of the inserted rows</param>
        /// <param name="sheetData"></param>
        /// <param name="howManyRowsToAdd">How many rows to insert</param>
        /// <param name="howManyColumns">How many columns to insert for each row</param>
        private static void InsertRows(uint rowIndex, SheetData sheetData, int howManyRowsToAdd, int howManyColumns)
        {
            if (howManyRowsToAdd != 0) {
                MoveRows(rowIndex, sheetData, howManyRowsToAdd);
                InsertNewRows(rowIndex, sheetData, howManyRowsToAdd, howManyColumns);
            }
        }
        /// <summary>
        /// Inserts new rows at index. It adds multiple rows and multiple columns for each row.
        /// </summary>
        /// <param name="rowIndex">Index of the row that will be the first of the inserted rows</param>
        /// <param name="sheetData"></param>
        /// <param name="howManyRows">How many rows to insert</param>
        /// <param name="howManyColumns">How many columns to insert for each row</param>
        private static void InsertNewRows(uint rowIndex, SheetData sheetData, int howManyRows, int howManyColumns)
        {
            Row RefRow = GetRow(sheetData.Parent as Worksheet, rowIndex);
            var styleIndex = RefRow.Elements<Cell>().First().StyleIndex;
            List<Row> RowsToAdd = new List<Row>();
            for (int i = 0; i < howManyRows; ++i) {
                int newRowIndex = (int)rowIndex + i;
                Row rowToAdd = new Row() { RowIndex = new UInt32Value((uint)newRowIndex) };
                for (uint cellColumnId = 0; cellColumnId <= howManyColumns; ++cellColumnId) {
                    Cell cellToAdd = new Cell() {
                        CellReference = new StringValue(GetExcelColumnName(cellColumnId) + newRowIndex),
                        CellFormula = new CellFormula(),
                        DataType = CellValues.String,
                        StyleIndex = styleIndex
                    };
                    rowToAdd.Append(cellToAdd);
                }
                RowsToAdd.Add(rowToAdd);
            }
            //  Insert 1st row above given row index. Insert all others below that.
            for (int i = 0; i < RowsToAdd.Count; ++i) {
                Row row = RowsToAdd[i];
                if (i == 0) {
                    sheetData.InsertBefore(row, RefRow);
                } else {
                    sheetData.InsertAfter(row, RowsToAdd[i - 1]);
                }
            }
            //  Row that had previously row index of the 1st row inserted must now update its cell references and row index.
            uint newIndexForFirstRow = RefRow.RowIndex.Value + (uint)howManyRows;
            foreach (Cell cell in RefRow.Elements<Cell>()) {
                cell.CellReference = new StringValue(cell.CellReference.Value.Replace(RefRow.RowIndex.Value.ToString(), newIndexForFirstRow.ToString()));
            }
            RefRow.RowIndex = new UInt32Value(newIndexForFirstRow);
        }
        /// <summary>
        /// Moves Rows given number of rows below. It keeps Cell References for each row.
        /// </summary>
        /// <param name="aboveRowIndex">Index of the row below of which the rows will be moved</param>
        /// <param name="sheetData"></param>
        /// <param name="howManyRowsToMove">How many rows to move below</param>
        private static void MoveRows(uint aboveRowIndex, SheetData sheetData, int howManyRowsToMove)
        {
            IEnumerable<Row> RowsToMove = sheetData.Descendants<Row>().Where(r => r.RowIndex.Value > aboveRowIndex);
            List<KeyValuePair<StringValue, int>> MergedCellsToAdd = new List<KeyValuePair<StringValue, int>>();
            foreach (Row rowToMove in RowsToMove) {
                var newRowIndex = Convert.ToUInt32(rowToMove.RowIndex.Value + howManyRowsToMove);
                foreach (Cell cellToMove in rowToMove.Elements<Cell>()) {
                    var m = GetMergedCell(sheetData.Parent as Worksheet, cellToMove.CellReference.Value);
                    if (m != null) {
                        MergedCellsToAdd.Add(new KeyValuePair<StringValue, int>(m.Reference.Value, (int)rowToMove.RowIndex.Value));
                        m.Remove();
                    }
                    cellToMove.CellReference = new StringValue(
                        cellToMove.CellReference.Value.Replace(rowToMove.RowIndex.Value.ToString(), newRowIndex.ToString()));
                }
                rowToMove.RowIndex = new UInt32Value(newRowIndex);
            }
            foreach(var mergeCell in MergedCellsToAdd) {
                string newRef = mergeCell.Key.ToString().Replace(mergeCell.Value.ToString(), (mergeCell.Value + howManyRowsToMove).ToString());
                GetMergedCells(sheetData.Parent as Worksheet).Append(new MergeCell() { Reference = new StringValue(newRef) });
            }
        }

        /// <summary>
        /// Returns MergedCell for given Cell CellReference.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellRef">Cell Reference to identify Cell.</param>
        /// <returns></returns>
        private static MergeCell GetMergedCell(Worksheet worksheet, string cellRef)
        {
            MergeCells mergedCells = GetMergedCells(worksheet);
            foreach(MergeCell mergedCell in mergedCells) {
                if (mergedCell.Reference.Value.Contains(cellRef)) {
                    return mergedCell;
                }
            }
            return null;
        }

        /// <summary>
        /// Returns Merged Cells table. If it is empty, it creates one.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private static MergeCells GetMergedCells(Worksheet worksheet)
        {
            MergeCells mergedCells;
            if (worksheet.Elements<MergeCells>().Count() > 0) {
                mergedCells = worksheet.Elements<MergeCells>().First();
            } else {
                mergedCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<CustomSheetView>().First());
                } else if (worksheet.Elements<DataConsolidate>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<DataConsolidate>().First());
                } else if (worksheet.Elements<SortState>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<SortState>().First());
                } else if (worksheet.Elements<AutoFilter>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<AutoFilter>().First());
                } else if (worksheet.Elements<Scenarios>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<Scenarios>().First());
                } else if (worksheet.Elements<ProtectedRanges>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<ProtectedRanges>().First());
                } else if (worksheet.Elements<SheetProtection>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<SheetProtection>().First());
                } else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0) {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<SheetCalculationProperties>().First());
                } else {
                    worksheet.InsertAfter(mergedCells, worksheet.Elements<SheetData>().First());
                }
            }
            return mergedCells;
        }
        /// <summary>
        /// Merges cells for 1 row given the cell reference and how many columns to be merged.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="leftCellRef"></param>
        /// <param name="howManyColumns"></param>
        private static void MergeTheseCells(Worksheet worksheet, string leftCellRef, int howManyColumns)
        {
            MergeCells mergedCells = GetMergedCells(worksheet);

            // Create the merged cell and append it to the MergeCells collection.
            MergeCell mergedCell = new MergeCell() { Reference = new StringValue(
                leftCellRef.ToString() + ":" + leftCellRef.ToString().Replace(leftCellRef.First(), GetExcelColumnName((uint)howManyColumns).First())) };
            mergedCells.Append(mergedCell);
        }

        /// <summary>
        /// Given a worksheet and a row index, return the row
        /// </summary>
        /// <param name="worksheet">Worksheet object that has row we need to get</param>
        /// <param name="rowIndex">Row Index</param>
        /// <returns></returns>
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            try {
                return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            } catch {
                return new Row();
            }
        }

        /// <summary>
        /// Returns cell name as shown in Excel spreadsheet, given its index. eg. given 0 should return "A", given 1 should return "B", etc
        /// </summary>
        /// <param name="columnIndex">Column Index</param>
        /// <returns></returns>
        private static string GetExcelColumnName(uint columnIndex)
        {
            //  Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Y, AA, AB, AC... AY, AZ, B1, B2..)
            //
            //  eg  GetExcelColumnName(0) should return "A"
            //      GetExcelColumnName(1) should return "B"
            //      GetExcelColumnName(25) should return "Z"
            //      GetExcelColumnName(26) should return "AA"
            //      GetExcelColumnName(27) should return "AB"
            //      ..etc..
            //
            if (columnIndex < 26)
                return ((char)('A' + columnIndex)).ToString();

            char firstChar = (char)('A' + (columnIndex / 26) - 1);
            char secondChar = (char)('A' + (columnIndex % 26));

            return string.Format("{0}{1}", firstChar, secondChar);
        }

        #endregion

    }
}
