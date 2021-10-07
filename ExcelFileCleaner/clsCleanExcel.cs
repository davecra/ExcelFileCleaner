using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;

namespace ExcelFileCleaner
{
    class clsCleanExcel
    {
        public delegate void CleanExcelEventHandler(object sender, EventArgs e);
        public delegate void CleanExcelEventHandler2(double value);
        public event CleanExcelEventHandler LogUpdated;
        public event CleanExcelEventHandler StatusUpdated;
        public event CleanExcelEventHandler UpdateProgress;
        public event CleanExcelEventHandler2 UpdateSmallProgress;
        private const string DIVIDER = "----------------------------------------------";

        public enum OPTIONS : int
        {
            CORRECT_STYLES,
            CORRECT_CELLS,
            CORRECT_NAMED,
            CORRECT_LINK,
            CORRECT_FORMULAS,
        }

        private class DefinedNameVal
        {
            // Defined name string (I will use this string as a reference to my data) 
            // Sheet name 
            // Start column 
            // Start row 
            // End column (only present if defined name is a range) 
            // End row (only present if defined name is a range) 
            public string Key { get; set; }
            public string SheetName { get; set; }
            public string StartColumn { get; set; }
            public string StartRow { get; set; }
            public string EndColumn { get; set; }
            public string EndRow { get; set; }
        }

        private class styleCount
        {
            public styleCount(CellStyle cs, CellFormat csf, bool b)
            {
                Found = false;
                cellStyleFormat = csf;
                cellStyle = cs;
                builtIn = b;
                // if we have a built-in style we flip it to 
                // also be in use in order to protect it becase
                // we do not want to delete it
                if (b)
                    inUse = true;
            }
            public CellStyle cellStyle { get; set; }
            public CellFormat cellStyleFormat { get; set; }
            public CellFormat cellFormat { get; set; }
            public bool inUse { get; set; }
            public bool builtIn { get; set; }
            public bool Found { get; set; }
        }

        // dictionary to hold all styles found while working with the file
        //Dictionary<int, styleCount> styleUsage = new Dictionary<int, styleCount>();
        public double getStylesCount(string docPath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docPath, false))
            {
                double val = (double)document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellStyles.Count;
                document.Close();
                return val;
            }
        }

        public int getSheetCount(string docPath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docPath, false))
            {
                int cnt = 0;
                foreach(Sheet s in document.WorkbookPart.Workbook.Sheets)
                    cnt++;
                document.Close();
                return cnt;
            }
        }

        private string logData;
        public string getLogData()
        {
            return logData;
        }

        private string statusData;
        public string getStatus()
        {
            return statusData;
        }

        private bool[] fOptions;
        public bool[] setOptions
        {
            set
            {
                fOptions = value;
            }
        }

        private void UpdateStatus(string status)
        {
            statusData = status;
            StatusUpdated(this, EventArgs.Empty);
        }

        private void Log(string data)
        {
            logData = data;
            LogUpdated(this, EventArgs.Empty);
        }

        /// <summary>
        /// This procedure cleans the Excel file and reports back the status
        /// via an event method (updated).
        /// </summary>
        /// <param name="docName">Full path and filename to the workbook</param>
        public bool CleanFile(string docPath)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docPath, true))
            {
                // ** CLEAN STYLES
                if (fOptions[(int)OPTIONS.CORRECT_STYLES])
                {
                    CleanStyles(document.WorkbookPart.WorkbookStylesPart.Stylesheet);
                }

                // ** DELETE INVALID RANGES
                if (fOptions[(int)OPTIONS.CORRECT_NAMED])
                {
                    RemoveInvalidNamedRanges(document.WorkbookPart);
                }

                // ** REMOVE EXTERNAL LINKS FROM THE FILE
                if (fOptions[(int)OPTIONS.CORRECT_LINK])
                {
                    Log(DIVIDER);
                    Log("Removing External links...");
                    foreach (ExternalReference ex in document.WorkbookPart.Workbook.ExternalReferences)
                    {
                        try
                        {
                            Log("Removing link to " + ex.Id);
                            ex.Remove();
                        }
                        catch { } // ignore
                    }
                }
  
                // ** DELETE UNUSED RANGES
                foreach (Sheet s in document.WorkbookPart.Workbook.Sheets)
                {
                    string relationshipId = s.Id.Value;
                    try
                    {
                        // this will fail if we have a chart sheet
                        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                        // Get the cell at the specified column and row.
                        Log("Analyzing " + s.Name);
                        try
                        {
                            CleanCells(worksheetPart.Worksheet, s.Name);
                        }
                        catch (Exception ex)
                        {
                            Log("Failed to clean: " + ex.Message);
                            // we will continue on at this point. There could be
                            // a number of reasons this failed, but since it it
                            // likely not a compelte failure, we continue...
                        }
                        Log("Completed analysis of " + s.Name + ".");
                        Log(DIVIDER);
                        worksheetPart = null;
                    }
                    catch
                    {
                        // workbook cast fails...
                        // we will inore this and continue on.
                    }
                    // tick an update on progress
                    UpdateProgress(this, EventArgs.Empty);
                }

                // ** DELETE CALCULATIONS CHAIN
                if (fOptions[(int)OPTIONS.CORRECT_FORMULAS])
                {
                    // since deleting forulas can have a detrimental affect on the 
                    // cached calculations, so we need Excel to build a new one
                    try
                    {
                        document.WorkbookPart.DeletePart(document.WorkbookPart.CalculationChainPart);
                    }
                    catch { } // ignore it fail
                }

                try
                {
                    // Compelted - now save
                    Log("Saving the workbook...");
                    UpdateStatus("... [PLEASE WAIT] ...");
                    document.WorkbookPart.Workbook.Save();
                    //document.Close();
                    Log("Workbook clean compelted!");
                    UpdateStatus("");
                }
                catch (Exception ex)
                {
                    Log("Failed to save: " + ex.Message);
                    // failed to save
                    return false;
                }
            }

            try
            {
                GC.Collect();
                GC.Collect();
                return true; // success
            }
            catch (Exception ex)
            {
                Log("Error finalizing cleaner - file many be corrupt. " + ex.Message);
                return false; // sometimes there is an out of memory error if the file
                              // is in a bad state. So we will fail out here.
            }
        }

        /// <summary>
        /// Here we clean the styles from the workbook that are not
        /// being used. We do this by catalogging the CellStyle and
        /// matching it to the correct CellStyleFormat and then
        /// looping through the CellFormats applied in the workbook
        /// and deleting the CellStyles and CellStyleFormats that 
        /// are not represented in the CellFormats list.
        /// </summary>
        /// <param name="stylesheet"></param>
        private void CleanStyles(Stylesheet stylesheet)
        {
            List<CellFormat> cfList = new List<CellFormat>();
            // build a list of cell formats
            // NOTE: In the CellStyle namespace there is a FormatId
            //       the CellStyleFormats are actually an array
            //       of styles and the FormatId in the CellStyle
            //       references the index position of the 
            //       CellStyleFormat. So what we have to do is
            //       load all these into an List and then using
            //       the formatID of the cell style get reference
            //       to the related CellStyleFormat
            // SEE: http://www.surveyxtreme.com/?p=126
            Log("Reviewing styles.");
            foreach (CellFormat cf in stylesheet.CellStyleFormats)
                cfList.Add(cf);

            // now get a list of the named Styles
            Dictionary<int,styleCount> styleList = new Dictionary<int,styleCount>();
            foreach (CellStyle cs in stylesheet.CellStyles)
            {
                try
                {
                    // add all styles and the cellformats
                    int formatId = int.Parse(cs.FormatId);
                    styleList.Add(formatId, new styleCount(cs, cfList[formatId], !(cs.BuiltinId == null)));
                }
                catch { } // ignore
            }
            Log("There are " + styleList.Count.ToString() + " styles in the workbook.");

            // now that we have aligned our cellstyleformats with cellformats
            // we need to delete everyone of them that are not related
            // to a specific CellFormat.
            int counter = 0;
            foreach (CellFormat xf in stylesheet.CellFormats)
            {
                try
                {
                    int formatId = int.Parse(xf.FormatId);
                    if (styleList.Keys.Contains(formatId))
                    {
                        styleList[formatId].inUse = true;
                        styleList[formatId].Found = true;
                        styleList[formatId].cellFormat = xf;
                        counter++;
                    }
                }
                catch { } // ignore
            }
            Log("There are " + counter.ToString() + " styles in use.");
            Log("Cleaning unused styles:");

            // now loop though the stylelist and delete anything that
            // is not in use and then log it
            foreach (KeyValuePair<int,styleCount> c in styleList)
            {
                // only those not in use and those that are not built in
                if (c.Value.Found == false && c.Value.builtIn == false)
                {
                    Log("Cleaning style: " + c.Value.cellStyle.Name);
                    c.Value.cellStyleFormat.Remove();
                    c.Value.cellStyle.Remove();
                }
            }

            // clean the list in reverse - deleting styles that were
            // removed in the operation above
            for (int i = (styleList.Count - 1); i >= 0; i--)
            {
                int keyVal = styleList.Keys.ToList<int>()[i];
                if (styleList[keyVal].inUse == false && styleList[keyVal].builtIn == false)
                    styleList.Remove(keyVal);
            }

            Log("Updating style indexes.");
            // now correct all the indexes
            foreach (KeyValuePair<int, styleCount> c in styleList)
            {
                int idx = 0; // keep track of the index
                // now we loop though the collection of the 
                // cellstyleformats (xsf) remaining in the
                // stylesheet. If the cellstyleformat is the
                // same cellstyleformat we are referencing in
                // our collection via styleList, then we will
                // update both the cellStyle xfId and the
                // cellFormat xfId in the stylesheet to 
                // reference the new index location...
                foreach (CellFormat xsf in stylesheet.CellStyleFormats)
                {
                    // CellStyleFormat (refersTo>) CellFormat
                    // CellStyle (refersTo>) CellFormat
                    if (c.Value.cellStyleFormat == xsf)
                    {
                        if (c.Value.Found) // only if matched to a cellformat
                        {
                            // we have a match... index them
                            c.Value.cellStyle.FormatId.Value = (uint)idx;
                            c.Value.cellFormat.FormatId.Value = (uint)idx;
                        }
                        break;
                    }
                    idx++;
                }
            }

            // update counts
            stylesheet.CellStyles.Count.Value = (uint)styleList.Count;
            stylesheet.CellStyleFormats.Count.Value = (uint)styleList.Count;    
    
            // clean
            stylesheet.Save();
            stylesheet = null;
            Log(DIVIDER);
        }

        /// <summary>
        /// Removes invalid named ranges from the file that reference invalid data #REF
        /// or reference an external file on the same system or on a network share
        /// </summary>
        /// <param name="wb"></param>
        private void RemoveInvalidNamedRanges(WorkbookPart wb)
        {
            Log(DIVIDER);
            Log("Removing invalid named ranges...");
            int removeCnt = 0;
            List<DefinedName> namesToRemove = new List<DefinedName>();
            try
            {

                if (wb.Workbook.GetFirstChild<DefinedNames>() != null)
                {
                    // catalog the names to be removed
                    foreach (DefinedName name in wb.Workbook.GetFirstChild<DefinedNames>())
                    {
                        string value = name.InnerText;
                        if (value.Contains("#REF") || // missing references
                            value.Contains("#N/A") || // invalid references
                            value.Contains(":\\") ||  // part of a file
                            value.Contains("\\\\") || //parts of a file
                            (value.Contains("[") && value.Contains("]"))) // part of a file
                            namesToRemove.Add(name);
                    }
                }

                // now remove the names
                foreach(DefinedName name in namesToRemove)
                {
                    try
                    {
                        Log("Removing Named Range: " + name.Name);
                        name.Remove();
                        removeCnt++;
                    }
                    catch 
                    {
                        Log(" >> REMOVE FAILED: " + name.Name);
                    }
                }
            }
            catch { } // ignore
            Log("Removed " + removeCnt.ToString() + " invalid named ranges.");
        }

        /// <summary>
        /// Workhorse function that will open the worksheet, iterate through the
        /// ranges in the passed in worksheet and then remove unused ranges.
        /// It also reviews the styles in use throughout the active range.
        /// </summary>
        /// <param name="worksheet">Worksheet part from the XML</param>
        private void CleanCells(Worksheet worksheet, string workbookName)
        {
            // get the sheet data
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            int totalNumRows = sheetData.Elements<Row>().Count();
            if(totalNumRows == 0)
            {
                // there are no rows.
                return;
            }

            // *** CATALOG THE USED RANGE ***
            double[] maxVals = getUsedRange(sheetData);

            // now verify there are cells to clean
            if (maxVals[2] == 0)
            {
                // everything is good.
                Log("Sheet is ok.");
                // if we are not trying to remove styles and also not tryiing to
                // remove formulas with links, then we will return - compelted...
                // otherwise, we need to continue scanning the sheet
                if(!fOptions[(int)OPTIONS.CORRECT_FORMULAS] && !fOptions[(int)OPTIONS.CORRECT_STYLES])
                    return;
            }
            else
            {
                // we state we are reviewing these cells since it will
                // likely not actually require this many cells to be 
                // removed. The calculations are based on all cells in
                // the entire range being activated and this is rarely
                // the case. So we need to "review" them.
                Log("Reviewing " + maxVals[2].ToString("#,#",CultureInfo.InvariantCulture) + " cells.");
            }

            // DELETE THE UNUSED RANGE AND CATALOG STYLES
            // for very large sheet ranges with hundreds of thousands to a 
            // million rows this process can take a very long time so we will
            // update the status with information on how long it will take
            DateTime dtStart = DateTime.Now;
            IEnumerable<Row> rows = sheetData.Elements<Row>();
            List<Row> rowsToDelete = new List<Row>();
            foreach(Row r in rows)
            {
                IEnumerable<Cell> cells = r.Elements<Cell>();
                if (cells.Count() == 0)
                {
                    // A cell does not exist at the specified column, in the specified row.
                    rowsToDelete.Add(r);
                    continue;
                }

                List<Cell> cellsToDelete = new List<Cell>();
                foreach(Cell c in cells)
                {
                    // remove cells -- if the user has set this value
                    int[] rc = splitRC(c.CellReference.Value);
                    if (fOptions[(int)OPTIONS.CORRECT_CELLS])
                    {
                        if (rc[1] > maxVals[1] || rc[0] > maxVals[0])
                        {
                            cellsToDelete.Add(c);
                            continue; // done
                        }
                    }

                    // if we are here we are on a cell that will not be
                    // removed from the sheet, so we can check the formula
                    // ** CHECK FORMULA
                    if (fOptions[(int)OPTIONS.CORRECT_FORMULAS])
                    {
                        // look at the cell formula to see if it contains a link to
                        // an external file. It is does, remove the formula from the file.
                        if (c.CellFormula.InnerText.Contains(":\\") || c.CellFormula.InnerText.Contains("\\\\") ||
                            (c.CellFormula.InnerText.Contains("[") && c.CellFormula.InnerText.Contains("]")))
                        {
                            Log("Removing Cell Formula with External Link for " + workbookName + ":" + c.CellReference);
                            Log(" >> Formula: " + c.CellFormula.InnerText);
                            c.CellFormula.Text = ""; // remove
                        }
                }
                    // clean-up
                    rc = null;
                }
                // are there cells to delete? If there are then
                // we will delete them here
                if(cellsToDelete.Count > 0)
                {
                    // delete all the cells
                    foreach (Cell c in cellsToDelete)
                    {
                        try
                        {
                            c.Remove();
                        }
                        catch
                        {
                            Log("Failed to remove cell " + c.CellReference.Value);
                        }
                    }
                }

                // Now that all the cells were deleted on the row, it is
                // acually empty? If so, then we need to delete the row...
                // is the row now empty?
                if(r.Elements<Cell>().Count() == 0)
                    rowsToDelete.Add(r);
                
                // cleanup
                cells = null;
            }

            // if we remove a row, update the log with the count and
            // then re-analyze the sheet
            if (rowsToDelete.Count > 0)
            {
                int delCnt = 0;
                foreach (Row r in rowsToDelete)
                {
                    try
                    {
                        r.Remove();
                    }
                    catch
                    {
                        Log("Unable to clean row #:" + r.RowIndex.Value.ToString());
                    }

                    if (delCnt % 1000 == 0)
                    {
                        UpdateSmallProgress((totalNumRows - delCnt) / totalNumRows);
                    }
                    //UpdateStatus((totalNumRows - delCnt).ToString("#,#", CultureInfo.InvariantCulture) +
                    //                   " rows left to review.");
                    delCnt++;
                }

                Log("Cleaned " + delCnt.ToString("#,#", CultureInfo.InvariantCulture) + " cells.");
                Log("Analyzing again...");
                getUsedRange(sheetData);
            }
            else
            {
                // otherwise everything is good.
                Log("Sheet is ok.");
            }
            UpdateStatus(""); // clean
            // clean up
            sheetData = null;
            maxVals = null;
        }

        /// <summary>
        /// Takes a rows collection and iterates through the rows and columns
        /// and then catalogs the unused range
        /// </summary>
        /// <param name="rows">The rows collection from the XML part</param>
        /// <returns>Integer[2] Row/Column Array for the used range</returns>
        private double[] getUsedRange(SheetData sheetData) //IEnumerable<Row> rows)
        {
            int maxRow = 0;
            int maxCol = 0;
            int colCount = 0;
            int cnt = 0;
            IEnumerable<Row> rows = sheetData.Elements<Row>();
            foreach(Row r in rows)
            {
                IEnumerable<Cell> cells = r.Elements<Cell>();
                if (cells.Count() == 0)
                    // A cell does not exist at the specified column, in the specified row.
                    continue;

                foreach (Cell c in cells)
                {
                    cnt++;
                    if (cnt > 0 && cnt % 10000 == 0)
                        UpdateStatus("Analyzing sheet (" + cnt.ToString() + ")... This may take a while.");
                    int[] rc = splitRC(c.CellReference.Value);
                    if (colCount < rc[0])
                        colCount = rc[0];

                    if ((c.CellValue != null && c.CellValue.Text.Length > 0) || 
                        (c.CellFormula != null && c.CellFormula.Text.Length > 0))
                    {
                        // get max values...    
                        if (rc[1] > maxRow)
                            maxRow = rc[1];
                        if (rc[0] > maxCol)
                            maxCol = rc[0];
                    }
                }
            }

            Log("REPORTED - r(" + rows.Count() + ")c(" + colCount.ToString() + ")");
            Log("ACTUAL   - r(" + maxRow.ToString() + ")c(" + maxCol.ToString() + ")");
            UpdateStatus("");

            double extraCells = (rows.Count() * colCount) - (maxRow * maxCol);

            double[] retVal = new double[3];
            retVal[0] = maxCol;
            retVal[1] = maxRow;
            retVal[2] = extraCells; // total number of invalid cells

            return retVal;
        }

        /// <summary>
        /// Function used to read the string data from the XML and split out
        ///  the Row/Column information from the XML, it returns a 2D array
        ///  for the row and column
        /// </summary>
        /// <param name="rc">string representing data like A34</param>
        /// <returns>The row and column after converting the alpha value for column to a number</returns>
        private int[] splitRC(string rc)
        {
            int breakPoint = 0;

            for (int i = 1; i < rc.Length; i++)
            {
                int num = 0;
                if (int.TryParse(rc.Substring(i, 1), out num))
                {
                    breakPoint = i;
                    break;
                }
            }

            int c = (int)ConvertFromBase26(rc.Substring(0, breakPoint));
            int r = int.Parse(rc.Substring(breakPoint));

            int[] vals = new int[2];
            vals[0] = c + 1; // column
            vals[1] = r;     // row

            return vals;
        }

        /// <summary>
        /// Used to convert alpha values for columns into a single number.
        /// Such as A = 1
        /// </summary>
        /// <param name="val">Takes a string representing the column</param>
        /// <returns>Number representing the column</returns>
        private int ConvertFromBase26(string val)
        {
            const double BASE = 26.0;
            int ret = 0;

            char[] vals = val.ToUpper().ToCharArray();
            int last = vals.Length - 1;

            for (int x = 0; x < vals.Length; x++)
            {
                if (vals[x] < 'A' || vals[x] > 'Z')
                    throw new ArgumentException("Not a valid Base26 string.", val);

                ret += (int)(Math.Pow(BASE, (double)x) * (vals[last - x] - 'A'));
            }

            return ret;
        }

    }
}
