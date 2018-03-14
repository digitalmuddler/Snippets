using Excel = Microsoft.Office.Interop.Excel;

internal static void CreateColumnIDPivot(Excel._Workbook wbIN)
        {
            Excel._Worksheet dataSheet;
            Excel._Worksheet pivotSheet;
            Excel.Range pivotData = null;
            Excel.Range pivotDestination = null;
            Excel.PivotTable pTable = null;
            string pTableName = @"PIVOT SUMMARY";
            int rowCounter = 0;

            dataSheet = wbIN.Worksheets[Globals.dataSheetID];
            pivotSheet = wbIN.Worksheets[Globals.pivotSheetID];

            rowCounter = priceSheet.UsedRange.Rows.Count;

            pivotData = priceSheet.get_Range("A1", "L" + rowCounter); // selects range of pricing sheet data

            pivotDestination = pivotSheet.get_Range("A2", Type.Missing);

            pivotSheet.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, pivotData, pivotDestination, pTableName, true, false, true, true,
                Type.Missing, Type.Missing, false, false, Excel.XlOrder.xlDownThenOver, 0, Type.Missing, Type.Missing);

            pTable = (Excel.PivotTable)pivotSheet.PivotTables(pTableName);

            pTable.Format(Excel.XlPivotFormatType.xlReport1);
            pTable.InGridDropZones = false;

            Excel.PivotField planNameField = (Excel.PivotField)pTable.PivotFields("ROW 1 FIELD NAME");
            planNameField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            planNameField.Position = 1;
            planNameField.set_Subtotals(1, false);
            planNameField.set_Subtotals(1, false);

            Excel.PivotField optionField = (Excel.PivotField)pTable.PivotFields("ROW 2 FIELD NAME");
            optionField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            optionField.Position = 2;

            Excel.PivotField typeField = (Excel.PivotField)pTable.PivotFields("COLUMN FIELD NAME");
            typeField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            typeField.Position = 1;

            pTable.AddDataField(pTable.PivotFields("CALCULATION FIELD NAME"), "Pivot Summary", Excel.XlConsolidationFunction.xlMin);
        }  // END CreatePivotTable
