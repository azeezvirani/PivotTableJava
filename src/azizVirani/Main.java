package azizVirani;

import org.apache.commons.math3.analysis.function.Add;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;

public class Main {

    public static void main(String[] args) throws Exception {

            //Assign Input Parameters/Variables
            String filePath = "C:\\Users\\Aziz.Virani\\Documents\\Automation Anywhere Files\\Automation Anywhere\\My Docs\\Java\\Sample_Data.xlsx";
            String sheetName = "Sample";
            String pivotTableSheetName = "Pivot Table";
            String areaReference = "A1:G173";
            String outputWorkBookName = "true";
            String valueField = "COUNT";
//            String row1 = "1";
//            String row2 = "2";
//            String row3 = "3";
            String rows = " 1, 2, 3 ";
            String sumRow = "6";

            // Assign File Path to a Variable
            FileInputStream fileInputStream = new FileInputStream(filePath);

            // Create Workbook variable
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            // Assign Data Sheet name to a Variable
            XSSFSheet sheet = workbook.getSheet(sheetName);


            //Create a new Sheet for Pivot Table
            XSSFSheet pivotSheet = workbook.createSheet(pivotTableSheetName);

            //Select area for Pivot Table
            AreaReference source = new AreaReference(sheetName+"!"+areaReference, SpreadsheetVersion.EXCEL2007);

            //Set Reference for Pivot Table ==>> Where to start Pivot Table
            CellReference position = new CellReference("A1");

            // Create a pivot table on Separate Sheet
            XSSFPivotTable pivotTable = pivotSheet.createPivotTable(source, position);

            //Convert String into Int
//            int iRow1 = Integer.parseInt(row1);
//            int iRow2 = Integer.parseInt(row2);
//            int iRow3 = Integer.parseInt(row3);
            int iSumRow = Integer.parseInt(sumRow);

            /* ********************************
                //Configure the pivot table
            ********************************** */
            //Remove white spaces
            rows = rows.replaceAll("\\s", "");
            System.out.println(rows);

            //Seperate Columns number and put it into String List Array
            List<String> iRows = Arrays.asList(rows.split(","));
            System.out.println(iRows);

            int count = 0;
            while (iRows.size() > count) {

                    //Covert String into Integer
                    int iRowNumber = Integer.parseInt(iRows.get(count));

                    //Add Row labels
                    pivotTable.addRowLabel(iRowNumber);
                    System.out.println(iRowNumber);

                    // Set Tabular Format
                    pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(iRowNumber).setOutline(false);

                    count++;
            }

            //pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(1).setDefaultSubtotal(false);


//            CTPivotField fld = pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldList().get(0);
//            fld.setOutline(false);



                    //.getPivotFields().getPivotFieldArray(0).setOutline(false);

            //Configure the pivot table
            //Add Row labels
//            pivotTable.addRowLabel(iRow1);
//            pivotTable.addRowLabel(iRow3);
//            pivotTable.addRowLabel(iRow2);


            //Sum up the second column
            if (valueField == "SUM") {
                    pivotTable.addColumnLabel(DataConsolidateFunction.SUM, iSumRow);
            } else if (valueField == "AVERAGE") {
                    pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, iSumRow);
            } else if (valueField == "COUNT") {
                    pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, iSumRow);
            } else if (valueField == "MAX") {
                    pivotTable.addColumnLabel(DataConsolidateFunction.MAX, iSumRow);
            } else if (valueField == "MIN") {
                    pivotTable.addColumnLabel(DataConsolidateFunction.MIN, iSumRow);
            }

            // Set Tabular Format
//            pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(iRow1).setOutline(false);
//            pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(iRow3).setOutline(false);
//            pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(iRow2).setOutline(false);



            //CTPivotField fld = pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldList().get(1);
            //fld.setOutline(false);

//            fld = pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldList().get(1);
//            fld.setOutline(false);
//
//            fld = pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldList().get(2);
//            fld.setOutline(false);

//            fld = pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldList().get(2);
//            fld.setOutline(false);

//            pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(row1).setOutline(false);
//            pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(row2).setOutline(false);

            // Get Rid of Sub Total ==> Not working - need to work more on logic and syntax
            //pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(0).setDefaultSubtotal(false);
            //CTPivotField fld = pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldList().get(2);
            //fld.setDefaultSubtotal(false);

            //Add Report filter example
            //pivotTable.addReportFilter(3);

            if (outputWorkBookName == "true") {
                    FileOutputStream fileOutputStream = new FileOutputStream(filePath);
                    workbook.write(fileOutputStream);
                    workbook.close();
            } else {
                    FileOutputStream fileOutputStream = new FileOutputStream(filePath);
                    workbook.write(fileOutputStream);
                    workbook.close();
            }

    }

}
