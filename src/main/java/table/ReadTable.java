package table;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ReadTable {
    public static final String SAMPLE_XLSX_FILE_PATH = "./origin01.xlsx";
    public static final int SHEET_TABLE_HEADER = 0;                                                                     //表头
    private static int mJianShiYiJianIndex;
    private static Map<Integer, List<String>> mAvailableValues = new HashMap<>();

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getRow().getRowNum() == SHEET_TABLE_HEADER) {
                    for (int i = 0; i < cell.getRow().getLastCellNum(); i++) {
                        Cell miniCell = cell.getRow().getCell(i);
                        String miniCellValue = miniCell.toString();

                        System.out.print("miniCellValue:" + miniCellValue + "\t");                                      //最小单元格内容

                        if ("检视意见".equals(miniCellValue)) {
                            mJianShiYiJianIndex = i;
                        } /*else if ("Salary".equals(rowName)) {
                            mNeedIndex1 = i;
                        }*/
                    }
                } else {
                    Cell miniCell = cell.getRow().getCell(mJianShiYiJianIndex);
                    String miniCellValue = miniCell.toString();

                    System.out.print("miniCellValue:" + miniCellValue + "\t");

                    if (mAvailableValues.containsKey(mJianShiYiJianIndex)) {
                        mAvailableValues.get(mJianShiYiJianIndex).add(miniCellValue);
                    } else {
                        List<String> miniCellValues = new ArrayList<>();
                        miniCellValues.add(miniCellValue);
                        mAvailableValues.put(mJianShiYiJianIndex, miniCellValues);
                    }

                  /*  Cell cell2 = cell.getRow().getCell(mNeedIndex1);
                    String rowName2 = cell2.toString();
                    System.out.print("rowName2:" + rowName2 + "\t");*/
                }
            }
        }
        // Closing the workbook
        workbook.close();

        modifyExistingWorkbook();
    }


    // Example to modify an existing excel file
    private static void modifyExistingWorkbook() throws InvalidFormatException, IOException {
        // Obtain a workbook from the excel file
        Workbook workbook = WorkbookFactory.create(new FileInputStream(new File("model01.xlsx")));

        // Get Sheet at index 0
        Sheet sheet = workbook.getSheetAt(0);

        Iterator<Integer> iterator = mAvailableValues.keySet().iterator();
        while (iterator.hasNext()) {
            Integer key = iterator.next();
            List<String> strings = mAvailableValues.get(key);
        }
        // Get Row at index 1
        Row row = sheet.getRow(1);

        // Get the Cell at index 2 from the above row
        Cell cell = row.getCell(2);

        // Create the cell if it doesn't exist
        if (cell == null)
            cell = row.createCell(2);

        // Update the cell's value
        cell.setCellType(CellType.STRING);
        cell.setCellValue("Updated Value333");

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("model01.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }
}
