package com.otmm.custom.migration;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.artesia.common.exception.BaseTeamsException;
import com.artesia.metadata.admin.DatabaseColumn;
import com.artesia.metadata.admin.DatabaseTable;
import com.artesia.metadata.admin.LookupTable;
import com.artesia.metadata.admin.services.MetadataAdminServices;
import com.artesia.security.SecuritySession;
import com.artesia.security.session.services.AuthenticationServices;

public class ReadTablesFromOTMMAndWriteToExcel
{

    private static XSSFSheet tablesSheet;
    private static int tableRownum = 0;
    private static XSSFSheet lookUpSheet;
    private static int lookUprownum = 0;

    public static void findTablesAndWriteToExcel(String userName, String password,
            String teamsHome)
    {

        // Set TEAMS_HOME value
        if (System.getenv("TEAMS_HOME") != null) {
            System.setProperty("TEAMS_HOME", System.getenv("TEAMS_HOME"));
        }
        else {
            System.setProperty("TEAMS_HOME", teamsHome);
        }

        try {

            // Blank workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            // Create a blank sheet
            tablesSheet = workbook.createSheet("OTMM Tables");

            List<String> headerTitle = new ArrayList<String>();
            headerTitle.addAll(Arrays.asList("TableName", "IsTabular", "Coloumn Name",
                    "Database Type", "MetadataType", "Size", "isNullable"));

            XSSFRow tableRow = tablesSheet.createRow(tableRownum++);

            int cellnum = 0;
            for (String cellValue : headerTitle) {
                XSSFCell cell = tableRow.createCell(cellnum++);
                cell.setCellValue(cellValue);
            }

            readFromOTandWriteInExcel(userName, password);

            // Create a blank sheet
            lookUpSheet = workbook.createSheet("OTMM LookUpTables");

            List<String> headerTitleLookUp = new ArrayList<String>();
            headerTitleLookUp.addAll(Arrays.asList("TableName", "Coloumn Name",
                    "Database Type", "MetadataType", "Size", "isNullable", "isPrimary"));
            XSSFRow lookupRow = lookUpSheet.createRow(lookUprownum++);

            int lookupCell = 0;
            for (String cellValue : headerTitleLookUp) {
                XSSFCell cell = lookupRow.createCell(lookupCell++);
                cell.setCellValue(cellValue);
            }

            readFromOTAndWriteLookUpTables(userName, password);

            // Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("CBP_Tables.xlsx"));
            workbook.write(out);
            workbook.close();
            out.close();
            System.out.print("CBP_Tables.xlsx written successfully on disk.");
        }
        catch (BaseTeamsException e1) {
            e1.printStackTrace();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Retrieves table from the OTMM and attempts to write those values in Excel
    private static void readFromOTAndWriteLookUpTables(String userName, String password)
            throws BaseTeamsException
    {
        SecuritySession session = null;
        try {
            session = AuthenticationServices.getInstance().login(userName, password);
            List<LookupTable> otmmLookUpDomainTables = MetadataAdminServices.getInstance()
                    .retrieveEditableLookupDomainTables(session);
            if (!otmmLookUpDomainTables.isEmpty()) {
                for (LookupTable lookupTable : otmmLookUpDomainTables) {
                    writeInOtherWorkBook(lookupTable, 0);
                }
            }
        }
        finally {
            AuthenticationServices.getInstance().logout(session);
        }

    }

    /**
     * Writes value to Excel Sheet
     * @param lookupTable
     * @param cellnum
     */
    private static void writeInOtherWorkBook(LookupTable lookupTable, int cellnum)
    {
        XSSFRow row = lookUpSheet.createRow(lookUprownum++);
        XSSFCell cell = row.createCell(cellnum);
        cell.setCellValue(lookupTable.getTableName());
        retrieveLookUpColumns(lookupTable.getColumns(), ++cellnum);
    }

    /**
     * Write all column values in excel
     * @param columns
     * @param cellnum
     */
    private static void retrieveLookUpColumns(DatabaseColumn[] columns, int cellnum)
    {
        int tempCell = cellnum;
        for (DatabaseColumn tableColums : columns) {
            XSSFRow row = lookUpSheet.createRow(lookUprownum++);
            XSSFCell coloumName = row.createCell(cellnum++);
            coloumName.setCellValue(tableColums.getColumnName());
            XSSFCell databaseDBType = row.createCell(cellnum++);
            databaseDBType.setCellValue(tableColums.getDatabaseDatatype());
            XSSFCell metadataType = row.createCell(cellnum++);
            metadataType.setCellValue(tableColums.getMetadataDatatype());
            XSSFCell sizeInOTMM = row.createCell(cellnum++);
            sizeInOTMM.setCellValue(tableColums.getSize());
            XSSFCell isNullable = row.createCell(cellnum++);
            isNullable.setCellValue(tableColums.isNullable());
            XSSFCell isPrimaryKey = row.createCell(cellnum++);
            isPrimaryKey.setCellValue(tableColums.isPrimaryKey());
            cellnum = tempCell;
        }

    }

    /**
     * Read editable metadata tables and write in Excel
     * @param userName
     * @param password
     * @throws BaseTeamsException
     */
    private static void readFromOTandWriteInExcel(String userName, String password)
            throws BaseTeamsException
    {
        SecuritySession session = null;
        try {
            session = AuthenticationServices.getInstance().login(userName, password);
            List<DatabaseTable> otmmTables = MetadataAdminServices.getInstance()
                    .retrieveEditableMetadataTables(session);
            for (DatabaseTable databaseTable : otmmTables) {
                writeInExcel(databaseTable, 0);
            }
        }
        finally {
            AuthenticationServices.getInstance().logout(session);
        }
    }

    /**
     * Persist values in Excel
     * @param databaseTable
     * @param cellnum
     */
    private static void writeInExcel(DatabaseTable databaseTable, int cellnum)
    {
        XSSFRow row = tablesSheet.createRow(tableRownum++);
        XSSFCell cell = row.createCell(cellnum);
        cell.setCellValue(databaseTable.getTableName());
        XSSFCell isTabular = row.createCell(++cellnum);
        isTabular.setCellValue(databaseTable.isTabular());
        retrieveColumns(databaseTable.getColumns(), ++cellnum);
    }

    /**
     * persist column values in excel
     * @param columns
     * @param cellnum
     */
    private static void retrieveColumns(DatabaseColumn[] columns, int cellnum)
    {
        int tempCell = cellnum;
        for (DatabaseColumn tableColums : columns) {
            XSSFRow row = tablesSheet.createRow(tableRownum++);
            XSSFCell coloumName = row.createCell(cellnum++);
            coloumName.setCellValue(tableColums.getColumnName());
            XSSFCell databaseDBType = row.createCell(cellnum++);
            databaseDBType.setCellValue(tableColums.getDatabaseDatatype());
            XSSFCell metadataType = row.createCell(cellnum++);
            metadataType.setCellValue(tableColums.getMetadataDatatype());
            XSSFCell sizeInOTMM = row.createCell(cellnum++);
            sizeInOTMM.setCellValue(tableColums.getSize());
            XSSFCell isNullable = row.createCell(cellnum++);
            isNullable.setCellValue(tableColums.isNullable());
            cellnum = tempCell;
        }
    }

}
