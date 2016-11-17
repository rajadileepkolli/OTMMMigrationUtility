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
import com.artesia.common.utils.LogUtils;
import com.artesia.metadata.admin.DatabaseColumn;
import com.artesia.metadata.admin.LookupTable;
import com.artesia.metadata.admin.services.MetadataAdminServices;
import com.artesia.security.SecuritySession;
import com.artesia.security.session.services.AuthenticationServices;

/**
 * @author rajakolli
 *
 */
public class ReadMetadataTablesFromOTMMAndWriteToExcel {

    private static XSSFSheet tablesSheet;
    private static int       tableRownum  = 0;
    private static XSSFSheet lookUpSheet;
    private static int       lookUprownum = 0;

    /**
     * @param userName
     * @param password
     * @param teamsHome
     */
    public static void findMetadataTablesAndWriteToExcel(String userName, String password,
            String teamsHome) {

        // Set TEAMS_HOME value
        if (System.getenv("TEAMS_HOME") != null) {
            System.setProperty("TEAMS_HOME", System.getenv("TEAMS_HOME"));
        }
        else {
            System.setProperty("TEAMS_HOME", teamsHome);
        }

        try {

            // Blank workbook
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

            // Create a blank sheet
            tablesSheet = xssfWorkbook.createSheet("OTMM Tables");

            List<String> headerTitle = new ArrayList<>();
            headerTitle.addAll(Arrays.asList("TableName", "IsTabular", "Coloumn Name",
                    "Database Type", "MetadataType", "Size", "isNullable"));

            XSSFRow tableRow = tablesSheet.createRow(tableRownum++);

            int cellnum = 0;
            for (String cellValue : headerTitle) {
                XSSFCell cell = tableRow.createCell(cellnum++);
                cell.setCellValue(cellValue);
            }

            readMetadataTablesFromOTandWriteInExcel(userName, password);

            // Create a blank sheet
            lookUpSheet = xssfWorkbook.createSheet("OTMM LookUpTables");

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
            FileOutputStream out = new FileOutputStream(
                    new File("OTMM_ImportExport.xlsx"));
            xssfWorkbook.write(out);
            xssfWorkbook.close();
            out.close();
            System.out.print("OTMM_ImportExport.xlsx written successfully on disk.");
        }
        catch (BaseTeamsException ex) {
            LogUtils.logException(ex);
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Retrieves table from the OTMM and attempts to write those values in Excel
     * 
     * @param userName
     * @param password
     * @throws BaseTeamsException
     */
    private static void readFromOTAndWriteLookUpTables(String userName, String password)
            throws BaseTeamsException {
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
     * Retrieves table from the OTMM and attempts to write those values in Excel
     * 
     * @param userName
     * @param password
     * @throws BaseTeamsException
     */
    private static void readMetadataTablesFromOTandWriteInExcel(String userName,
            String password) throws BaseTeamsException {
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
    private static void writeInOtherWorkBook(LookupTable lookupTable, int cellnum) {
        XSSFRow xssfRow = lookUpSheet.createRow(lookUprownum++);
        XSSFCell xssfCell = xssfRow.createCell(cellnum);
        xssfCell.setCellValue(lookupTable.getTableName());
        retrieveLookUpColumns(lookupTable.getColumns(), ++cellnum);
    }

    /**
     * Write all column values in excel
     * @param columns
     * @param cellnum
     */
    private static void retrieveLookUpColumns(DatabaseColumn[] columns, int cellnum) {
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

}
