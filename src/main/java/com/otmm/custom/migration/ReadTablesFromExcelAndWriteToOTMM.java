package com.otmm.custom.migration;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.artesia.common.exception.BaseTeamsException;
import com.artesia.metadata.admin.DatabaseColumn;
import com.artesia.metadata.admin.DatabaseTable;
import com.artesia.metadata.admin.LookupTable;
import com.artesia.metadata.admin.services.MetadataAdminServices;
import com.artesia.security.SecuritySession;
import com.artesia.security.session.services.AuthenticationServices;

public class ReadTablesFromExcelAndWriteToOTMM {

    private static XSSFSheet sheet;
    private static XSSFSheet lookUpSheet;

    /**
     * Read values from Excel and create Tables in OTMM
     * @param userName
     * @param password
     * @param teamsHome
     */
    public static void createTablesInOTMM(String userName, String password,
            String teamsHome) {
        if (System.getenv("TEAMS_HOME") != null) {
            System.setProperty("TEAMS_HOME", System.getenv("TEAMS_HOME"));
        }
        else {
            System.setProperty("TEAMS_HOME", teamsHome);
        }
        try {

            FileInputStream file = new FileInputStream(new File("CBP_Tables.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get desired sheet from the workbook
            sheet = workbook.getSheet("OTMM Tables");

            createDatabaseTable(sheet, userName, password);

            System.out.println("Created table in otmm");

            lookUpSheet = workbook.getSheet("OTMM LookUpTables");

            createLookUpTable(lookUpSheet, userName, password);
            System.out.println("Created lookUp in otmm");
            workbook.close();
            file.close();
        }
        catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * Create Lookuptables in OTMM
     * @param lookUpSheet
     * @param userName
     * @param password
     */
    private static void createLookUpTable(XSSFSheet lookUpSheet, String userName,
            String password) {
        LookupTable lookUpTable = null;
        List<LookupTable> otmmTables = new ArrayList<LookupTable>();
        List<DatabaseColumn> lstDatabaseColumn = new ArrayList<>();
        boolean isCustomTable = false;
        for (Row row : lookUpSheet) {
            if (row.getRowNum() != 0) {
                Cell cell = row.getCell(0, Row.RETURN_NULL_AND_BLANK);
                if (null != cell && cell.getCellType() != 3
                        && cell.getStringCellValue().startsWith("SAMPLE")) {
                    if (lookUpTable != null) {
                        lookUpTable.setColumns(lstDatabaseColumn
                                .toArray(new DatabaseColumn[lstDatabaseColumn.size()]));
                        otmmTables.add(lookUpTable);
                        lookUpTable = null;
                        lstDatabaseColumn.clear();
                    }
                    lookUpTable = new LookupTable();
                    lookUpTable.setTableName(cell.getStringCellValue());
                    isCustomTable = true;
                    continue;
                }
                else if (null != cell && cell.getCellType() != 3
                        && !cell.getStringCellValue().startsWith("SAMPLE")) {
                    isCustomTable = false;
                }
                else if (isCustomTable) {
                    DatabaseColumn column = new DatabaseColumn();
                    column.setColumnName(row.getCell(1).getStringCellValue());
                    column.setDatabaseDatatype(row.getCell(2).getStringCellValue());
                    column.setMetadataDatatype(row.getCell(3).getStringCellValue());
                    if (row.getCell(2).getStringCellValue().equals("VARCHAR2")) {
                        column.setSize((int) row.getCell(4).getNumericCellValue());
                    }
                    column.setNullable(row.getCell(5).getBooleanCellValue());
                    column.setPrimaryKey(row.getCell(6).getBooleanCellValue());
                    lstDatabaseColumn.add(column);
                }
            }
        }
        if (lookUpTable != null) {
            lookUpTable.setColumns(lstDatabaseColumn
                    .toArray(new DatabaseColumn[lstDatabaseColumn.size()]));
            otmmTables.add(lookUpTable);
        }
        SecuritySession session = null;
        try {
            session = AuthenticationServices.getInstance().login(userName, password);
            for (DatabaseTable customTable : otmmTables) {
                System.out.println("Attempting to create lookupdomain table "
                        + customTable.getTableName());
                MetadataAdminServices.getInstance().createLookupDatabaseTable(customTable,
                        session);
            }
        }
        catch (BaseTeamsException ex) {
            System.out.println(ex.getDebugMessage());
            ex.printStackTrace();
        }
        finally {
            try {
                AuthenticationServices.getInstance().logout(session);
            }
            catch (BaseTeamsException e) {
                e.printStackTrace();
            }
        }

    }

    /**
     * Creates metadata tables in OTMM after reading from ExcelSheet
     * @param sheet
     * @param userName
     * @param password
     */
    private static void createDatabaseTable(XSSFSheet sheet, String userName,
            String password) {
        DatabaseTable metadataTable = null;
        List<DatabaseTable> otmmTables = new ArrayList<DatabaseTable>();
        List<DatabaseColumn> lstDatabaseColumn = new ArrayList<>();
        boolean isCustomTable = false;
        for (Row row : sheet) {
            if (row.getRowNum() != 0) {
                Cell cell = row.getCell(0, Row.RETURN_NULL_AND_BLANK);
                if (null != cell && cell.getCellType() != 3
                        && cell.getStringCellValue().startsWith("SAMPLE")) {
                    if (metadataTable != null) {
                        metadataTable.setColumns(lstDatabaseColumn
                                .toArray(new DatabaseColumn[lstDatabaseColumn.size()]));
                        otmmTables.add(metadataTable);
                        metadataTable = null;
                        lstDatabaseColumn.clear();
                    }
                    metadataTable = new DatabaseTable();
                    metadataTable.setTableName(cell.getStringCellValue());
                    metadataTable.setTabular(row.getCell(1).getBooleanCellValue());
                    isCustomTable = true;
                    continue;
                }
                else if (null != cell && cell.getCellType() != 3
                        && !cell.getStringCellValue().startsWith("SAMPLE")) {
                    isCustomTable = false;
                }
                else if (isCustomTable) {
                    DatabaseColumn column = new DatabaseColumn();
                    column.setColumnName(row.getCell(2).getStringCellValue());
                    column.setDatabaseDatatype(row.getCell(3).getStringCellValue());
                    column.setMetadataDatatype(row.getCell(4).getStringCellValue());
                    if (row.getCell(3).getStringCellValue().equals("VARCHAR2")) {
                        column.setSize((int) row.getCell(5).getNumericCellValue());
                    }
                    column.setNullable(row.getCell(6).getBooleanCellValue());
                    lstDatabaseColumn.add(column);
                }
            }
        }
        if (metadataTable != null) {
            metadataTable.setColumns(lstDatabaseColumn
                    .toArray(new DatabaseColumn[lstDatabaseColumn.size()]));
            otmmTables.add(metadataTable);
        }
        SecuritySession session = null;
        try {
            session = AuthenticationServices.getInstance().login(userName, password);
            System.out.println("Attempting to Save Tables");
            for (DatabaseTable customTable : otmmTables) {
                System.out.println(
                        "Attempting to create table " + customTable.getTableName());
                MetadataAdminServices.getInstance()
                        .createMetadataDatabaseTable(customTable, session);
            }
        }
        catch (BaseTeamsException ex) {
            System.out.println(ex.getDebugMessage());
            ex.printStackTrace();
        }
        finally {
            try {
                AuthenticationServices.getInstance().logout(session);
            }
            catch (BaseTeamsException e) {
                e.printStackTrace();
            }
        }
    }

}
