package com.otmm.custom.migration;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.artesia.common.exception.BaseTeamsException;
import com.artesia.common.utils.LogUtils;
import com.artesia.security.SecuritySession;
import com.artesia.security.UserGroup;
import com.artesia.security.services.UserGroupServices;
import com.artesia.security.session.services.AuthenticationServices;

public class ReadUserGroupsFromOTMMAndWriteToExcel
{

    private static XSSFWorkbook workbook = null;
    private static XSSFSheet userGroupsSheet;
    private static int tableRownum = 0;

    public static void findUserGroupsInOTMMAndWriteToExcel(String userName,
            String password, String teamsHome)
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
            workbook = new XSSFWorkbook();

            // Create a blank sheet
            userGroupsSheet = workbook.createSheet("UserGroups");

            List<String> headerTitle = new ArrayList<String>();
            headerTitle.addAll(Arrays.asList("UserGroup Name", "UserGroup Description"));

            XSSFRow tableRow = userGroupsSheet.createRow(tableRownum++);

            int cellnum = 0;
            for (String cellValue : headerTitle) {
                XSSFCell cell = tableRow.createCell(cellnum++);
                cell.setCellValue(cellValue);
            }

            readUserGroupsFromOTandWriteInExcel(userName, password);

            FileOutputStream out = new FileOutputStream(new File("MigrationSheet.xlsx"));
            workbook.write(out);
            workbook.close();
            out.close();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        catch (BaseTeamsException e) {
            LogUtils.logException(e);
        }

    }

    private static void readUserGroupsFromOTandWriteInExcel(String userName,
            String password) throws BaseTeamsException
    {
        SecuritySession session = null;
        try {
            session = AuthenticationServices.getInstance().login(userName, password);
            UserGroup[] listOfUserGroups = UserGroupServices.getInstance()
                    .retrieveAllUserGroups(session);
            for (UserGroup userGroups : listOfUserGroups) {
                XSSFRow tableRow = userGroupsSheet.createRow(tableRownum++);
                XSSFCell nameCell = tableRow.createCell(0);
                nameCell.setCellValue(userGroups.getName());
                XSSFCell descriptionCell = tableRow.createCell(1);
                descriptionCell.setCellValue(userGroups.getDescription());
            }
        }
        finally {
            AuthenticationServices.getInstance().logout(session);
        }
    }

}
