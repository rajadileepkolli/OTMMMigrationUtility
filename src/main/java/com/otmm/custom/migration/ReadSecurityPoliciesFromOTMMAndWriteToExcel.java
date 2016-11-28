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
import com.artesia.security.SecurityPolicy;
import com.artesia.security.SecuritySession;
import com.artesia.security.services.SecurityPolicyServices;
import com.artesia.security.session.services.AuthenticationServices;

public class ReadSecurityPoliciesFromOTMMAndWriteToExcel
{

    private static XSSFWorkbook workbook;
    private static XSSFSheet securityPoliciesSheet;
    private static int tableRownum = 0;

    public static void findSecurityPoliciesAndWriteToExcel(String userName,
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
            securityPoliciesSheet = workbook.createSheet("SecurityPolicies");

            List<String> headerTitle = new ArrayList<String>();
            headerTitle.addAll(
                    Arrays.asList("SecurityPolicy Name", "SecurityPolicy Description"));

            XSSFRow tableRow = securityPoliciesSheet.createRow(tableRownum++);

            int cellnum = 0;
            for (String cellValue : headerTitle) {
                XSSFCell cell = tableRow.createCell(cellnum++);
                cell.setCellValue(cellValue);
            }

            readSecurityPolicyFromOTandWriteInExcel(userName, password);

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

    private static void readSecurityPolicyFromOTandWriteInExcel(String userName,
            String password) throws BaseTeamsException
    {
        SecuritySession session = null;
        try {
            session = AuthenticationServices.getInstance().login(userName, password);
            SecurityPolicy[] listOfSecurityPolicy = SecurityPolicyServices.getInstance()
                    .retrieveAllSecurityPolicies(session);
            for (SecurityPolicy securityPolicies : listOfSecurityPolicy) {
                XSSFRow tableRow = securityPoliciesSheet.createRow(tableRownum++);
                XSSFCell nameCell = tableRow.createCell(0);
                nameCell.setCellValue(securityPolicies.getName());
                XSSFCell descriptionCell = tableRow.createCell(1);
                descriptionCell.setCellValue(securityPolicies.getDescription());
            }
        }
        finally {
            AuthenticationServices.getInstance().logout(session);
        }
    }

}
