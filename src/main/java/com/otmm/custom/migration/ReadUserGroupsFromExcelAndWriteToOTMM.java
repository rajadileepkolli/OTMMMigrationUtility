package com.otmm.custom.migration;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.artesia.common.exception.BaseTeamsException;
import com.artesia.common.utils.LogUtils;
import com.artesia.entity.TeamsNumberIdentifier;
import com.artesia.security.SecuritySession;
import com.artesia.security.UserGroup;
import com.artesia.security.services.UserGroupServices;
import com.artesia.security.session.services.AuthenticationServices;

public class ReadUserGroupsFromExcelAndWriteToOTMM
{

    private static XSSFSheet userGroupSheet;
    private static SecuritySession session;
    private static TeamsNumberIdentifier parentUserGroupId = new TeamsNumberIdentifier(
            new Long(1));
    private static List<String> existingUsergroups = null;

    public static void createUserGroupsInOTMM(String userName, String password,
            String teamsHome)
    {
        if (System.getenv("TEAMS_HOME") != null) {
            System.setProperty("TEAMS_HOME", System.getenv("TEAMS_HOME"));
        }
        else {
            System.setProperty("TEAMS_HOME", teamsHome);
        }

        // Create Workbook instance holding reference to .xlsx file
        try (XSSFWorkbook workbook = new XSSFWorkbook(
                new FileInputStream(new File("MigrationSheet.xlsx")))) {

            // Get desired sheet from the workbook
            userGroupSheet = workbook.getSheet("UserGroups");

            createUserGroups(userName, password);

            System.out.println("Created UserGroups in otmm");

        }
        catch (IOException e) {
            e.printStackTrace();
        }
        catch (BaseTeamsException e) {
            LogUtils.logException(e);
        }

    }

    private static void createUserGroups(String userName, String password)
            throws BaseTeamsException
    {
        try {
            session = AuthenticationServices.getInstance().login(userName, password);
            for (Row row : userGroupSheet) {
                if (row.getRowNum() != 0) {
                    String userGroupName = row.getCell(0).getStringCellValue();
                    if (!userGroupCreated(userGroupName)
                            && !Objects.equals(userGroupName, "Everyone")) {
                        UserGroup userGroup = new UserGroup();
                        userGroup.setName(userGroupName);
                        userGroup.setDescription(row.getCell(1).getStringCellValue());
                        UserGroupServices.getInstance().createUserGroup(parentUserGroupId,
                                userGroup, session);
                        System.out.println("Created UserGroup with Name" + userGroupName);
                    }
                }
            }
        }
        finally {
            AuthenticationServices.getInstance().logout(session);
        }

    }

    private static boolean userGroupCreated(String userGroupName)
            throws BaseTeamsException
    {
        if (existingUsergroups == null) {
            existingUsergroups = UserGroupServices.getInstance()
                    .retrieveAllChildUserGroups(parentUserGroupId, session).stream()
                    .map(userGroup -> userGroup.getName()).collect(Collectors.toList());
        }
        return existingUsergroups.contains(userGroupName);
    }

}
