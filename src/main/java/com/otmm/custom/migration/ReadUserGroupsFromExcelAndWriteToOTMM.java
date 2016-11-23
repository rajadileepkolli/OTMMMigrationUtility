package com.otmm.custom.migration;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.artesia.common.exception.BaseTeamsException;
import com.artesia.entity.TeamsNumberIdentifier;
import com.artesia.security.SecuritySession;
import com.artesia.security.UserGroup;
import com.artesia.security.services.UserGroupServices;
import com.artesia.security.session.services.AuthenticationServices;

public class ReadUserGroupsFromExcelAndWriteToOTMM {
	private static XSSFSheet userGroupSheet;
	private static SecuritySession session;

	public static void createUserGroupsInOTMM(String userName, String password, String teamsHome) {
		if (System.getenv("TEAMS_HOME") != null) {
			System.setProperty("TEAMS_HOME", System.getenv("TEAMS_HOME"));
		} else {
			System.setProperty("TEAMS_HOME", teamsHome);
		}
		try {

			FileInputStream file = new FileInputStream(new File("MigrationSheet.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get desired sheet from the workbook
			userGroupSheet = workbook.getSheet("UserGroups");

			createUserGroups(userGroupSheet, userName, password);

			System.out.println("Created UserGroups in otmm");

			workbook.close();
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static void createUserGroups(XSSFSheet userGroupSheet, String userName, String password)
			throws BaseTeamsException {
		session = AuthenticationServices.getInstance().login(userName, password);
		for (Row row : userGroupSheet) {
			TeamsNumberIdentifier parentGroupId = new TeamsNumberIdentifier(new Long(1));
			UserGroup userGroup = new UserGroup();
			int tableNum = 1;
			row.setRowNum(tableNum++);
			String userGroupName = row.getCell(0).getStringCellValue();
			String userGroupDescription = row.getCell(1).getStringCellValue();
			UserGroup[] listOfUserGroups = UserGroupServices.getInstance().retrieveAllUserGroups(session);
			for (UserGroup userGroups : listOfUserGroups) {
				if (userGroups.getName().equals(userGroupName)) {
					System.out.println("Duplicate User Group");
				} else {
					userGroup.setName(userGroupName);
					userGroup.setDescription(userGroupDescription);
					TeamsNumberIdentifier userGroupId = UserGroupServices.getInstance().createUserGroup(parentGroupId,
							userGroup, session);

				}
			}

		}

	}

}
