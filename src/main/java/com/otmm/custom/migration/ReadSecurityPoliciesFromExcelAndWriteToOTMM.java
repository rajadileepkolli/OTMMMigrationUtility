package com.otmm.custom.migration;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.artesia.common.exception.BaseTeamsException;
import com.artesia.entity.TeamsNumberIdentifier;
import com.artesia.security.SecurityPolicy;
import com.artesia.security.SecuritySession;
import com.artesia.security.UserGroup;
import com.artesia.security.services.SecurityPolicyServices;
import com.artesia.security.services.UserGroupServices;
import com.artesia.security.session.services.AuthenticationServices;
import com.artesia.user.TeamsUser;

public class ReadSecurityPoliciesFromExcelAndWriteToOTMM {

	private static XSSFSheet securityPolicySheet;
	private static SecuritySession session;

	public static void createSecurityPoliciesInOTMM(String userName, String password, String teamsHome) {

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
			securityPolicySheet = workbook.getSheet("SecurityPolicies");

			createSecurityPolicies(securityPolicySheet, userName, password);

			System.out.println("Created SecurityPolicies in otmm");

			workbook.close();
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static void createSecurityPolicies(XSSFSheet securityPolicySheet, String userName, String password)
			throws BaseTeamsException {
		session = AuthenticationServices.getInstance().login(userName, password);
		for (Row row : securityPolicySheet) {

			int tableNum = 1;
			row.setRowNum(tableNum++);
			String policyName = row.getCell(0).getStringCellValue();
			String policyDesc = row.getCell(1).getStringCellValue();
			SecurityPolicy[] listOfSecurityPolicy = SecurityPolicyServices.getInstance()
					.retrieveAllSecurityPolicies(session);
			for (SecurityPolicy securityPolicies : listOfSecurityPolicy) {
				if (securityPolicies.getName().equals(policyName)) {
					System.out.println("Duplicate SecurityPolicy");
				} else {
				    SecurityPolicy securityPolicy = new SecurityPolicy();
				    securityPolicy.setName(policyName);
				    securityPolicy.setDescription(policyDesc);
				    
					TeamsNumberIdentifier securityPolicyId = SecurityPolicyServices.getInstance().createSecurityPolicy(securityPolicy,session);

				}
			}

		}

	}

}
