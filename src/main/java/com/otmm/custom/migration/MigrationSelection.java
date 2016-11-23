package com.otmm.custom.migration;

import java.util.Scanner;

public class MigrationSelection {

	public static void main(String[] args) {

		String userName = null;
		String password = null;
		String teamsHome = null;

		Scanner userNameReadValue = new Scanner(System.in);
		System.out.println("Enter User Name:");
		userName = userNameReadValue.next();

		Scanner passwordReadValue = new Scanner(System.in);
		System.out.println("Enter Password:");
		password = passwordReadValue.next();

		if (System.getenv("TEAMS_HOME") == null) {
			Scanner teamsHomeReadValue = new Scanner(System.in);
			System.out.println("Enter TEAMS_HOME path of Environment:");
			teamsHome = teamsHomeReadValue.next();
			teamsHomeReadValue.close();
		}

		Scanner selectionReadValue = new Scanner(System.in);
		System.out.println(
				"1: BackUp Existing OTMM TABLES Structure to an Excel Sheet \n2: Create the Custom Tables in OTMM \n3: BackUp Existing UserGroups to an Excel Sheet \n4: Create the UserGroups in OTMM from an Excel Sheet \n5:BackUp exisiting Security Policies to an Excel Sheet \n6:Create the Security Policies in OTMM from an Excel Sheet  ");
		String selection = selectionReadValue.next();

		// Export Tables to Excel Sheet
		if (selection.equalsIgnoreCase("1")) {
			ReadTablesFromOTMMAndWriteToExcel.findTablesAndWriteToExcel(userName, password, teamsHome);
		} else if (selection.equalsIgnoreCase("2")) {
			ReadTablesFromExcelAndWriteToOTMM.createTablesInOTMM(userName, password, teamsHome);
		} else if (selection.equalsIgnoreCase("3")) {
			ReadUserGroupsFromOTMMAndWriteToExcel.findUserGroupsInOTMMAndWriteToExcel(userName, password, teamsHome);
		} else if (selection.equalsIgnoreCase("4")) {
			ReadUserGroupsFromExcelAndWriteToOTMM.createUserGroupsInOTMM(userName, password, teamsHome);
		} else if (selection.equalsIgnoreCase("5")) {
			ReadSecurityPoliciesFromOTMMAndWriteToExcel.findSecurityPoliciesAndWriteToExcel(userName, password, teamsHome);
		} else if (selection.equalsIgnoreCase("6")) {
			ReadSecurityPoliciesFromExcelAndWriteToOTMM.createSecurityPoliciesInOTMM(userName, password, teamsHome);
		}

		userNameReadValue.close();
		passwordReadValue.close();
		selectionReadValue.close();
	}

}
