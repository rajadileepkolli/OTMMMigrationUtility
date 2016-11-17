package com.otmm.custom.migration;

import java.util.Scanner;

public class Selection {

    public static void main(String[] args) {

        Scanner selectionReadValue = new Scanner(System.in);
        System.out.println(
                "1: Export data from OTMM to an Excel Sheet \n2: Import data to OTMM  ");
        String selection = selectionReadValue.next();
        String userName = null;
        String password = null;
        String teamsHome = null;
        // Export Tables to Excelsheet
        if (selection.equalsIgnoreCase("1")) {
            Scanner userNameReadValue = new Scanner(System.in);
            System.out.println("Enter User Name:");
            userName = userNameReadValue.next();

            Scanner passwordReadValue = new Scanner(System.in);
            System.out.println("Enter Password:");
            password = passwordReadValue.next();

            if (System.getenv("TEAMS_HOME") == null) {
                Scanner teamsHomeReadValue = new Scanner(System.in);
                System.out.println("Enter TEAMS_HOME path of Import Environment:");
                teamsHome = teamsHomeReadValue.next();
                teamsHomeReadValue.close();
            }
            userNameReadValue.close();
            passwordReadValue.close();
            selectionReadValue.close();

            ReadMetadataTablesFromOTMMAndWriteToExcel
                    .findMetadataTablesAndWriteToExcel(userName, password, teamsHome);
        }
        else if (selection.equalsIgnoreCase("2")) {
            Scanner userNameReadValue = new Scanner(System.in);
            System.out.println("Enter User Name:");
            userName = userNameReadValue.next();

            Scanner passwordReadValue = new Scanner(System.in);
            System.out.println("Enter Password:");
            password = passwordReadValue.next();

            userNameReadValue.close();
            passwordReadValue.close();
            selectionReadValue.close();
            if (System.getenv("TEAMS_HOME") == null) {
                Scanner teamsHomeReadValue = new Scanner(System.in);
                System.out.println("Enter TEAMS_HOME path of Export Environment:");
                teamsHome = teamsHomeReadValue.next();
                teamsHomeReadValue.close();
            }
            ReadFromExcelAndCreateMetadataTableInOTMM.createMetaDataTablesInOTMM(userName,
                    password, teamsHome);
        }
    }

}
