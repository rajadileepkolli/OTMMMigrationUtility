# OTMMMigrationUtility

To install CustomJars to maven local run below commands from relative path
mvn install:install-file -Dfile=C:\Users\rajakolli\Documents\workspace-sts-3.8.1.RELEASE\OTMMMigrationUtility.git\lib\TEAMS-sdk.jar -DgroupId=com.artesia -DartifactId=TEAMS-sdk -Dversion=16.0.2 -Dpackaging=jar
mvn install:install-file -Dfile=C:\Users\rajakolli\Documents\workspace-sts-3.8.1.RELEASE\OTMMMigrationUtility.git\lib\TEAMS-common.jar -DgroupId=com.artesia -DartifactId=TEAMS-common -Dversion=16.0.2 -Dpackaging=jar
mvn install:install-file -Dfile=C:\Users\rajakolli\Documents\workspace-sts-3.8.1.RELEASE\OTMMMigrationUtility.git\lib\jboss-cli-client.jar -DgroupId=com.wildfly -DartifactId=jboss-cli-client -Dversion=9.0.2.Final -Dpackaging=jar
mvn install:install-file -Dfile=C:\Users\rajakolli\Documents\workspace-sts-3.8.1.RELEASE\OTMMMigrationUtility.git\lib\jboss-client.jar -DgroupId=com.wildfly -DartifactId=jboss-client -Dversion=9.0.2.Final -Dpackaging=jar