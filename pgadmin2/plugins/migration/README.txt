pgAdmin II Migration Wizard
---------------------------

The pgAdmin II Migration Wizard is a Plugin that allows the user to migrate tables, indexes, 
foreign keys and data from an ODBC datasource or a Microsoft Access .mdb file. The code for 
this plugin is based on the original pgAdmin Migration Wizard which was released under the 
GNU General Public Licence (see the included LICENCE.txt file) and for this reason is not 
distributed with pgAdmin II.

Pre-requisites
--------------

1) pgAdmin
2) PostgreSQL
3) MDAC 2.6 or higher (http://www.microsoft.com/data)
4) Jet 4.0 SP3 or higher (http://www.microsoft.com/data)

Installation
------------

1) Unzip pgMigration.dll from the distribution archive into the pgAdmin II plugins folder.
   Normally this is C:\Program Files\pgAdmin2\Plugins\

2) Click Start -> Run

3) In the Open: textbox, enter the following command:

   regsvr32 "C:\Program Files\pgAdmin2\Plugins\pgMigration.dll"

   You may need to alter the path to the dll file on your system.

4) Click OK, and you should see a message box indicating success. The Migration Wizard 
   should now appear on the Plugins menu in pgAdmin when you next start it, and connect to 
   a server

NOTE: The Migration Wizard and pgAdmin II are a matched set of files - if you upgrade pgAdmin, 
      you should also upgrade the Migration Wizard.

Please email any queries or bug reports to the pgAdmin Support Mailing List 
(pgadmin-support@postgresql.org).

