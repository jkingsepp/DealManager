Attribute VB_Name = "aNotes"
' DEVELOPMENT NOTES

' This module contains information about the procedures run as part of the model

' The first primary activity of this workbook is to format the current report, which
' involves moving data around from various worksheets to the Report tab.  Examples
' of data movement are to sort/filter the Government assets and to sort/filter the
' non-Rated assets and the Rated assets.

' The second primary activity is to generate the "final" report.  This action is performed
' by the user after review and approval of the template version.  The final version creates a
' "values-only" version of the "Report" worksheet.  It also asks the user whether updated to the
' database should be made.  It is necesary to run this data update at least once, so that
' required values are taken from the Servicer Report and Stored in the db.  These mappings
' are found on the "KDI-CI" worksheet.

' The second activity is fairly sophisticated.  It usesthe following amacros:
' 1. "ExportReport" - this prepares the report, creates a string for the file name and calls the
    ' the next procedure
' 2. "ExportDatatoDM" - this creates the database connection, using the .ini file located
'   in the users settings folder.  This location can be different depending upon operating
'   system.  XP and Windows 7 have created different locations. IF there is a problem
'   in running this procedure, users should look in this procedure for the file location
'   of dealmanager.ini.  do this by searching for  the variable DMIniFileName. the current location is
'   "C:\Documents and Settings\" & UserNameWindows & "\Application Data\NorthBound Solutions\DealManager Suite\dealmanager.ini"

'Changes/modifications

' 20161218 - moving code repository to Github, will manage changes to VBA modules, including forms, 
' through Github instead of individual workbooks

'    20140629 - changes to settings tab and format macro to reflect two assets sales per period

'   20120131 - updated the report to reflect the new PNC transaction.  Updates included:
'       1. Added in the "Final Report", "Schedule 1", "Schedule 2" and "Hedge Schedule" worksheets
'       2. Updated the "Final Report" worksheet with all links in the old "Report" tab; maintained
'           formulas and calculations provided by PNC on the "Final Report" worksheet
'       3. Updated the

'   20111212 - removed "datevalue" formatting from 'WSLINKS' macro that used the term as part
'       of the formula to get the Effective Date.  the "Datevalue' was causing error on
'       Bill Malloy's machine.

'   20111208 - changed ExportDatatoDM to reflect new database structure.  Also made changes
'       to ensure that updates to existing values are handled correctly.  Instead of simply
'       inserting records into the DealMetricValues talbe, we must first check to see
'       if a value exists for thegiven DealMetric, Deal and Effectidt.  If only a null value,
'       then we update the value with new information from KDI-CI page.  If a non-null value
'       has already been entered, we insert the entire row into DealMetricValueHistory table
'       and then insert a row with current value from kdi-CI into the DMV table.  We must
'       also check to make sure that the non-null value is different than the new values.
'       if the values are identical, then leave them alone and do not update.
'       - need to insert into historytable and to get the existing values to update
