# Excel SQL Exporter

This tool imports Excel files into SQL Tables and also optionally first downloads them from an FTP site for integrations with cloud systems and provide an easier solution compared with SSIS which will often not work well with large text fields

## Purpose

The tool was created as a replacement for Microsoft SQL Integration Services (SSIS) which can work well with smaller files but these days has a lot of limitations which this tool overcomes:
- Excel columns that contain a large number of characters can be exported without any errors or changes being made to settings
- All rows are evaluated when setting column sizes to avoid errors you get with SSIS when the first rows contain less data than subseqent rows and the column size is set to the maximum size needed (for the importer)
- Data types are detected automatically so will export correctly without code page errors, truncated values, missing values where a column mixes text and numbers
- The tool is simpler to use as just requires .NET 9 runtime to be installed and does not require Excel binaries or data access components or any other special settings

## Setting Up

Download the latest release from: https://github.com/robinwilson16/ExcelSQLImporter/releases/latest

If you have an Intel/AMD machine (most likely) then pick the `amd64` version but if you have a an ARM Snapdragon device then pick the 'arm64` version.

Downlaod and extract the zip file to a folder of your choice.

Now edit the appsettings.json to fill in details for:
| Item | Purpose |
| ---- | ---- |
| Excel File | Where you are getting the data from |
| Database Connection | Where you are saving the data to |
| Database Table | Allows you to specify a prefix for the SQL table that is created and name it from either the Excel Sheet or the File Name |
| FTP Connection (Optional) | Where you are first downloading the Excel file from if the source is remote rather than local |

Once all settings are entered then just click on `ExcelSQLImporter.exe` to run the program.
If you notice any errors appearing in the window then review these, change the settings file and try again.

## Importing Multiple Files

By default configuration values are picked up from `appsettings.json` but in case you want to use the tool to import multiple Database Tables then when running from the commandline specify the name of the config file after the .exe so to pick up settings from a config file called `FinanceImport.json` execute:

```
ExcelSQLImporter.exe FinanceImport.json
```

## Setting Up a Schedule

You can just click on may wish to set up a schedule to import one or more Excel files each night and the best way to do this in Windows is to use Task Scheduler which is available in all modern versions of Windows.

Create a new task and name it based on the Excel file it will import so for example:
```
ExcelSQLImporter - Finance Data
```

Pick a user account to run the task against. If you used Windows Authentication in your settings file then you will need to pick a user account with sufficient permissions to create the database table you are importing as well as read the Excel file if it is on a network drive.

On the Triggers tab select a schedule such as each day at 18:00.

On the Actions tab specify the location of the Excel SQL Import tool under Program/script (you can use browse to pick it). It should show as something similar to:
```
D:\ExcelSQLImporter\ExcelSQLImporter.exe
```

Optionally if you are importing more than one file then enter the name of this into the arguments box - e.g.:
```
UsersTable.json
```
