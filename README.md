# Compact and Repair MS Access DB Externally

Placing this code in an open MS Access database will allow the user to compact and repair an external, unopened MS Access database. Works well for MS Access databases that keep corrupting when selecting "Compact & Repair" from the file menu.

## Getting Started

Download the mod_CompactRepair.bas file to local machine. Open desired MS Access database, or a new blank MS Access database. Import mod_CompactRepair.bas file from VBA code window. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

MS Access version 2016 or higher 
64-bit processor

### Installing

Step 1

```
Download the mod_CompactRepair.bas file
```

Step 2

```
Either open a new blank MS Access database, or open desired database to compact and repair from
```

Step 3

```
Open the VBA code window either by the shortcut button or the ALT+F11 shortcut keys
```

Step 4

```
Right click on the Project window on the left side of the code window, and select "Import File..."
```

Step 5

```
Find and select the downloaded mod_CompactRepair.bas file to import
```

## Deployment

See Installing first.

Once mod_CompactRepair.bas is installed in selected and opened MS Access database, go to the Immediate window and type: compactRepairMSAccessDB "[filePath]"

Example: 
compactRepairMSAccessDB "C:\Temp\MSAccessDBName.accdb"

Press Enter on keyboard, and compact & repair process should run.

## Built With

* [MS Access VBA](https://docs.microsoft.com/en-us/office/vba/api/overview/access) 

## Authors

* **Jennifer Rushton** - *Initial work* - https://github.com/JNRushton/

## Acknowledgements

* Thank you PurpleBooth for your ReadMe template: https://gist.githubusercontent.com/PurpleBooth/109311bb0361f32d87a2/raw/8254b53ab8dcb18afc64287aaddd9e5b6059f880/README-Template.md
