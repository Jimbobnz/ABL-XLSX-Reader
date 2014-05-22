ABL-XLSX-Reader
===============

XLSX Reader for OpenEdge ABL

~~~
{OfficeOpenXML/OfficeOpenXMLCommonShared.i}

RUN OfficeOpenXML/OfficeOpenXMLSuper.p PERSISTENT SET hnSuperProc.

    IF VALID-HANDLE(hnSuperProc) THEN
    DO:
        RUN PackageUnpack IN hnSuperProc (INPUT pchSourceDirectory, 
                                          INPUT pchFilename).
        
        LOG-MANAGER:WRITE-MESSAGE( 'Excel File: ' + pchFilename, 'DailyTrack' ). 

        /** Fetch all the workbook names and all the worksheet data.**/
        RUN GetData IN hnSuperProc (OUTPUT TABLE ttSheet,
                                    OUTPUT TABLE ttSheetData).
        
        LOG-MANAGER:WRITE-MESSAGE( 'Worksheet Found: '      + STRING(TEMP-TABLE ttSheet:HAS-RECORDS), 'DailyTrack' ).  
        LOG-MANAGER:WRITE-MESSAGE( 'Worksheet Data Found: ' + STRING(TEMP-TABLE ttSheetData:HAS-RECORDS) ).  

        RUN PackageTidy IN hnSuperProc.

        LOG-MANAGER:WRITE-MESSAGE( 'Import :' + pchFilename ).  
        
        RUN ParseWorkSheets IN THIS-PROCEDURE (INPUT pchFilename).
        
        DELETE OBJECT hnSuperProc.
    END.
    
~~~
