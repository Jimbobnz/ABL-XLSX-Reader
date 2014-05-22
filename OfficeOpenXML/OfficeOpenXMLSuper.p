&ANALYZE-SUSPEND _VERSION-NUMBER AB_v10r12
&ANALYZE-RESUME
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS Procedure 
/*------------------------------------------------------------------------
    File        : 
    Purpose     :

    Syntax      :

    Description :

    Author(s)   :
    Created     :
    Notes       :
  ----------------------------------------------------------------------*/
/*          This .W file was created with the Progress AppBuilder.      */
/*----------------------------------------------------------------------*/

/* ***************************  Definitions  ************************** */

DEFINE VARIABLE chExportDir  AS CHARACTER   NO-UNDO.
DEFINE VARIABLE chSAXHandlerProc AS CHARACTER   NO-UNDO.

DEFINE STREAM sHTML.

{src/OfficeOpenXML/OfficeOpenXMLCommonShared.i NEW SHARED}
                             
DEFINE VARIABLE chrelationtarget    AS CHARACTER   NO-UNDO.
DEFINE VARIABLE chFileName          AS CHARACTER   NO-UNDO.
DEFINE VARIABLE chPathName          AS CHARACTER   NO-UNDO.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Procedure
&Scoped-define DB-AWARE no



/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME


/* ************************  Function Prototypes ********************** */

&IF DEFINED(EXCLUDE-ConvertToDate) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD ConvertToDate Procedure 
FUNCTION ConvertToDate RETURNS DATE
  ( INPUT pinDate AS INTEGER )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-GetRelalationshipTargetByType) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD GetRelalationshipTargetByType Procedure 
FUNCTION GetRelalationshipTargetByType RETURNS CHARACTER
  ( INPUT pcType AS CHARACTER )  FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-NormaliseNotation) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION-FORWARD NormaliseNotation Procedure 
FUNCTION NormaliseNotation RETURNS CHARACTER
    (INPUT pchNumber AS CHARACTER) FORWARD.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF


/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Procedure
   Allow: 
   Frames: 0
   Add Fields to: Neither
   Other Settings: CODE-ONLY COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
/* DESIGN Window definition (used by the UIB) 
  CREATE WINDOW Procedure ASSIGN
         HEIGHT             = 22.76
         WIDTH              = 52.8.
/* END WINDOW DEFINITION */
                                                                        */
&ANALYZE-RESUME

 


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK Procedure 


/* ***************************  Main Block  *************************** */


/* RUN PackageUnpack. */

MAIN-BLOCK:
DO ON ERROR   UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK
   ON END-KEY UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK:
    
/*     IF NOT THIS-PROCEDURE:PERSISTENT THEN                                                                                                                                    */
/*     DO:                                                                                                                                                                      */
/*                                                                                                                                                                              */
/* /*         RUN ParseContentTypeXML.                                                                                                                */                        */
/* /*         RUN ParseRelationsXML(chExportDir + '_rels/.rels').                                                                                     */                        */
/* /*                                                                                                                                                 */                        */
/* /*         chRelationTarget = GetRelalationshipTargetByType('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'). */                        */
/*                                                                                                                                                                              */
/*         /* Fetch the Work Sheets */                                                                                                                                          */
/* /*         RUN ParseTargetPart (chExportDir + chRelationTarget). */                                                                                                          */
/*                                                                                                                                                                              */
/*         RUN SplitFilename(INPUT chRelationTarget,                                                                                                                            */
/*                           OUTPUT chPathName,                                                                                                                                 */
/*                           OUTPUT chFileName).                                                                                                                                */
/*                                                                                                                                                                              */
/* /*         MESSAGE chPathName chFileName. */                                                                                                                                 */
/*         ASSIGN                                                                                                                                                               */
/*             chRelationTarget = SUBSTITUTE('&1_rels/&2.rels',                                                                                                                 */
/*                                           chPathName,                                                                                                                        */
/*                                           chFileName).                                                                                                                       */
/*                                                                                                                                                                              */
/*         /*  contains paths to pivot table cache fragments and their IDs */                                                                                                   */
/*         RUN ParseRelationsXML(chExportDir + chRelationTarget).                                                                                                               */
/*                                                                                                                                                                              */
/*         /* Fetch the shared strings into Memory */                                                                                                                           */
/*         RUN ParseTargetPart (chExportDir + chPathName + GetRelalationshipTargetByType('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings')). */
/*                                                                                                                                                                              */
/*         FOR EACH ttRelationShip                                                                                                                                              */
/*             WHERE ttRelationShip.TYPE EQ 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet':                                                    */
/*                                                                                                                                                                              */
/*             RUN ParseTargetPart (chExportDir + chPathName + ttRelationShip.Target).                                                                                          */
/*                                                                                                                                                                              */
/*             DATASET dsSheetData:WRITE-XML('file', ttRelationShip.Id + '.xml' ,TRUE).                                                                                         */
/*                                                                                                                                                                              */
/*             RUN TempTable2HTML (INPUT ttRelationShip.Id + '.html').                                                                                                          */
/*                                                                                                                                                                              */
/*         END.                                                                                                                                                                 */
/*     END.                                                                                                                                                                     */
END.
    

/* RUN PackageTidy. */

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&IF DEFINED(EXCLUDE-GetData) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE GetData Procedure 
PROCEDURE GetData :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    DEFINE OUTPUT PARAMETER TABLE FOR ttSheet.
    DEFINE OUTPUT PARAMETER TABLE FOR ttSheetData.

    DEFINE VARIABLE inNumRows   AS INTEGER     NO-UNDO.
    DEFINE VARIABLE inNumCols   AS INTEGER     NO-UNDO.
    DEFINE VARIABLE INCOLCOUNT  AS INTEGER     NO-UNDO.
           
    /**  Parse the main content type file [content_type].xml file **/
    RUN ParseContentTypeXML.                                                                                                                
    RUN ParseRelationsXML(chExportDir + '_rels/.rels').                                                                                     

    ASSIGN
        chRelationTarget = GetRelalationshipTargetByType('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'). 
        
    RUN ParseTargetPart (chExportDir + chRelationTarget). 
    RUN SplitFilename(INPUT chRelationTarget,
                      OUTPUT chPathName,
                      OUTPUT chFileName).


/*         MESSAGE chPathName chFileName. */
    ASSIGN
        chRelationTarget = SUBSTITUTE('&1_rels/&2.rels',
                                      chPathName,
                                      chFileName).
    
    /*  contains paths to pivot table cache fragments and their IDs */
    RUN ParseRelationsXML(chExportDir + chRelationTarget).
    
    /* Fetch the shared strings into Memory */
    RUN ParseTargetPart (chExportDir + chPathName + GetRelalationshipTargetByType('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings')).
    
    MESSAGE chExportDir + chPathName.

    FOR EACH ttRelationShip
        WHERE ttRelationShip.TYPE EQ 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet':
        
/*         MESSAGE ttRelationShip.Id ttRelationShip.TARGET. */

        RUN ParseTargetPart (chExportDir + chPathName + ttRelationShip.Target).
        
        
        /** reset to initial values.**/
        ASSIGN
            INCOLCOUNT = 0
            inNumCols  = 0
            inNumRows  = 0. 

        FOR EACH ttSheetData
            WHERE ttSheetData.relid EQ ''
            BREAK BY ttSheetData.cellRow
                  BY ttSheetData.cellCol:

            IF FIRST-OF(ttSheetData.cellRow) THEN
                ASSIGN
                    INCOLCOUNT  = 0
                    inNumRows   = inNumRows + 1.

            ASSIGN
                INCOLCOUNT = INCOLCOUNT + 1.

            IF LAST-OF(ttSheetData.cellCol) THEN
                inNumCols = MAXIMUM(INCOLCOUNT,inNumCols).

            ASSIGN 
                ttSheetData.relid = ttRelationShip.Id.

            /** Save on memory, perge empty cell values...**/
            IF ttSheetData.cellValue EQ '' THEN
                DELETE ttSheetData.
        END.
        
        /** Find the sheet and assign it's how many rows and columns it has.**/

        FIND FIRST ttSheet
            WHERE ttSheet.relid EQ ttRelationShip.Id
            NO-ERROR.

        IF AVAILABLE ttSheet THEN
        DO:
            
            ASSIGN
                ttSheet.numRows = inNumRows
                ttSheet.numCols = inNumCols.

        END.
    
        /*DATASET dsSheetData:WRITE-XML('file', ttRelationShip.Id + '.xml' ,TRUE).*/
        /*RUN TempTable2HTML (INPUT ttRelationShip.Id + '.html').*/
            
    END.

    RETURN.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-PackageTidy) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE PackageTidy Procedure 
PROCEDURE PackageTidy :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    
    MESSAGE SUBSTITUTE('Package Tidy: &1',
                       chExportDir).
    
    OS-DELETE VALUE(chExportDir) RECURSIVE.
    
    RETURN.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-PackageUnpack) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE PackageUnpack Procedure 
PROCEDURE PackageUnpack :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    
    DEFINE INPUT PARAMETER pcImportDirectory AS CHARACTER.
    DEFINE INPUT PARAMETER pcXSLXFilename    AS CHARACTER.
    
    DEFINE VARIABLE chOSCommand AS CHARACTER   NO-UNDO.
    DEFINE VARIABLE inOSError AS INTEGER     NO-UNDO.
    
    IF pcImportDirectory MATCHES '*/' THEN
        .
    ELSE
        pcImportDirectory = pcImportDirectory + '/'.

    ASSIGN
        chExportDir = SUBSTITUTE('&1&2/',
                                 pcImportDirectory,
                                 HEX-ENCODE(GENERATE-UUID)
                                 ).

    OS-CREATE-DIR VALUE(chExportDir).

    chOSCommand = SUBSTITUTE('unzip -q "&1" -d "&2"',
                             pcImportDirectory + pcXSLXFilename,
                             chExportDir ).

    MESSAGE chOSCommand.

    OS-COMMAND SILENT VALUE(chOSCommand).
    
    inOSError = OS-ERROR.

    IF inOSError NE 0 THEN 
        RETURN ERROR.
    
/*     RUN ParseContentTypeXML.                                                                                                                */
/*     RUN ParseRelationsXML(chExportDir + '_rels/.rels').                                                                                     */
/*                                                                                                                                             */
/*     chRelationTarget = GetRelalationshipTargetByType('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'). */


    RETURN.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-ParseContentTypeXML) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE ParseContentTypeXML Procedure 
PROCEDURE ParseContentTypeXML :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    MESSAGE 'Parsing Content Type XML file [Content_Types].xml'.
    
    DEFINE VARIABLE chContentTypeXMLFile AS CHARACTER   NO-UNDO INITIAL '[Content_Types].xml'.
    
    IF SEARCH(chExportDir + chContentTypeXMLFile) EQ ? THEN
    DO:
        MESSAGE chExportDir + chContentTypeXMLFile + ' is missing'
            VIEW-AS ALERT-BOX ERROR.
        RETURN ERROR.
    END.

    DATASET dsContentTypes:READ-XML('FILE',
                                    chExportDir + '/' + chContentTypeXMLFile,
                                    'EMPTY',
                                    ?,
                                    ?).

    FOR EACH ttOveride:
        IF SEARCH(chExportDir + ttOveride.PartName) EQ ? THEN 
        DO:
            MESSAGE SUBSTITUTE('Unable TO FIND: &1',
                               ttOveride.PartName). 
            NEXT.
        END.

        CASE ttOveride.ContentType:
            WHEN 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' THEN
                ASSIGN
                    chSAXHandlerProc = 'src/OfficeOpenXML/OpenOfficeSpreadsheetMLSAXHandler.p'.
        END CASE.

    END.
    
    RETURN.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-ParseRelationsXML) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE ParseRelationsXML Procedure 
PROCEDURE ParseRelationsXML :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    
    DEFINE INPUT PARAMETER pcRelFile AS CHARACTER.           
    
    MESSAGE  'Parsing Relations XML file ' + pcRelFile.

    DATASET dsRelationShips:READ-XML('FILE', pcRelFile,'APPEND',?,?).

    /*DATASET dsRelationShips:WRITE-XML('FILE', 'relationship.xml',TRUE).*/

    RETURN.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-ParseTargetPart) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE ParseTargetPart Procedure 
PROCEDURE ParseTargetPart :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    
    DEFINE INPUT PARAMETER pcTargetPartFilename AS CHARACTER.
    
    DEFINE VARIABLE hnSAXReader                 AS HANDLE      NO-UNDO.
    DEFINE VARIABLE hnSAXHandler                AS HANDLE      NO-UNDO.
        
    MESSAGE SUBSTITUTE('Parsing Target Part: &1', 
                       pcTargetPartFilename).

    RUN src/OfficeOpenXML/OfficeOpenSpreadsheetMLSAXHandler.p PERSISTENT SET hnSAXHandler.

    CREATE SAX-READER hnSAXReader.

    hnSAXReader:SET-INPUT-SOURCE('file', pcTargetPartFilename).

    hnSAXReader:HANDLER = hnSAXHandler.

    hnSAXReader:SAX-PARSE().

    DELETE PROCEDURE hnSAXHandler.
    DELETE OBJECT hnSAXReader.
    
    RETURN.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-SplitFilename) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE SplitFilename Procedure 
PROCEDURE SplitFilename :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    DEFINE INPUT  PARAMETER pchfilename  AS CHARACTER NO-UNDO.
    DEFINE OUTPUT PARAMETER opchPathName AS CHARACTER NO-UNDO.
    DEFINE OUTPUT PARAMETER opchFileName AS CHARACTER NO-UNDO.

    ASSIGN
        opchPathName = SUBSTRING(pchfilename,1, R-INDEX(pchfilename,'~/'))
        opchFileName = SUBSTRING(pchfilename, R-INDEX(pchfilename,'~/') + 1).
           
    RETURN.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-TempTable2html) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE TempTable2html Procedure 
PROCEDURE TempTable2html :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    DEFINE INPUT PARAMETER pcOutputFilename AS CHARACTER.

    FOR EACH ttSheetData
        BREAK BY ttSheetData.cellRow
              BY ttSheetData.cellCol:

        IF FIRST(ttSheetData.cellRow)  THEN
        DO: 
            OUTPUT STREAM sHTML TO VALUE(pcOutputFilename) NO-ECHO NO-CONVERT.
            PUT STREAM sHTML UNFORMATTED '<table>' SKIP.
        END.

        IF FIRST-OF(ttSheetData.cellRow)  THEN
        DO:
            PUT STREAM sHTML UNFORMATTED '~t<tr>' SKIP.
            
        END.

        IF FIRST-OF(ttSheetData.cellRow) AND FIRST-OF(ttSheetData.cellCol) THEN
            PUT STREAM sHTML UNFORMATTED '~t~t'.

        PUT STREAM sHTML UNFORMATTED SUBSTITUTE('<td data-cell="&1" data-row="&3" data-col="&4">&2</td>',
                                                ttSheetData.cellRef,
                                                ttSheetData.cellValue,
                                                ttSheetData.cellRow,
                                                ttSheetData.cellCol).

        IF LAST-OF(ttSheetData.cellRow)  THEN
        DO:
            PUT STREAM sHTML UNFORMATTED SKIP.
            PUT STREAM sHTML UNFORMATTED '~t</tr>' SKIP.
        END.

        IF LAST(ttSheetData.cellRow) THEN
        DO:
            PUT STREAM sHTML UNFORMATTED '</table>' SKIP.
            OUTPUT STREAM sHTML CLOSE.
        END.

    END.

    RETURN.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

/* ************************  Function Implementations ***************** */

&IF DEFINED(EXCLUDE-ConvertToDate) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION ConvertToDate Procedure 
FUNCTION ConvertToDate RETURNS DATE
  ( INPUT pinDate AS INTEGER ) :
/*------------------------------------------------------------------------------
  Purpose:  
    Notes:  
------------------------------------------------------------------------------*/

  RETURN ADD-INTERVAL(DATE(1,1,1900), pinDate - 2, 'days').

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-GetRelalationshipTargetByType) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION GetRelalationshipTargetByType Procedure 
FUNCTION GetRelalationshipTargetByType RETURNS CHARACTER
  ( INPUT pcType AS CHARACTER ) :
/*------------------------------------------------------------------------------
  Purpose:  
    Notes:  
------------------------------------------------------------------------------*/

    FIND FIRST ttRelationShip 
        WHERE ttRelationShip.TYPE EQ pcType
        NO-ERROR.

    IF NOT AVAILABLE ttRelationShip THEN
        RETURN "".   /* Function return value. */
    ELSE
        RETURN ttRelationShip.TARGET.          

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

&IF DEFINED(EXCLUDE-NormaliseNotation) = 0 &THEN

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _FUNCTION NormaliseNotation Procedure 
FUNCTION NormaliseNotation RETURNS CHARACTER
    (INPUT pchNumber AS CHARACTER):

    DEFINE VARIABLE deNotation  AS DECIMAL     NO-UNDO DECIMALS 10.
    DEFINE VARIABLE inPower     AS INTEGER     NO-UNDO .
    
    IF pchNumber MATCHES '*E*':U AND NUM-ENTRIES(pchNumber,'E') EQ 2 THEN
    DO:
        ASSIGN
            deNotation = DECIMAL(ENTRY(1,pchNumber,'E'))
            inPower    = INTEGER(ENTRY(2,pchNumber,'E'))
            NO-ERROR.
            
        IF ERROR-STATUS:ERROR THEN
            RETURN ERROR 'Invalid String Notation'.
        
        RETURN STRING(deNotation * EXP(10, inPower)).
    END.
    ELSE
        RETURN pchNumber.    

END FUNCTION.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ENDIF

