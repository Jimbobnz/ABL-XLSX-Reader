DEFINE TEMP-TABLE ttOveride XML-NODE-NAME 'Override':U
    FIELD PartName    AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'PartName'
    FIELD ContentType AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'ContentType'.

DEFINE TEMP-TABLE ttDefault XML-NODE-NAME 'Default':U
    FIELD Extension   AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'Extension'
    FIELD ContentType AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'ContentType'.

DEFINE DATASET dsContentTypes NAMESPACE-URI "http://schemas.openxmlformats.org/package/2006/content-types" XML-NODE-NAME 'Types'
    FOR ttDefault, ttOveride.

DEFINE TEMP-TABLE ttRelationShip  XML-NODE-NAME 'Relationship':U
    FIELD id         AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'Id'
    FIELD type       AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'Type'
    FIELD Target     AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'Target'
    FIELD TargetMode AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'TargetMode'.

DEFINE DATASET dsRelationShips NAMESPACE-URI "http://schemas.openxmlformats.org/package/2006/relationships" XML-NODE-NAME 'Relationships'
    FOR ttRelationShip.

DEFINE {1} {2} TEMP-TABLE ttSheet XML-NODE-NAME 'sheet':U
    FIELD sheetname AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'name'     LABEL 'Sheet Name':U 
    FIELD state     AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'state'    LABEL 'Visible State':U INITIAL 'visible':U
    FIELD sheetid   AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'sheetID'  LABEL 'Sheet Tab Id':U
    FIELD relid     AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'id'       LABEL 'Relationship Id':U
    FIELD numRows   AS INTEGER   LABEL 'Number of Rows':U
    FIELD numCols   AS INTEGER   LABEL 'Number of Columns':U.

DEFINE DATASET dsSheets XML-NODE-NAME 'sheets':U
    FOR ttSheet.
    
DEFINE {1} {2} TEMP-TABLE ttSharedString XML-NODE-NAME 'sharedstring':U
    FIELD lookupID      AS INTEGER      XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'lookupID' LABEL 'Lookup Position ID':U
    FIELD sharedstring  AS CHARACTER    XML-NODE-TYPE 'Text':U
    INDEX idxttSharedString IS PRIMARY
        LookupID.
    
DEFINE DATASET dsSharedStrings XML-NODE-NAME 'sharedstrings':U
    FOR ttSharedString.

DEFINE {1} {2} TEMP-TABLE ttSheetData XML-NODE-NAME 'SheetDataRow':U
    FIELD relid        AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'id'    LABEL 'Relationship Id':U INITIAL ''
    FIELD cellRow      AS INTEGER   XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'row':U LABEL 'Cell Row':U 
    FIELD cellCol      AS INTEGER   XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'col':U LABEL 'Cell Coll':U 
    FIELD cellRef      AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 'r':U   LABEL 'Cell Ref':U
    FIELD cellStyleID  AS INTEGER   XML-NODE-TYPE 'Attribute' XML-NODE-NAME 's':U   LABEL 'Cell Style':U
    FIELD sharedString AS CHARACTER XML-NODE-TYPE 'Attribute' XML-NODE-NAME 't':U   LABEL 'Shared String':U
    FIELD cellValue    AS CHARACTER XML-NODE-TYPE 'Text':U                          LABEL 'Cell Value':U
    INDEX idxSheetData IS PRIMARY cellRow cellCol
    INDEX idxRelID     relid.

DEFINE DATASET dsSheetData XML-NODE-NAME 'SheetData':U
    FOR ttSheetData.
