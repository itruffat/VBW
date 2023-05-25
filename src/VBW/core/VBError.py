class VB_ERROR(Exception):

    def __init__(self, msg, code):
        super().__init__(msg)
        self.code = code

    errorDictionary = {
        '5': '5-Invalid procedure call or argument',
        '6': '6-Overflow',
        '7': '7-Out of memory',
        '9': '9-Subscript out of range',
        '10': '10-Array fixed or temporarily locked',
        '11': '11-Division by zero',
        '13': '13-Type mismatch',
        '14': '14-Out of string space',
        '28': '28-Out of stack space',
        '35': '35-Sub or Function not defined',
        '48': '48-Error in loading DLL',
        '51': '51-Internal error',
        '53': '53-File not found',
        '57': '57-Device I/O error',
        '58': '58-File already exists',
        '61': '61-Disk full',
        '67': '67-Too many files',
        '70': '70-Permission denied',
        '75': '75-Path/File access error',
        '76': '76-Path not found',
        '91': '91-Object variable or With block variable not set',
        '92': '92-For loop not initialized',
        '94': '94-Invalid use of Null',
        '322': "322-Can't create necessary temporary file",
        '424': '424-Object required',
        '429': "429-ActiveX component can't create object",
        '430': "430-Class doesn't support Automation",
        '432': '432-File name or class name not found during Automation operation',
        '438': "438-Object doesn't support this property or method",
        '440': '440-Automation error',
        '445': "445-Object doesn't support this action",
        '446': "446-Object doesn't support named arguments",
        '447': "447-Object doesn't support current locale setting",
        '448': '448-Named argument not found',
        '449': '449-Argument not optional',
        '450': '450-Wrong number of arguments or invalid property assignment',
        '451': '451-Object not a collection',
        '453': '453-Specified DLL function not found',
        '455': '455-Code resource lock error',
        '457': '457-This key already associated with an element of this collection',
        '458': '458-Variable uses an Automation type not supported in VBScript',
        '500': '500-Variable is undefined',
        '501': '501-Illegal assignment',
        '502': '502-Object not safe for scripting',
        '503': '503-Object not safe for initializing',
        '1001': '1001-Out of memory',
        '1002': '1002-Syntax error',
        '1003': "1003-Expected ':'",
        '1004': "1004-Expected ';'",
        '1005': "1005-Expected '('",
        '1006': "1006-Expected ')'",
        '1007': "1007-Expected ']'",
        '1008': "1008-Expected '{'",
        '1009': "1009-Expected '}'",
        '1010': '1010-Expected identifier',
        '1011': "1011-Expected '='",
        '1012': "1012-Expected 'If'",
        '1013': "1013-Expected 'To'",
        '1014': "1014-Expected 'End'",
        '1015': "1015-Expected 'Function'",
        '1016': "1016-Expected 'Sub'",
        '1017': "1017-Expected 'Then'",
        '1018': "1018-Expected 'Wend'",
        '1019': "1019-Expected 'Loop'",
        '1020': "1020-Expected 'Next'",
        '1021': "1021-Expected 'Case'",
        '1022': "1022-Expected 'Select'",
        '1023': '1023-Expected expression',
        '1024': '1024-Expected statement',
        '1025': '1025-Expected end of statement',
        '1026': '1026-Expected integer constant',
        '1027': "1027-Expected 'While' or 'Until'",
        '1028': "1028-Expected 'While','Until',or end of statement",
        '1029': '1029-Too many locals or arguments',
        '1030': '1030-Identifier too long',
        '1031': '1031-Invalid number',
        '1032': '1032-Invalid character',
        '1033': '1033-Unterminated string constant',
        '1034': '1034-Unterminated comment',
        '1035': '1035-Nested comment',
        '1037': "1037-Invalid use of 'Me' keyword",
        '1038': "1038-'Loop' without 'Do'",
        '1039': "1039-Invalid 'Exit' statement",
        '1040': "1040-Invalid 'For' loop control variable",
        '1041': '1041-Name redefined',
        '1042': '1042-Must be first statement on the line',
        '1043': "1043-Can't assign to non-ByVal argument",
        '1044': "1044-Can't use parens when calling a Sub",
        '1045': '1045-Expected literal constant',
        '1046': "1046-Expected 'In'",
        '32766': '32766-True',
        '32767': '32767-False',
        '32811': '32811-Element not found'
    }
