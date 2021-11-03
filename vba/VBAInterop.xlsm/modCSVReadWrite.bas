Attribute VB_Name = "modCSVReadWrite"
Option Private Module
' VBA-CSV

' Copyright (C) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/VBA-CSV#readme
' This version at: https://github.com/PGS62/VBA-CSV/releases/tag/v0.15

'Installation:
'1) Import this module into your project (Open VBA Editor, Alt + F11; File > Import File).

'2) Add two references (In VBA Editor Tools > References)
'   Microsoft Scripting Runtime
'   Microsoft VBScript Regular Expressions 5.5 (or the latest version available)

'3) If you plan to call the functions from spreadsheet formulas then you might like to tell
'   Excel's Function Wizard about them by adding calls to RegisterCSVRead and RegisterCSVWrite
'   to the project's Workbook_Open event, which lives in the ThisWorkbook class module.

'        Private Sub Workbook_Open()
'            RegisterCSVRead
'            RegisterCSVWrite
'        End Sub

Option Explicit

Private m_FSO As Scripting.FileSystemObject

#If VBA7 And Win64 Then
    'for 64-bit Excel
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
#Else
    'for 32-bit Excel
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Private Enum enmErrorStyle
    es_ReturnString = 0
    es_RaiseError = 1
End Enum

Private Const m_ErrorStyle As Long = es_ReturnString

Private Const m_LBound As Long = 1 'Sets the array lower bounds of the return from CSVRead.
                                   'To return zero-based arrays, change the value of this constant to 0.

Private Enum enmSourceType
    st_File = 0
    st_URL = 1
    st_String = 2
End Enum

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CSVRead
' Purpose   : Returns the contents of a comma-separated file on disk as an array.
' Arguments
' FileName  : The full name of the file, including the path, or else a URL of a file, or else a string in
'             CSV format.
' ConvertTypes: Controls whether fields in the file are converted to typed values or remain as strings, and
'             sets the treatment of "quoted fields" and space characters.
'
'             ConvertTypes should be a string of zero or more letters from allowed characters `NDBETQ`.
'
'             The most commonly useful letters are:
'             1) `N` number fields are returned as numbers (Doubles).
'             2) `D` date fields (that respect DateFormat) are returned as Dates.
'             3) `B` fields matching TrueStrings or FalseStrings are returned as Booleans.
'
'             ConvertTypes is optional and defaults to the null string for no type conversion. `TRUE` is
'             equivalent to `NDB` and `FALSE` to the null string.
'
'             Three further options are available:
'             4) `E` fields that match Excel errors are converted to error values. There are fourteen of
'             these, including `#N/A`, `#NAME?`, `#VALUE!` and `#DIV/0!`.
'             5) `T` leading and trailing spaces are trimmed from fields. In the case of quoted fields,
'             this will not remove spaces between the quotes.
'             6) `Q` conversion happens for both quoted and unquoted fields; otherwise only unquoted fields
'             are converted.
'
'             For most files, correct type conversion can be achieved with ConvertTypes as a string which
'             applies for all columns, but type conversion can also be specified on a per-column basis.
'
'             Enter an array (or range) with two columns or two rows, column numbers on the left/top and
'             type conversion (subset of `NDBETQ`) on the right/bottom. Instead of column numbers, you can
'             enter strings matching the contents of the header row, and a column number of zero applies to
'             all columns not otherwise referenced.
'
'             For convenience when calling from VBA, you can pass an array of two element arrays such as
'             `Array(Array(0,"N"),Array(3,""),Array("Phone",""))` to convert all numbers in a file into
'             numbers in the return except for those in column 3 and in the column(s) headed "Phone".
' Delimiter : By default, CSVRead will try to detect a file's delimiter as the first instance of comma, tab,
'             semi-colon, colon or pipe found outside quoted regions in the first 10,000 characters of the
'             file. If it can't auto-detect the delimiter, it will assume comma. If your file includes a
'             different character or string delimiter you should pass that as the Delimiter argument.
'
'             Alternatively, enter `FALSE` as the delimiter to treat the file as "not a delimited file". In
'             this case the return will mimic how the file would appear in a text editor such as NotePad.
'             The file will be split into lines at all line breaks (irrespective of double quotes) and each
'             element of the return will be a line of the file.
' IgnoreRepeated: Whether delimiters which appear at the start of a line, the end of a line or immediately
'             after another delimiter should be ignored while parsing; useful for fixed-width files with
'             delimiter padding between fields.
' DateFormat: The format of dates in the file such as `Y-M-D` (the default), `M-D-Y` or `Y/M/D`. Also `ISO`
'             for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g.
'             2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time.
' Comment   : Rows that start with this string will be skipped while parsing.
' IgnoreEmptyLines: Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each
'             column will be assigned ShowMissingsAs for that empty row).
' HeaderRowNum: The row in the file containing headers. Type conversion is not applied to fields in the
'             header row, though leading and trailing spaces are trimmed.
'
'             This argument is most useful when calling from VBA, with SkipToRow set to one more than
'             HeaderRowNum. In that case the function returns the rows starting from SkipToRow, and the
'             header row is returned via the by-reference argument HeaderRow. Optional and defaults to 0.
' SkipToRow : The first row in the file that's included in the return. Optional and defaults to one more
'             than HeaderRowNum.
' SkipToCol : The column in the file at which reading starts. Optional and defaults to 1 to read from the
'             first column.
' NumRows   : The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the
'             end of the file are read.
' NumCols   : The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol
'             are read.
' TrueStrings: Indicates how `TRUE` values are represented in the file. May be a string, an array of strings
'             or a range containing strings; by default, `TRUE`, `True` and `true` are recognised.
' FalseStrings: Indicates how `FALSE` values are represented in the file. May be a string, an array of
'             strings or a range containing strings; by default, `FALSE`, `False` and `false` are
'             recognised.
' MissingStrings: Indicates how missing values are represented in the file. May be a string, an array of
'             strings or a range containing strings. By default, only an empty field (consecutive
'             delimiters) is considered missing.
' ShowMissingsAs: Fields which are missing in the file (consecutive delimiters) or match one of the
'             MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null
'             string or `#N/A!` error value can be good alternatives.
'
'             If NumRows is greater than the number of rows in the file then the return is "padded" with
'             the value of ShowMissingsAs. Likewise, if NumCols is greater than the number of columns in
'             the file.
' Encoding  : Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be
'             omitted and CSVRead will detect the file's encoding. If auto-detection does not work then
'             it's possible that the file is encoded `UTF-8` or `UTF-16` but without a byte option mark to
'             identify the encoding. Experiment with Encoding as each of `UTF-8` and `UTF-16`.
' DecimalSeparator: In many places in the world, floating point number decimals are separated with a comma
'             instead of a period (3,14 vs. 3.14). CSVRead can correctly parse these numbers by passing in
'             the DecimalSeparator as a comma, in which case comma ceases to be a candidate if the parser
'             needs to guess the Delimiter.
' HeaderRow : This by-reference argument is for use from VBA (as opposed to from Excel formulas). It is
'             populated with the contents of the header row, with no type conversion, though leading and
'             trailing spaces are removed.
'
' Notes     : See also companion function CSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
' -----------------------------------------------------------------------------------------------------------------------
Public Function CSVRead(ByVal FileName As String, Optional ByVal ConvertTypes As Variant = False, _
    Optional ByVal Delimiter As Variant, Optional ByVal IgnoreRepeated As Boolean, _
    Optional ByVal DateFormat As String = "Y-M-D", Optional ByVal Comment As String, _
    Optional ByVal IgnoreEmptyLines As Boolean, Optional ByVal HeaderRowNum As Long, _
    Optional ByVal SkipToRow As Long, Optional ByVal SkipToCol As Long = 1, _
    Optional ByVal NumRows As Long, Optional ByVal NumCols As Long, _
    Optional ByVal TrueStrings As Variant, Optional ByVal FalseStrings As Variant, _
    Optional ByVal MissingStrings As Variant, Optional ByVal ShowMissingsAs As Variant, _
    Optional ByVal Encoding As Variant, Optional ByVal DecimalSeparator As String, _
    Optional ByRef HeaderRow As Variant) As Variant


    Const DQ As String = """"
    Const Err_Delimiter As String = "Delimiter character must be passed as a string, FALSE for no delimiter. " & _
        "Omit to guess from file contents"
    Const Err_Delimiter2 As String = "Delimiter must have at least one character and cannot start with a double " & _
        "quote, line feed or carriage return"
    Const Err_FileEmpty As String = "File is empty"
    Const Err_FunctionWizard  As String = "Disabled in Function Wizard"
    Const Err_NumCols As String = "NumCols must be positive to read a given number of columns, or zero or omitted " & _
        "to read all columns from SkipToCol to the maximum column encountered."
    Const Err_NumRows As String = "NumRows must be positive to read a given number of rows, or zero or omitted to " & _
        "read all rows from SkipToRow to the end of the file."
    Const Err_Seps1 As String = "DecimalSeparator must be a single character"
    Const Err_Seps2 As String = "DecimalSeparator must not be equal to the first character of Delimiter or to a " & _
        "line-feed or carriage-return"
    Const Err_SkipToCol As String = "SkipToCol must be at least 1."
    Const Err_SkipToRow As String = "SkipToRow must be at least 1."
    Const Err_Comment As String = "Comment must not contain double-quote, line feed or carriage return"
    Const Err_HeaderRowNum As String = "HeaderRowNum must be greater than or equal to zero and less than or equal to SkipToRow"
    
    Dim AcceptWithoutTimeZone As Boolean
    Dim AcceptWithTimeZone As Boolean
    Dim Adj As Long
    Dim AnyConversion As Boolean
    Dim AnySentinels As Boolean
    Dim CallingFromWorksheet As Boolean
    Dim CharSet As String
    Dim ColByColFormatting As Boolean
    Dim ColIndexes() As Long
    Dim ConvertQuoted As Boolean
    Dim CSVContents As String
    Dim CTDict As Scripting.Dictionary
    Dim DateOrder As Long
    Dim DateSeparator As String
    Dim Err_StringTooLong As String
    Dim ErrRet As String
    Dim i As Long
    Dim ISO8601 As Boolean
    Dim j As Long
    Dim k As Long
    Dim lengths() As Long
    Dim m As Long
    Dim MaxSentinelLength As Long
    Dim MSLIA As Long
    Dim NeedToFill As Boolean
    Dim NotDelimited As Boolean
    Dim NumColsFound As Long
    Dim NumColsInReturn As Long
    Dim NumFields As Long
    Dim NumRowsFound As Long
    Dim NumRowsInReturn As Long
    Dim QuoteCounts() As Long
    Dim Ragged As Boolean
    Dim ReturnArray() As Variant
    Dim RowIndexes() As Long
    Dim Sentinels As Scripting.Dictionary
    Dim SepStandard As Boolean
    Dim ShowBooleansAsBooleans As Boolean
    Dim ShowDatesAsDates As Boolean
    Dim ShowErrorsAsErrors As Boolean
    Dim ShowMissingsAsEmpty As Boolean
    Dim ShowNumbersAsNumbers As Boolean
    Dim SourceType As enmSourceType
    Dim Starts() As Long
    Dim strDelimiter As String
    Dim Stream As Object 'either ADODB.Stream or Scripting.TextStram
    Dim SysDateOrder As Long
    Dim SysDateSeparator As String
    Dim SysDecimalSeparator As String
    Dim TempFile As String
    Dim TrimFields As Boolean
    Dim TriState As Long
    Dim useADODB As Boolean
    
    On Error GoTo ErrHandler

    SourceType = InferSourceType(FileName)

    'Download file from internet to local temp folder
    If SourceType = st_URL Then
        TempFile = Environ$("Temp") & "\VBA-CSV\Downloads\DownloadedFile.csv"
        FileName = Download(FileName, TempFile)
        SourceType = st_File
    End If

    'Parse and validate inputs...
    If SourceType <> st_String Then
        ParseEncoding FileName, Encoding, TriState, CharSet, useADODB
    End If

    If VarType(Delimiter) = vbBoolean Then
        If Not Delimiter Then
            NotDelimited = True
        Else
            Throw Err_Delimiter
        End If
    ElseIf VarType(Delimiter) = vbString Then
        If Len(Delimiter) = 0 Then
            strDelimiter = InferDelimiter(SourceType, FileName, TriState, DecimalSeparator)
        ElseIf Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
            Throw Err_Delimiter2
        Else
            strDelimiter = Delimiter
        End If
    ElseIf IsEmpty(Delimiter) Or IsMissing(Delimiter) Then
        strDelimiter = InferDelimiter(SourceType, FileName, TriState, DecimalSeparator)
    Else
        Throw Err_Delimiter
    End If

    SysDecimalSeparator = Application.DecimalSeparator
    If DecimalSeparator = vbNullString Then DecimalSeparator = SysDecimalSeparator
    If DecimalSeparator = SysDecimalSeparator Then
        SepStandard = True
    ElseIf Len(DecimalSeparator) <> 1 Then
        Throw Err_Seps1
    ElseIf DecimalSeparator = strDelimiter Or DecimalSeparator = vbLf Or DecimalSeparator = vbCr Then
        Throw Err_Seps2
    End If

    Set CTDict = New Scripting.Dictionary

    ParseConvertTypes ConvertTypes, ShowNumbersAsNumbers, _
        ShowDatesAsDates, ShowBooleansAsBooleans, ShowErrorsAsErrors, _
        ConvertQuoted, TrimFields, ColByColFormatting, HeaderRowNum, CTDict

    Set Sentinels = New Scripting.Dictionary
    MakeSentinels Sentinels, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
        ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings
    
    If ShowDatesAsDates Then
        ParseDateFormat DateFormat, DateOrder, DateSeparator, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone
        SysDateOrder = Application.International(xlDateOrder)
        SysDateSeparator = Application.International(xlDateSeparator)
    End If

    If HeaderRowNum < 0 Then Throw Err_HeaderRowNum
    If SkipToRow = 0 Then SkipToRow = HeaderRowNum + 1
    If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
    If SkipToCol = 0 Then SkipToCol = 1
    If SkipToRow < 1 Then Throw Err_SkipToRow
    If SkipToCol < 1 Then Throw Err_SkipToCol
    If NumRows < 0 Then Throw Err_NumRows
    If NumCols < 0 Then Throw Err_NumCols

    If HeaderRowNum > SkipToRow Then Throw Err_HeaderRowNum
       
    If InStr(Comment, DQ) > 0 Or InStr(Comment, vbLf) > 0 Or InStr(Comment, vbCrLf) > 0 Then Throw Err_Comment
    'End of input validation
    
    CallingFromWorksheet = TypeName(Application.Caller) = "Range"
    
    If CallingFromWorksheet Then
        If FunctionWizardActive() Then
            CSVRead = "#" & Err_FunctionWizard & "!"
            Exit Function
        End If
    End If
    
    If NotDelimited Then
        CSVRead = ParseTextFile(FileName, SourceType <> st_String, useADODB, CharSet, TriState, SkipToRow, NumRows, CallingFromWorksheet)
        Exit Function
    End If
          
    If SourceType = st_String Then
        CSVContents = FileName
        
        ParseCSVContents CSVContents, useADODB, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
            IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, _
            NumFields, Ragged, Starts, lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow
    Else
        If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
            
        If useADODB Then
            Set Stream = CreateObject("ADODB.Stream")
            Stream.CharSet = CharSet
            Stream.Open
            Stream.LoadFromFile FileName
            If Stream.EOS Then Throw Err_FileEmpty
        Else
            Set Stream = m_FSO.GetFile(FileName).OpenAsTextStream(ForReading, TriState)
            If Stream.atEndOfStream Then Throw Err_FileEmpty
        End If

        If SkipToRow = 1 And NumRows = 0 Then
            CSVContents = ReadAllFromStream(Stream)
            Stream.Close
            
            ParseCSVContents CSVContents, useADODB, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, _
                Ragged, Starts, lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow
        Else
            CSVContents = ParseCSVContents(Stream, useADODB, DQ, strDelimiter, Comment, IgnoreEmptyLines, _
                IgnoreRepeated, SkipToRow, HeaderRowNum, NumRows, NumRowsFound, NumColsFound, NumFields, _
                Ragged, Starts, lengths, RowIndexes, ColIndexes, QuoteCounts, HeaderRow)
            Stream.Close
        End If
    End If
                     
    If NumCols = 0 Then
        NumColsInReturn = NumColsFound - SkipToCol + 1
        If NumColsInReturn <= 0 Then
            Throw "SkipToCol (" & CStr(SkipToCol) & _
                ") exceeds the number of columns in the file (" & CStr(NumColsFound) & ")"
        End If
    Else
        NumColsInReturn = NumCols
    End If
    If NumRows = 0 Then
        NumRowsInReturn = NumRowsFound
    Else
        NumRowsInReturn = NumRows
    End If
        
    AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowBooleansAsBooleans Or ShowErrorsAsErrors Or TrimFields
        
    Adj = m_LBound - 1
    ReDim ReturnArray(1 + Adj To NumRowsInReturn + Adj, 1 + Adj To NumColsInReturn + Adj)
    MSLIA = MaxStringLengthInArray()
    ShowMissingsAsEmpty = IsEmpty(ShowMissingsAs)
        
    For k = 1 To NumFields
        i = RowIndexes(k)
        j = ColIndexes(k) - SkipToCol + 1
        If j >= 1 And j <= NumColsInReturn Then
            If CallingFromWorksheet Then
                If lengths(k) > MSLIA Then
                    Err_StringTooLong = "The file has a field (row " + CStr(i + SkipToRow - 1) & _
                        ", column " & CStr(j + SkipToCol - 1) & ") of length " + Format$(lengths(k), "###,###")
                    If MSLIA >= 32767 Then
                        Err_StringTooLong = Err_StringTooLong & ". Excel cells cannot contain strings longer than " + Format$(MSLIA, "####,####")
                    Else
                        Err_StringTooLong = Err_StringTooLong & _
                            ". An array containing a string longer than " + Format$(MSLIA, "###,###") + _
                            " cannot be returned from VBA to an Excel worksheet"
                    End If
                    Throw Err_StringTooLong
                End If
            End If
        
            If ColByColFormatting Then
                ReturnArray(i + Adj, j + Adj) = Mid$(CSVContents, Starts(k), lengths(k))
            Else
                ReturnArray(i + Adj, j + Adj) = ConvertField(Mid$(CSVContents, Starts(k), lengths(k)), AnyConversion, _
                    lengths(k), TrimFields, DQ, QuoteCounts(k), ConvertQuoted, ShowNumbersAsNumbers, SepStandard, _
                    DecimalSeparator, SysDecimalSeparator, ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, _
                    AcceptWithTimeZone, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, AnySentinels, _
                    Sentinels, MaxSentinelLength, ShowMissingsAs)
            End If
            
            'File has variable number of fields per line...
            If Ragged Then
                If Not ShowMissingsAsEmpty Then
                    If k = NumFields Then
                        NeedToFill = j < NumColsInReturn
                    ElseIf RowIndexes(k + 1) > RowIndexes(k) Then
                        NeedToFill = j < NumColsInReturn
                    Else
                        NeedToFill = False
                    End If
                    If NeedToFill Then
                        For m = j + 1 To NumColsInReturn
                            ReturnArray(i + Adj, m + Adj) = ShowMissingsAs
                        Next m
                    End If
                End If
            End If
        End If
    Next k

    If Ragged Then
        If Not IsEmpty(HeaderRow) Then
            If NCols(HeaderRow) < NCols(ReturnArray) + SkipToCol - 1 Then
                ReDim Preserve HeaderRow(1 To 1, 1 To NCols(ReturnArray) + SkipToCol - 1)
            End If
        End If
    End If
    If SkipToCol > 1 Then
        If Not IsEmpty(HeaderRow) Then
            Dim HeaderRowTruncated() As String
            ReDim HeaderRowTruncated(1 To 1, 1 To NumColsInReturn)
            For i = 1 To NumColsInReturn
                HeaderRowTruncated(1, i) = HeaderRow(1, i + SkipToCol - 1)
            Next i
            HeaderRow = HeaderRowTruncated
        End If
    End If
    
    'In this case no type conversion should be applied to the top row of the return
    If HeaderRowNum = SkipToRow Then
        If AnyConversion Then
            For i = 1 To MinLngs(NCols(HeaderRow), NumColsInReturn)
                ReturnArray(1 + Adj, i + Adj) = HeaderRow(1, i)
            Next
        End If
    End If

    If ColByColFormatting Then
        Dim CT As Variant
        Dim Field As String
        Dim NC As Long
        Dim NCH As Long
        Dim NR As Long
        Dim QC As Long
        Dim UnQuotedHeader As String
        NR = NRows(ReturnArray)
        NC = NCols(ReturnArray)
        If IsEmpty(HeaderRow) Then
            NCH = 0
        Else
            NCH = NCols(HeaderRow) 'possible that headers has fewer than expected columns if file is ragged
        End If

        For j = 1 To NC
            If j + SkipToCol - 1 <= NCH Then
                UnQuotedHeader = HeaderRow(1, j + SkipToCol - 1)
            Else
                UnQuotedHeader = -1 'Guaranteed not to be a key of the Dictionary
            End If
            If CTDict.Exists(j + SkipToCol - 1) Then
                CT = CTDict.item(j + SkipToCol - 1)
            ElseIf CTDict.Exists(UnQuotedHeader) Then
                CT = CTDict.item(UnQuotedHeader)
            ElseIf CTDict.Exists(0) Then
                CT = CTDict.item(0)
            Else
                CT = False
            End If
            
            ParseCTString CT, ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
                ShowErrorsAsErrors, ConvertQuoted, TrimFields
            
            AnyConversion = ShowNumbersAsNumbers Or ShowDatesAsDates Or _
                ShowBooleansAsBooleans Or ShowErrorsAsErrors
                
            Set Sentinels = New Scripting.Dictionary
            
            MakeSentinels Sentinels, MaxSentinelLength, AnySentinels, ShowBooleansAsBooleans, _
                ShowErrorsAsErrors, ShowMissingsAs, TrueStrings, FalseStrings, MissingStrings

            For i = 1 To NR
                If Not IsEmpty(ReturnArray(i + Adj, j + Adj)) Then
                    Field = CStr(ReturnArray(i + Adj, j + Adj))
                    QC = CountQuotes(Field, DQ)
                    ReturnArray(i + Adj, j + Adj) = ConvertField(Field, AnyConversion, _
                        Len(ReturnArray(i + Adj, j + Adj)), TrimFields, DQ, QC, ConvertQuoted, _
                        ShowNumbersAsNumbers, SepStandard, DecimalSeparator, SysDecimalSeparator, _
                        ShowDatesAsDates, ISO8601, AcceptWithoutTimeZone, AcceptWithTimeZone, DateOrder, _
                        DateSeparator, SysDateOrder, SysDateSeparator, AnySentinels, Sentinels, _
                        MaxSentinelLength, ShowMissingsAs)
                End If
            Next i
        Next j
    End If

    'Pad if necessary
    If Not ShowMissingsAsEmpty Then
        If NumColsInReturn > NumColsFound - SkipToCol + 1 Then
            For i = 1 To NumRowsInReturn
                For j = NumColsFound - SkipToCol + 2 To NumColsInReturn
                    ReturnArray(i + Adj, j + Adj) = ShowMissingsAs
                Next j
            Next i
        End If
        If NumRowsInReturn > NumRowsFound Then
            For i = NumRowsFound + 1 To NumRowsInReturn
                For j = 1 To NumColsInReturn
                    ReturnArray(i + Adj, j + Adj) = ShowMissingsAs
                Next j
            Next i
        End If
    End If

    CSVRead = ReturnArray

    Exit Function

ErrHandler:
    ErrRet = "#CSVRead: " & Err.Description & "!"
    If m_ErrorStyle = es_ReturnString Then
        CSVRead = ErrRet
    Else
        Throw ErrRet
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVRead
' Purpose    : Register the function CSVRead with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterCSVRead()
    Const Description As String = "Returns the contents of a comma-separated file on disk as an array."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 19)
    ArgDescs(1) = "The full name of the file, including the path, or else a URL of a file, or else a string in CSV " & _
                  "format."
    ArgDescs(2) = "Type conversion: Boolean or string, subset of letters NDBETQ. N = convert Numbers, D = convert " & _
                  "Dates, B = Convert Booleans, E = convert Excel errors, T = trim leading & trailing spaces, Q = " & _
                  "quoted fields also converted. TRUE = NDB, FALSE = no conversion."
    ArgDescs(3) = "Delimiter string. Defaults to the first instance of comma, tab, semi-colon, colon or pipe found " & _
                  "outside quoted regions within the first 10,000 characters. Enter FALSE to  see the file's " & _
                  "contents as would be displayed in a text editor."
    ArgDescs(4) = "Whether delimiters which appear at the start of a line, the end of a line or immediately after " & _
                  "another delimiter should be ignored while parsing; useful for fixed-width files with delimiter " & _
                  "padding between fields."
    ArgDescs(5) = "The format of dates in the file such as `Y-M-D` (the default), `M-D-Y` or `Y/M/D`. Also `ISO` " & _
                  "for ISO8601 (e.g., 2021-08-26T09:11:30) or `ISOZ` (time zone given e.g. " & _
                  "2021-08-26T13:11:30+05:00), in which case dates-with-time are returned in UTC time."
    ArgDescs(6) = "Rows that start with this string will be skipped while parsing."
    ArgDescs(7) = "Whether empty rows/lines in the file should be skipped while parsing (if `FALSE`, each column " & _
                  "will be assigned ShowMissingsAs for that empty row)."
    ArgDescs(8) = "The row in the file containing headers. Optional and defaults to 0. Type conversion is not " & _
                  "applied to fields in the header row, though leading and trailing spaces are trimmed."
    ArgDescs(9) = "The first row in the file that's included in the return. Optional and defaults to one more than " & _
                  "HeaderRowNum."
    ArgDescs(10) = "The column in the file at which reading starts. Optional and defaults to 1 to read from the " & _
                   "first column."
    ArgDescs(11) = "The number of rows to read from the file. If omitted (or zero), all rows from SkipToRow to the " & _
                   "end of the file are read."
    ArgDescs(12) = "The number of columns to read from the file. If omitted (or zero), all columns from SkipToCol " & _
                   "are read."
    ArgDescs(13) = "Indicates how `TRUE` values are represented in the file. May be a string, an array of strings " & _
                   "or a range containing strings; by default, `TRUE`, `True` and `true` are recognised."
    ArgDescs(14) = "Indicates how `FALSE` values are represented in the file. May be a string, an array of strings " & _
                   "or a range containing strings; by default, `FALSE`, `False` and `false` are recognised."
    ArgDescs(15) = "Indicates how missing values are represented in the file. May be a string, an array of strings " & _
                   "or a range containing strings. By default, only an empty field (consecutive delimiters) is " & _
                   "considered missing."
    ArgDescs(16) = "Fields which are missing in the file (consecutive delimiters) or match one of the " & _
                   "MissingStrings are returned in the array as ShowMissingsAs. Defaults to Empty, but the null " & _
                   "string or `#N/A!` error value can be good alternatives."
    ArgDescs(17) = "Allowed entries are `ASCII`, `ANSI`, `UTF-8`, or `UTF-16`. For most files this argument can be " & _
                   "omitted and CSVRead will detect the file's encoding."
    ArgDescs(18) = "The character that represents a decimal point. If omitted, then the value from Windows " & _
                   "regional settings is used."
    ArgDescs(19) = "For use from VBA only."
    Application.MacroOptions "CSVRead", Description, , , , , , , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function CSVRead failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferSourceType
' Purpose    : Guess whether FileName is in fact a file, a URL or a string in CSV format
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferSourceType(FileName As String) As enmSourceType

    On Error GoTo ErrHandler
    If InStr(FileName, vbLf) > 0 Then 'vbLf and vbCr are not permitted characters in file names or urls
        InferSourceType = st_String
    ElseIf InStr(FileName, vbCr) > 0 Then
        InferSourceType = st_String
    ElseIf Mid$(FileName, 2, 2) = ":\" Then
        InferSourceType = st_File
    ElseIf Left$(FileName, 2) = "\\" Then
        InferSourceType = st_File
    ElseIf Left$(FileName, 8) = "https://" Then
        InferSourceType = st_URL
    ElseIf Left$(FileName, 7) = "http://" Then
        InferSourceType = st_URL
    Else
        'Doesn't look like either file with path, url or string in CSV format
        InferSourceType = st_String
        If Len(FileName) < 1000 Then
            If FileExists(FileName) Then 'file exists in current working directory
                InferSourceType = st_File
            End If
        End If
    End If

    Exit Function
ErrHandler:
    Throw "#InferSourceType: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MaxStringLengthInArray
' Purpose    : Different versions of Excel have different limits for the longest string that can be an element of an
'              array passed from a VBA UDF back to Excel. I know the limit is 255 for Excel 2010 and earlier, and is
'              32,767 for Excel 365 (as of Sep 2021). But don't yet know the limit for Excel 2013, 2016 and 2019.
' Tried to get info from StackOverflow, without much joy:
' https://stackoverflow.com/questions/69303804/excel-versions-and-limits-on-the-length-of-string-elements-in-arrays-returned-by
' -----------------------------------------------------------------------------------------------------------------------
Private Function MaxStringLengthInArray() As Long
    Static res As Long
    If res = 0 Then
        Select Case Val(Application.Version)
            Case Is <= 14 'Excel 2010
                res = 255
            Case 15
                res = 32767 'Don't yet know if this is correct for Excel 2013
            Case Else
                res = 32767 'Excel 2016, 2019, 365. Hopefully these versions (which all _
                             return 16 as Application.Version) have the same limit.
        End Select
    End If
    MaxStringLengthInArray = res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Download
' Purpose   : Downloads bits from the Internet and saves them to a file.
'             See https://msdn.microsoft.com/en-us/library/ms775123(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Private Function Download(URLAddress As String, ByVal FileName As String) As String
    Dim EN As Long
    Dim ErrString As String
    Dim res As Long
    Dim TargetFolder As String

    On Error GoTo ErrHandler
    
    TargetFolder = FileFromPath(FileName, False)
    CreatePath TargetFolder
    If FileExists(FileName) Then
        On Error Resume Next
        FileDelete FileName
        EN = Err.Number
        On Error GoTo ErrHandler
        If EN <> 0 Then
            Throw "Cannot download from URL '" + URLAddress + "' because target file '" + FileName + _
                "' already exists and cannot be deleted. Is the target file open in a program such as Excel?"
        End If
    End If
    
    res = URLDownloadToFile(0, URLAddress, FileName, 0, 0)
    If res <> 0 Then
        ErrString = ParseDownloadError(CLng(res))
        Throw "Windows API function URLDownloadToFile returned error code " & CStr(res) & _
            " with description '" & ErrString & "'"
    End If
    If Not FileExists(FileName) Then Throw "Windows API function URLDownloadToFile did not report an error, " & _
        "but appears to have not successfuly downloaded a file from " & URLAddress & " to " & FileName
        
    Download = FileName

    Exit Function
ErrHandler:
    Throw "#Download: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDownloadError, sub of Download
'              https://www.vbforums.com/showthread.php?882757-URLDownloadToFile-error-codes
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseDownloadError(ErrNum As Long) As String
    Dim ErrString As String
    Select Case ErrNum
        Case &H80004004
            ErrString = "Aborted"
        Case &H800C0001
            ErrString = "Destination File Exists"
        Case &H800C0002
            ErrString = "Invalid Url"
        Case &H800C0003
            ErrString = "No Session"
        Case &H800C0004
            ErrString = "Cannot Connect"
        Case &H800C0005
            ErrString = "Resource Not Found"
        Case &H800C0006
            ErrString = "Object Not Found"
        Case &H800C0007
            ErrString = "Data Not Available"
        Case &H800C0008
            ErrString = "Download Failure"
        Case &H800C0009
            ErrString = "Authentication Required"
        Case &H800C000A
            ErrString = "No Valid Media"
        Case &H800C000B
            ErrString = "Connection Timeout"
        Case &H800C000C
            ErrString = "Invalid Request"
        Case &H800C000D
            ErrString = "Unknown Protocol"
        Case &H800C000E
            ErrString = "Security Problem"
        Case &H800C000F
            ErrString = "Cannot Load Data"
        Case &H800C0010
            ErrString = "Cannot Instantiate Object"
        Case &H800C0014
            ErrString = "Redirect Failed"
        Case &H800C0015
            ErrString = "Redirect To Dir"
        Case &H800C0016
            ErrString = "Cannot Lock Request"
        Case Else
            ErrString = "Unknown"
    End Select
    ParseDownloadError = ErrString
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReadAllFromStream
' Purpose    : Handles both ADOB.Stream and Scripting.TextStream. Note that ADODB.ReadText(-1) to read all of a stream
'              in a single operation has _very_ poor performance for large files, but reading 10,000 characters at a time
'              in a loop appears to solve that problem.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ReadAllFromStream(Stream As Object) As String
      
    Dim Chunk As String
    Dim Contents As String
    Dim i As Long
    Const ChunkSize As Long = 10000

    On Error GoTo ErrHandler
    If TypeName(Stream) = "Stream" Then
        Contents = String(ChunkSize, " ")

        i = 1
        Do While Not Stream.EOS
            Chunk = Stream.ReadText(ChunkSize)
            If i - 1 + Len(Chunk) > Len(Contents) Then
                Contents = Contents & String(i - 1 + Len(Chunk), " ")
            End If

            Mid$(Contents, i, Len(Chunk)) = Chunk
            i = i + Len(Chunk)
        Loop

        If (i - 1) < Len(Contents) Then
            Contents = Left$(Contents, i - 1)
        End If

        ReadAllFromStream = Contents
    ElseIf TypeName(Stream) = "TextStream" Then
        ReadAllFromStream = Stream.ReadAll
    Else
        Throw "Stream has unknown type: " + TypeName(Stream)
    End If

    Exit Function
ErrHandler:
    Throw "#ReadAllFromStream: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseEncoding
' Purpose    : Set by-ref arguments
' Parameters :
'  FileName:
'  Encoding: Optional argument passed in to CSVRead. If not passed, we delegate to DetectEncoding.
'  TriState: Set by reference. Needed only when we read files using Scripting.TextStream, i.e. when useADODB is False.
'  CharSet : Set by reference. Needed only when we read files using ADODB.Stream, i.e. when useADODB is True.
'  useADODB: Should file be read via ADODB.Stream, which is capable of reading UTF-8 files
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseEncoding(FileName As String, Encoding As Variant, ByRef TriState As Long, _
    ByRef CharSet As String, ByRef useADODB As Boolean)

    Const Err_Encoding As String = "Encoding argument can usually be omitted, but otherwise Encoding be " & _
        "either ""ASCII"", ""ANSI"", ""UTF-8"", ""UTF-8-BOM"", ""UTF-16"" or ""UTF-16-BOM"""
    
    On Error GoTo ErrHandler
    If IsEmpty(Encoding) Or IsMissing(Encoding) Then
        DetectEncoding FileName, TriState, CharSet, useADODB
    ElseIf VarType(Encoding) = vbString Then
        Select Case UCase$(Replace(Replace(Encoding, "-", vbNullString), " ", vbNullString))
            Case "ASCII"
                CharSet = "ascii" 'not actually relevant, since we won't use ADODB
                TriState = TristateFalse
                useADODB = False
            Case "ANSI"
                CharSet = "_autodetect_all" 'not actually relevant, since we won't use ADODB
                TriState = TristateFalse
                useADODB = False
            Case "UTF8"
                CharSet = "utf-8"
                TriState = TristateFalse
                useADODB = True 'Use ADODB because Scripting.TextStream can't cope with UTF-8
            Case "UTF8BOM"
                CharSet = "utf-8"
                TriState = TristateFalse
                useADODB = True 'Use ADODB because Scripting.TextStream can't cope with UTF-8
            Case "UTF16"
                CharSet = "utf-16"
                TriState = TristateTrue
                useADODB = False
            Case "UTF16BOM"
                CharSet = "utf-16"
                TriState = TristateTrue
                useADODB = False
            Case Else
                Throw Err_Encoding
        End Select
    Else
        Throw Err_Encoding
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseEncoding: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsCTValid
' Purpose    : Is a "Convert Types string" (which can in fact be either a string or a Boolean) valid?
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsCTValid(CT As Variant) As Boolean

    Static rx As VBScript_RegExp_55.RegExp

    On Error GoTo ErrHandler
    If rx Is Nothing Then
        Set rx = New RegExp
        With rx
            .IgnoreCase = True
            .Pattern = "^[NDBETQ]*$"
            .Global = False        'Find first match only
        End With
    End If

    If VarType(CT) = vbBoolean Then
        IsCTValid = True
    ElseIf VarType(CT) = vbString Then
        IsCTValid = rx.Test(CT)
    Else
        IsCTValid = False
    End If

    Exit Function
ErrHandler:
    Throw "#IsCTValid: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CTsEqual
' Purpose    : Test if two CT strings (strings to define type conversion) are equal, i.e. will have the same effect
' -----------------------------------------------------------------------------------------------------------------------
Private Function CTsEqual(CT1 As Variant, CT2 As Variant) As Boolean
    On Error GoTo ErrHandler
    If VarType(CT1) = VarType(CT2) Then
        If CT1 = CT2 Then
            CTsEqual = True
            Exit Function
        End If
    End If
    CTsEqual = StandardiseCT(CT1) = StandardiseCT(CT2)
    Exit Function
ErrHandler:
    Throw "#CTsEqual: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StandardiseCT
' Purpose    : Put a CT string into standard form so that two such can be compared.
' -----------------------------------------------------------------------------------------------------------------------
Private Function StandardiseCT(CT As Variant) As String
    On Error GoTo ErrHandler
    If VarType(CT) = vbBoolean Then
        If CT Then
            StandardiseCT = "BDN"
        Else
            StandardiseCT = vbNullString
        End If
        Exit Function
    ElseIf VarType(CT) = vbString Then
        StandardiseCT = IIf(InStr(1, CT, "B", vbTextCompare), "B", vbNullString) & _
            IIf(InStr(1, CT, "D", vbTextCompare), "D", vbNullString) & _
            IIf(InStr(1, CT, "E", vbTextCompare), "E", vbNullString) & _
            IIf(InStr(1, CT, "N", vbTextCompare), "N", vbNullString) & _
            IIf(InStr(1, CT, "Q", vbTextCompare), "Q", vbNullString) & _
            IIf(InStr(1, CT, "T", vbTextCompare), "T", vbNullString)
    End If

    Exit Function
ErrHandler:
    Throw "#StandardiseCT: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OneDArrayToTwoDArray
' Purpose    : Convert 1-d array of 2-element 1-d arrays into a 1-based, two-column, 2-d array.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OneDArrayToTwoDArray(x As Variant) As Variant
    Const Err_1DArray As String = "If ConvertTypes is given as a 1-dimensional array, each element must " & _
        "be a 1-dimensional array with two elements"

    Dim i As Long
    Dim k As Long
    Dim TwoDArray() As Variant
    On Error GoTo ErrHandler
    ReDim TwoDArray(1 To UBound(x) - LBound(x) + 1, 1 To 2)
    For i = LBound(x) To UBound(x)
        k = k + 1
        If Not IsArray(x(i)) Then Throw Err_1DArray
        If NumDimensions(x(i)) <> 1 Then Throw Err_1DArray
        If UBound(x(i)) - LBound(x(i)) <> 1 Then Throw Err_1DArray
        TwoDArray(k, 1) = x(i)(LBound(x(i)))
        TwoDArray(k, 2) = x(i)(1 + LBound(x(i)))
    Next i
    OneDArrayToTwoDArray = TwoDArray
    Exit Function
ErrHandler:
    Throw "#OneDArrayToTwoDArray: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseConvertTypes
' Purpose    : There is flexibility in how the ConvertTypes argument is provided to CSVRead:
'              a) As a string or Boolean for the same type conversion rules for every field in the file; or
'              b) An array to define different type conversion rules by column, in which case ConvertTypes can be passed
'                 as a two-column or two-row array (convenient from Excel) or as an array of two-element arrays
'                 (convenient from VBA).
'              If an array, then the left col(or top row or first element) can contain either column numbers or strings
'              that match the elements of the SkipToRow row of the file
'
' Parameters :
'  ConvertTypes          :
'  ShowNumbersAsNumbers  : Set only if ConvertTypes is not an array
'  ShowDatesAsDates      : Set only if ConvertTypes is not an array
'  ShowBooleansAsBooleans: Set only if ConvertTypes is not an array
'  ShowErrorsAsErrors    : Set only if ConvertTypes is not an array
'  ConvertQuoted         : Set only if ConvertTypes is not an array
'  TrimFields            : Set only if ConvertTypes is not an array
'  ColByColFormatting    : Set to True if ConvertTypes is an array
'  HeaderRowNum          : As passed to CSVRead, used to throw an error if HeaderRowNum has not been specified when
'                          it needs to have been.
'  CTDict                : Set to a dictionary keyed on the elements of the left column (or top row) of ConvertTypes,
'                          each element containing the corresponding right (or bottom) element.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseConvertTypes(ByVal ConvertTypes As Variant, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean, _
    ByRef ColByColFormatting As Boolean, HeaderRowNum As Long, ByRef CTDict As Scripting.Dictionary)
    
    Const Err_2D As String = "If ConvertTypes is given as a two dimensional array then the " & _
        " lower bounds in each dimension must be 1"
    Const Err_Ambiguous As String = "ConvertTypes is ambiguous, it can be interpreted as two rows, or as two columns"
    Const Err_BadColumnIdentifier As String = "Column identifiers in the left column (or top row) of " & _
        "ConvertTypes must be strings or non-negative whole numbers"
    Const Err_BadCT As String = "Type Conversion given in bottom row (or right column) of ConvertTypes must be " & _
        "Booleans or strings containing letters NDBETQ"
    Const Err_ConvertTypes As String = "ConvertTypes must be a Boolean, a string with allowed letters ""NDBETQ"" or an array"
    Const Err_HeaderRowNum As String = "ConvertTypes specifies columns by their header (instead of by number), " & _
        "but HeaderRowNum has not been specified"
    
    Dim ColIdentifier As Variant
    Dim CT As Variant
    Dim i As Long
    Dim LCN As Long 'Left column number
    Dim NC As Long 'Number of columns
    Dim ND As Long 'Number of dimensions
    Dim NR As Long 'Number of rows
    Dim RCN As Long 'Right Column Number
    Dim Transposed As Boolean
    
    On Error GoTo ErrHandler
    If VarType(ConvertTypes) = vbString Or VarType(ConvertTypes) = vbBoolean Or IsEmpty(ConvertTypes) Then
        ParseCTString CStr(ConvertTypes), ShowNumbersAsNumbers, ShowDatesAsDates, ShowBooleansAsBooleans, _
            ShowErrorsAsErrors, ConvertQuoted, TrimFields
        ColByColFormatting = False
        Exit Sub
    End If

    If TypeName(ConvertTypes) = "Range" Then ConvertTypes = ConvertTypes.Value2
    ND = NumDimensions(ConvertTypes)
    If ND = 1 Then
        ConvertTypes = OneDArrayToTwoDArray(ConvertTypes)
    ElseIf ND = 2 Then
        If LBound(ConvertTypes, 1) <> 1 Or LBound(ConvertTypes, 2) <> 1 Then
            Throw Err_2D
        End If
    End If

    NR = NRows(ConvertTypes)
    NC = NCols(ConvertTypes)
    If NR = 2 And NC = 2 Then
        'Tricky - have we been given two rows or two columns?
        If Not IsCTValid(ConvertTypes(2, 2)) Then Throw Err_ConvertTypes
        If IsCTValid(ConvertTypes(1, 2)) And IsCTValid(ConvertTypes(2, 1)) Then
            If StandardiseCT(ConvertTypes(1, 2)) <> StandardiseCT(ConvertTypes(2, 1)) Then
                Throw Err_Ambiguous
            End If
        End If
        If IsCTValid(ConvertTypes(2, 1)) Then
            ConvertTypes = Transpose(ConvertTypes)
            Transposed = True
        End If
    ElseIf NR = 2 Then
        ConvertTypes = Transpose(ConvertTypes)
        Transposed = True
        NR = NC
    ElseIf NC <> 2 Then
        Throw Err_ConvertTypes
    End If
    LCN = LBound(ConvertTypes, 2)
    RCN = LCN + 1
    For i = LBound(ConvertTypes, 1) To UBound(ConvertTypes, 1)
        ColIdentifier = ConvertTypes(i, LCN)
        CT = ConvertTypes(i, RCN)
        If IsNumber(ColIdentifier) Then
            If ColIdentifier <> CLng(ColIdentifier) Then
                Throw Err_BadColumnIdentifier & _
                    " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                    ") is " & CStr(ColIdentifier)
            ElseIf ColIdentifier < 0 Then
                Throw Err_BadColumnIdentifier & " but ConvertTypes(" & _
                    IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & ") is " & CStr(ColIdentifier)
            End If
        ElseIf VarType(ColIdentifier) <> vbString Then
            Throw Err_BadColumnIdentifier & " but ConvertTypes(" & IIf(Transposed, "1," & CStr(i), CStr(i) & ",1") & _
                ") is of type " & TypeName(ColIdentifier)
        End If
        If Not IsCTValid(CT) Then
            If VarType(CT) = vbString Then
                Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                    ") is string """ & CStr(CT) & """"
            Else
                Throw Err_BadCT & " but ConvertTypes(" & IIf(Transposed, "2," & CStr(i), CStr(i) & ",2") & _
                    ") is of type " & TypeName(CT)
            End If
        End If

        If CTDict.Exists(ColIdentifier) Then
            If Not CTsEqual(CTDict.item(ColIdentifier), CT) Then
                Throw "ConvertTypes is contradictory. Column " & CStr(ColIdentifier) & _
                    " is specified to be converted using two different conversion rules: " & CStr(CT) & _
                    " and " & CStr(CTDict.item(ColIdentifier))
            End If
        Else
            CT = StandardiseCT(CT)
            'Need this line to ensure that we parse the DateFormat provided when doing Col-by-col type conversion
            If InStr(CT, "D") > 0 Then ShowDatesAsDates = True
            If VarType(ColIdentifier) = vbString Then
                If HeaderRowNum = 0 Then
                    Throw Err_HeaderRowNum
                End If
            End If
            CTDict.Add ColIdentifier, CT
        End If
    Next i
    ColByColFormatting = True
    Exit Sub
ErrHandler:
    Throw "#ParseConvertTypes: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCTString
' Purpose    : Parse the input ConvertTypes to set seven Boolean flags which are passed by reference.
' Parameters :
'  ConvertTypes          : The argument to CSVRead
'  ShowNumbersAsNumbers  : Should fields in the file that look like numbers be returned as Numbers? (Doubles)
'  ShowDatesAsDates      : Should fields in the file that look like dates with the specified DateFormat be returned as
'                          Dates?
'  ShowBooleansAsBooleans: Should fields in the file that match one of the TrueStrings or FalseStrings be returned as
'                          Booleans?
'  ShowErrorsAsErrors    : Should fields in the file that look like Excel errors (#N/A #REF! etc) be returned as errors?
'  ConvertQuoted         : Should the four conversion rules above apply even to quoted fields?
'  TrimFields            : Should leading and trailing spaces be trimmed from fields?
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseCTString(ByVal ConvertTypes As String, ByRef ShowNumbersAsNumbers As Boolean, _
    ByRef ShowDatesAsDates As Boolean, ByRef ShowBooleansAsBooleans As Boolean, _
    ByRef ShowErrorsAsErrors As Boolean, ByRef ConvertQuoted As Boolean, ByRef TrimFields As Boolean)

    Const Err_ConvertTypes As String = "ConvertTypes must be Boolean or string with allowed letters NDBETQ. " & _
        """N"" show numbers as numbers, ""D"" show dates as dates, ""B"" show Booleans " & _
        "as Booleans, ""E"" show Excel errors as errors, ""T"" to trim leading and trailing " & _
        "spaces from fields, ""Q"" rules NDBE apply even to quoted fields, TRUE = ""NDB"" " & _
        "(convert unquoted numbers, dates and Booleans), FALSE = no conversion"
    Const Err_Quoted As String = "ConvertTypes is incorrect, ""Q"" indicates that conversion should apply even to " & _
        "quoted fields, but none of ""N"", ""D"", ""B"" or ""E"" are present to indicate which type conversion to apply"
    Dim i As Long

    On Error GoTo ErrHandler

    If ConvertTypes = "True" Or ConvertTypes = "False" Then
        ConvertTypes = StandardiseCT(CBool(ConvertTypes))
    End If

    ShowNumbersAsNumbers = False
    ShowDatesAsDates = False
    ShowBooleansAsBooleans = False
    ShowErrorsAsErrors = False
    ConvertQuoted = False
    For i = 1 To Len(ConvertTypes)
        'Adding another letter? Also change method IsCTValid!
        Select Case UCase$(Mid$(ConvertTypes, i, 1))
            Case "N"
                ShowNumbersAsNumbers = True
            Case "D"
                ShowDatesAsDates = True
            Case "B"
                ShowBooleansAsBooleans = True
            Case "E"
                ShowErrorsAsErrors = True
            Case "Q"
                ConvertQuoted = True
            Case "T"
                TrimFields = True
            Case Else
                Throw Err_ConvertTypes & " Found unrecognised character '" _
                    & Mid$(ConvertTypes, i, 1) & "'"
        End Select
    Next i
    
    If ConvertQuoted And Not (ShowNumbersAsNumbers Or ShowDatesAsDates Or _
        ShowBooleansAsBooleans Or ShowErrorsAsErrors) Then
        Throw Err_Quoted
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseCTString: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Min4
' Purpose    : Returns the minimum of four inputs and an indicator of which of the four was the minimum
' -----------------------------------------------------------------------------------------------------------------------
Private Function Min4(N1 As Long, N2 As Long, N3 As Long, _
     N4 As Long, ByRef Which As Long) As Long

    If N1 < N2 Then
        Min4 = N1
        Which = 1
    Else
        Min4 = N2
        Which = 2
    End If

    If N3 < Min4 Then
        Min4 = N3
        Which = 3
    End If

    If N4 < Min4 Then
        Min4 = N4
        Which = 4
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DetectEncoding
' Purpose    : Guesses whether a file needs to be opened with the "format" argument to File.OpenAsTextStream set to
'              TriStateTrue or TriStateFalse.
'              The documentation at
'              https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/openastextstream-method
'              is limited but I believe that:
'            * TriStateTrue needs to passed for files which (as reported by NotePad++) are encoded as either
'              "UTF-16 LE BOM" or "UTF-16 BE BOM"
'            * TristateFalse needs to be passed for files encoded as "ANSI"
'            * UTF-8 files are not correctly handled by OpenAsTextStream, instead we use ADODB.Stream, setting CharSet
'              to "UTF-8".
' -----------------------------------------------------------------------------------------------------------------------
Private Sub DetectEncoding(FilePath As String, ByRef TriState As Long, _
    ByRef CharSet As String, ByRef useADODB As Boolean)

    Dim intAsc1Chr As Long
    Dim intAsc2Chr As Long
    Dim intAsc3Chr As Long
    Dim t As Scripting.TextStream

    On Error GoTo ErrHandler
    
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    
    If (m_FSO.FileExists(FilePath) = False) Then
        Throw "File not found!"
    End If

    ' 1=Read-only, False=do not create if not exist, -1=Unicode 0=ASCII
    Set t = m_FSO.OpenTextFile(FilePath, 1, False, 0)
    If t.atEndOfStream Then
        t.Close: Set t = Nothing
        TriState = TristateFalse
        CharSet = "_autodetect_all"
        useADODB = False
        Exit Sub
    End If
    intAsc1Chr = Asc(t.Read(1))
    If t.atEndOfStream Then
        t.Close: Set t = Nothing
        TriState = TristateFalse
        CharSet = "_autodetect_all"
        useADODB = False
        Exit Sub
    End If
    
    intAsc2Chr = Asc(t.Read(1))
    
    If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
        'File is probably encoded UTF-16 LE BOM (little endian, with Byte Option Marker)
        TriState = TristateTrue
        CharSet = "utf-16"
        useADODB = False
    ElseIf (intAsc1Chr = 254) And (intAsc2Chr = 255) Then
        'File is probably encoded UTF-16 BE BOM (big endian, with Byte Option Marker)
        TriState = TristateTrue
        CharSet = "utf-16"
        useADODB = False
    Else
        If t.atEndOfStream Then
            TriState = TristateFalse
            Exit Sub
        End If
        intAsc3Chr = Asc(t.Read(1))
        If (intAsc1Chr = 239) And (intAsc2Chr = 187) And (intAsc3Chr = 191) Then
            'File is probably encoded UTF-8 with BOM
            CharSet = "utf-8"
            TriState = TristateFalse
            useADODB = True
        Else
            'We don't know, assume ANSI but that may be incorrect.
            CharSet = vbNullString
            TriState = TristateFalse
            useADODB = False
        End If
    End If

    t.Close: Set t = Nothing
    Exit Sub
ErrHandler:
    Throw "#DetectEncoding: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InferDelimiter
' Purpose    : Infer the delimiter in a file by looking for first occurrence outside quoted regions of comma, tab,
'              semi-colon, colon or pipe (|). Only look in the first 10,000 characters, Would prefer to look at the first
'              10 lines, but that presents a problem for files with Mac line endings as T.ReadLine doesn't work for them.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InferDelimiter(st As enmSourceType, FileNameOrContents As String, _
     TriState As Long, DecimalSeparator As String) As String
    
    Const CHUNK_SIZE As Long = 1000
    Const Err_SourceType As String = "Cannot infer delimiter directly from URL"
    Const MAX_CHUNKS As Long = 10
    Const QuoteChar As String = """"
    Dim Contents As String
    Dim CopyOfErr As String
    Dim EvenQuotes As Boolean
    Dim F As Scripting.File
    Dim i As Long
    Dim j As Long
    Dim MaxChars As Long
    Dim t As TextStream

    On Error GoTo ErrHandler

    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject

    EvenQuotes = True
    If st = st_File Then

        Set F = m_FSO.GetFile(FileNameOrContents)
        Set t = F.OpenAsTextStream(ForReading, TriState)

        If t.atEndOfStream Then
            t.Close: Set t = Nothing: Set F = Nothing
            Throw "File is empty"
        End If

        Do While Not t.atEndOfStream And j <= MAX_CHUNKS
            j = j + 1
            Contents = t.Read(CHUNK_SIZE)
            For i = 1 To Len(Contents)
                Select Case Mid$(Contents, i, 1)
                    Case QuoteChar
                        EvenQuotes = Not EvenQuotes
                    Case ",", vbTab, "|", ";", ":"
                        If EvenQuotes Then
                            If Mid$(Contents, i, 1) <> DecimalSeparator Then
                                InferDelimiter = Mid$(Contents, i, 1)
                                t.Close: Set t = Nothing: Set F = Nothing
                                Exit Function
                            End If
                        End If
                End Select
            Next i
        Loop
        t.Close
    ElseIf st = st_String Then
        Contents = FileNameOrContents
        MaxChars = MAX_CHUNKS * CHUNK_SIZE
        If MaxChars > Len(Contents) Then MaxChars = Len(Contents)

        For i = 1 To MaxChars
            Select Case Mid$(Contents, i, 1)
                Case QuoteChar
                    EvenQuotes = Not EvenQuotes
                Case ",", vbTab, "|", ";", ":"
                    If EvenQuotes Then
                        If Mid$(Contents, i, 1) <> DecimalSeparator Then
                            InferDelimiter = Mid$(Contents, i, 1)
                            Exit Function
                        End If
                    End If
            End Select
        Next i
    Else
        Throw Err_SourceType
    End If

    'No commonly-used delimiter found in the file outside quoted regions _
     and in the first MAX_CHUNKS * CHUNK_SIZE characters. Assume comma _
     unless that's the decimal separator.
    
    If DecimalSeparator = "," Then
        InferDelimiter = ";"
    Else
        InferDelimiter = ","
    End If

    Exit Function
ErrHandler:
    CopyOfErr = "#InferDelimiter: " & Err.Description & "!"
    If Not t Is Nothing Then
        t.Close
        Set t = Nothing: Set F = Nothing
    End If
    Throw CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDateFormat
' Purpose    : Populate DateOrder and DateSeparator by parsing DateFormat.
' Parameters :
'  DateFormat   : String such as D/M/Y or Y-M-D
'  DateOrder    : ByRef argument is set to DateFormat using same convention as Application.International(xlDateOrder)
'                 (0 = MDY, 1 = DMY, 2 = YMD)
'  DateSeparator: ByRef argument is set to the DateSeparator, typically "-" or "/"
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ParseDateFormat(ByVal DateFormat As String, ByRef DateOrder As Long, ByRef DateSeparator As String, _
    ByRef ISO8601 As Boolean, ByRef AcceptWithoutTimeZone As Boolean, ByRef AcceptWithTimeZone As Boolean)

    Dim Err_DateFormat As String

    On Error GoTo ErrHandler
    
    If UCase$(DateFormat) = "ISO" Then
        ISO8601 = True
        AcceptWithoutTimeZone = True
        AcceptWithTimeZone = False
        Exit Sub
    ElseIf UCase$(DateFormat) = "ISOZ" Then
        ISO8601 = True
        AcceptWithoutTimeZone = False
        AcceptWithTimeZone = True
        Exit Sub
    End If
    
    Err_DateFormat = "DateFormat not valid should be one of 'ISO', 'ISOZ', 'M-D-Y', 'D-M-Y', 'Y-M-D', " & _
        "'M/D/Y', 'D/M/Y' or 'Y/M/D'" & ". Omit to use the default date format of 'Y-M-D'"
     
    'Replace repeated D's with a single D, etc since CastToDate only needs _
     to know the order in which the three parts of the date appear.
    If Len(DateFormat) > 5 Then
        DateFormat = UCase$(DateFormat)
        ReplaceRepeats DateFormat, "D"
        ReplaceRepeats DateFormat, "M"
        ReplaceRepeats DateFormat, "Y"
    End If
       
    If Len(DateFormat) = 0 Then 'use "Y-M-D"
        DateOrder = 2
        DateSeparator = "-"
    ElseIf Len(DateFormat) <> 5 Then
        Throw Err_DateFormat
    ElseIf Mid$(DateFormat, 2, 1) <> Mid$(DateFormat, 4, 1) Then
        Throw Err_DateFormat
    Else
        DateSeparator = Mid$(DateFormat, 2, 1)
        If DateSeparator <> "/" And DateSeparator <> "-" Then Throw Err_DateFormat
        Select Case UCase$(Left$(DateFormat, 1) & Mid$(DateFormat, 3, 1) & Right$(DateFormat, 1))
            Case "MDY"
                DateOrder = 0
            Case "DMY"
                DateOrder = 1
            Case "YMD"
                DateOrder = 2
            Case Else
                Throw Err_DateFormat
        End Select
    End If

    Exit Sub
ErrHandler:
    Throw "#ParseDateFormat: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReplaceRepeats
' Purpose    : Replace repeated instances of a character in a string with a single instance.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReplaceRepeats(ByRef TheString As String, TheChar As String)
    Dim ChCh As String
    ChCh = TheChar & TheChar
    Do While InStr(TheString, ChCh) > 0
        TheString = Replace(TheString, ChCh, TheChar)
    Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseCSVContents
' Purpose    : Parse the contents of a CSV file. Returns a string Buffer together with arrays which assist splitting
'              Buffer into a two-dimensional array.
' Parameters :
'  ContentsOrStream: The contents of a CSV file as a string, or else a Scripting.TextStream.
'  useADODB        : Pass as True when ContentsOrStream is ADODB.Stream, False when it's Scripting.TextStream,
'                    ignored when it's a string.
'  QuoteChar       : The quote character, usually ascii 34 ("), which allow fields to contain characters that would
'                    otherwise be significant to parsing, such as delimiters or new line characters.
'  Delimiter       : The string that separates fields within each line. Typically a single character, but needn't be.
'  SkipToRow       : Rows in the file prior to SkipToRow are ignored.
'  Comment         : Lines in the file that start with these characters will be ignored, handled by method SkipLines.
'  IgnoreRepeated  : If true then parsing ignores delimiters at the start of lines, consecutive delimiters and delimiters
'                    at the end of lines.
'  SkipToRow       : The first line of the file to appear in the return from CSVRead. However, we need to parse earlier
'                    lines to identify where SkipToRow starts in the file - see variable HaveReachedSkipToRow.
'  HeaderRowNum    : The row number of the headers in the file, must be less than or equal to SkipToRow.
'  NumRows         : The number of rows to parse. 0 for all rows from SkipToRow to the end of the file.
'  NumRowsFound    : Set to the number of rows in the file that are on or after SkipToRow.
'  NumColsFound    : Set to the number of columns in the file, i.e. the maximum number of fields in any single line.
'  NumFields       : Set to the number of fields in the file that are on or after SkipToRow.  May be less than
'                    NumRowsFound times NumColsFound if not all lines have the same number of fields.
'  Ragged          : Set to True if not all rows of the file have the same number of fields.
'  Starts          : Set to an array of size at least NumFields. Element k gives the point in Buffer at which the
'                    kth field starts.
'  Lengths         : Set to an array of size at least NumFields. Element k gives the length of the kth field.
'  RowIndexes      : Set to an array of size at least NumFields. Element k gives the row at which the kth field should
'                    appear in the return from CSVRead.
'  ColIndexes      : Set to an array of size at least NumFields. Element k gives the column at which the kth field would
'                    appear in the return from CSVRead under the assumption that argument SkipToCol is 1.
'  QuoteCounts     : Set to an array of size at least NumFields. Element k gives the number of QuoteChars that appear in
'                    the kth field.
'  HeaderRow       : Set equal to the contents of the header row in the file, no type conversion, but quoted fields are
'                    unquoted and leading and trailing spaces are removed.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseCSVContents(ContentsOrStream As Variant, useADODB As Boolean, QuoteChar As String, _
     Delimiter As String, Comment As String, IgnoreEmptyLines As Boolean, _
     IgnoreRepeated As Boolean, SkipToRow As Long, HeaderRowNum As Long, NumRows As Long, _
    ByRef NumRowsFound As Long, ByRef NumColsFound As Long, ByRef NumFields As Long, ByRef Ragged As Boolean, _
    ByRef Starts() As Long, ByRef lengths() As Long, ByRef RowIndexes() As Long, ByRef ColIndexes() As Long, _
    ByRef QuoteCounts() As Long, ByRef HeaderRow As Variant) As String

    Const Err_Delimiter As String = "Delimiter must not be the null string"
    Dim Buffer As String
    Dim BufferUpdatedTo As Long
    Dim ColNum As Long
    Dim DoSkipping As Boolean
    Dim EvenQuotes As Boolean
    Dim HaveReachedSkipToRow As Boolean
    Dim i As Long 'Index to read from Buffer
    Dim j As Long 'Index to write to Starts, Lengths, RowIndexes and ColIndexes
    Dim LComment As Long
    Dim LDlm As Long
    Dim NumRowsInFile As Long
    Dim OrigLen As Long
    Dim PosCR As Long
    Dim PosDL As Long
    Dim PosLF As Long
    Dim PosQC As Long
    Dim QuoteArray() As String
    Dim quoteCount As Long
    Dim RowNum As Long
    Dim SearchFor() As String
    Dim Stream As Object
    Dim Streaming As Boolean
    Dim Tmp As Long
    Dim Which As Long

    On Error GoTo ErrHandler
    On Error GoTo ErrHandler
    HeaderRow = Empty
    
    If VarType(ContentsOrStream) = vbString Then
        Buffer = ContentsOrStream
        Streaming = False
    Else
        Set Stream = ContentsOrStream
        If NumRows = 0 Then
            Buffer = ReadAllFromStream(Stream)
            Streaming = False
        Else
            GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
            Streaming = True
        End If
    End If
       
    LComment = Len(Comment)
    If LComment > 0 Or IgnoreEmptyLines Then
        DoSkipping = True
    End If
       
    If Streaming Then
        ReDim SearchFor(1 To 4)
        SearchFor(1) = Delimiter
        SearchFor(2) = vbLf
        SearchFor(3) = vbCr
        SearchFor(4) = QuoteChar
        ReDim QuoteArray(1 To 1)
        QuoteArray(1) = QuoteChar
    End If

    ReDim Starts(1 To 8): ReDim lengths(1 To 8): ReDim RowIndexes(1 To 8)
    ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
    
    LDlm = Len(Delimiter)
    If LDlm = 0 Then Throw Err_Delimiter 'Avoid infinite loop!
    OrigLen = Len(Buffer)
    If Not Streaming Then
        'Ensure Buffer terminates with vbCrLf
        If Right$(Buffer, 1) <> vbCr And Right$(Buffer, 1) <> vbLf Then
            Buffer = Buffer & vbCrLf
        ElseIf Right$(Buffer, 1) = vbCr Then
            Buffer = Buffer & vbLf
        End If
        BufferUpdatedTo = Len(Buffer)
    End If
    
    i = 0: j = 1
    
    If DoSkipping Then
        SkipLines Streaming, useADODB, Comment, LComment, IgnoreEmptyLines, _
            Stream, Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
    End If
    
    If IgnoreRepeated Then
        'IgnoreRepeated: Handle repeated delimiters at the start of the first line
        Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
            i = i + LDlm
        Loop
    End If
    
    ColNum = 1: RowNum = 1
    EvenQuotes = True
    Starts(1) = i + 1
    If SkipToRow = 1 Then HaveReachedSkipToRow = True

    Do
        If EvenQuotes Then
            If Not Streaming Then
                If PosDL <= i Then PosDL = InStr(i + 1, Buffer, Delimiter): If PosDL = 0 Then PosDL = BufferUpdatedTo + 1
                If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
                If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
                If PosQC <= i Then PosQC = InStr(i + 1, Buffer, QuoteChar): If PosQC = 0 Then PosQC = BufferUpdatedTo + 1
                i = Min4(PosDL, PosLF, PosCR, PosQC, Which)
            Else
                i = SearchInBuffer(SearchFor, i + 1, Stream, useADODB, Delimiter, QuoteChar, Which, Buffer, BufferUpdatedTo)
            End If

            If i >= BufferUpdatedTo + 1 Then
                Exit Do
            End If

            If j + 1 > UBound(Starts) Then
                ReDim Preserve Starts(1 To UBound(Starts) * 2)
                ReDim Preserve lengths(1 To UBound(lengths) * 2)
                ReDim Preserve RowIndexes(1 To UBound(RowIndexes) * 2)
                ReDim Preserve ColIndexes(1 To UBound(ColIndexes) * 2)
                ReDim Preserve QuoteCounts(1 To UBound(QuoteCounts) * 2)
            End If

            Select Case Which
                Case 1
                    'Found Delimiter
                    lengths(j) = i - Starts(j)
                    If IgnoreRepeated Then
                        Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
                            i = i + LDlm
                        Loop
                    End If
                    
                    Starts(j + 1) = i + LDlm
                    ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                    ColNum = ColNum + 1
                    QuoteCounts(j) = quoteCount: quoteCount = 0
                    j = j + 1
                    NumFields = NumFields + 1
                    i = i + LDlm - 1
                Case 2, 3
                    'Found line ending
                    lengths(j) = i - Starts(j)
                    If Which = 3 Then 'Found a vbCr
                        If Mid$(Buffer, i + 1, 1) = vbLf Then
                            'Ending is Windows rather than Mac or Unix.
                            i = i + 1
                        End If
                    End If
                    
                    If DoSkipping Then
                        SkipLines Streaming, useADODB, Comment, LComment, IgnoreEmptyLines, Stream, _
                            Delimiter, Buffer, i, QuoteChar, PosLF, PosCR, BufferUpdatedTo
                    End If
                    
                    If IgnoreRepeated Then
                        'IgnoreRepeated: Handle repeated delimiters at the end of the line, _
                         all but one will have already been handled.
                        If lengths(j) = 0 Then
                            If ColNum > 1 Then
                                j = j - 1
                                ColNum = ColNum - 1
                                NumFields = NumFields - 1
                            End If
                        End If
                        'IgnoreRepeated: handle delimiters at the start of the next line to be parsed
                        Do While Mid$(Buffer, i + LDlm, LDlm) = Delimiter
                            i = i + LDlm
                        Loop
                    End If
                    Starts(j + 1) = i + 1

                    If ColNum > NumColsFound Then
                        If NumColsFound > 0 Then
                            Ragged = True
                        End If
                        NumColsFound = ColNum
                    ElseIf ColNum < NumColsFound Then
                        Ragged = True
                    End If
                    
                    ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                    QuoteCounts(j) = quoteCount: quoteCount = 0
                    
                    If RowNum = 1 Then
                        If SkipToRow = 1 Then
                            If HeaderRowNum = 1 Then
                                HeaderRow = GetLastParsedRow(Buffer, Starts, lengths, _
                                     ColIndexes, QuoteCounts, j)
                            End If
                        End If
                    End If
                    If Not HaveReachedSkipToRow Then
                        If RowNum = HeaderRowNum Then
                            HeaderRow = GetLastParsedRow(Buffer, Starts, lengths, _
                                 ColIndexes, QuoteCounts, j)
                        End If
                    End If
                    
                    ColNum = 1: RowNum = RowNum + 1
                    
                    j = j + 1
                    NumFields = NumFields + 1
                    
                    If HaveReachedSkipToRow Then
                        If RowNum = NumRows + 1 Then
                            Exit Do
                        End If
                    Else
                        If RowNum = SkipToRow Then
                            HaveReachedSkipToRow = True
                            Tmp = Starts(j)
                            ReDim Starts(1 To 8): ReDim lengths(1 To 8): ReDim RowIndexes(1 To 8)
                            ReDim ColIndexes(1 To 8): ReDim QuoteCounts(1 To 8)
                            RowNum = 1: j = 1: NumFields = 0
                            Starts(1) = Tmp
                        End If
                    End If
                Case 4
                    'Found QuoteChar
                    EvenQuotes = False
                    quoteCount = quoteCount + 1
            End Select
        Else
            If Not Streaming Then
                PosQC = InStr(i + 1, Buffer, QuoteChar)
            Else
                If PosQC <= i Then PosQC = SearchInBuffer(QuoteArray, i + 1, Stream, useADODB, _
                    Delimiter, QuoteChar, 0, Buffer, BufferUpdatedTo)
            End If
            
            If PosQC = 0 Then
                'Malformed Buffer (not RFC4180 compliant). There should always be an even number of double quotes. _
                 If there are an odd number then all text after the last double quote in the file will be (part of) _
                 the last field in the last line.
                lengths(j) = OrigLen - Starts(j) + 1
                ColIndexes(j) = ColNum: RowIndexes(j) = RowNum
                
                RowNum = RowNum + 1
                If ColNum > NumColsFound Then NumColsFound = ColNum
                NumFields = NumFields + 1
                Exit Do
            Else
                i = PosQC
                EvenQuotes = True
                quoteCount = quoteCount + 1
            End If
        End If
    Loop

    NumRowsFound = RowNum - 1
    
    If HaveReachedSkipToRow Then
        NumRowsInFile = SkipToRow - 1 + RowNum - 1
    Else
        NumRowsInFile = RowNum - 1
    End If
    
    If SkipToRow > NumRowsInFile Then
        If NumRows = 0 Then 'Attempting to read from SkipToRow to the end of the file, but that would be zero or _
                             a negative number of rows. So throw an error.
            Dim RowDescription As String
            If IgnoreEmptyLines And Len(Comment) > 0 Then
                RowDescription = "not commented, not empty "
            ElseIf IgnoreEmptyLines Then
                RowDescription = "not empty "
            ElseIf Len(Comment) > 0 Then
                RowDescription = "not commented "
            End If
                             
            Throw "SkipToRow (" & CStr(SkipToRow) & ") exceeds the number of " & RowDescription & _
                "rows in the file (" & CStr(NumRowsInFile) & ")"
        Else
            'Attempting to read a set number of rows, function CSVRead will return an array of Empty values.
            NumFields = 0
            NumRowsFound = 0
        End If
    End If

    ParseCSVContents = Buffer

    Exit Function
ErrHandler:
    Throw "#ParseCSVContents: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLastParsedRow
' Purpose    : For use during parsing (fn ParseCSVContents) to grab the header row (which may or may not be part of the
'              function return). The argument j should point into the Starts, Lengths etc arrays, pointing to the last
'              field in the header row
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLastParsedRow(Buffer As String, Starts() As Long, lengths() As Long, _
    ColIndexes() As Long, QuoteCounts() As Long, j As Long) As Variant
    Dim NC As Long

    Dim Field As String
    Dim i As Long
    Dim res() As String

    On Error GoTo ErrHandler
    NC = ColIndexes(j)

    ReDim res(1 To 1, 1 To NC)
    For i = j To j - NC + 1 Step -1
        Field = Mid$(Buffer, Starts(i), lengths(i))
        res(1, NC + i - j) = Unquote(Trim$(Field), """", QuoteCounts(i))
    Next i
    GetLastParsedRow = res

    Exit Function
ErrHandler:
    Throw "#GetLastParsedRow: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SkipLines
' Purpose    : Sub-routine of ParseCSVContents. Skip a commented or empty row by incrementing i to the position of
'              the line feed just before the next not-commented line.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub SkipLines(Streaming As Boolean, useADODB As Boolean, Comment As String, _
    LComment As Long, IgnoreEmptyLines As Boolean, Stream As Object, ByVal Delimiter As String, _
    ByRef Buffer As String, ByRef i As Long, QuoteChar As String, ByVal PosLF As Long, ByVal PosCR As Long, _
    ByRef BufferUpdatedTo As Long)
    
    Dim atEndOfStream As Boolean
    Dim LookAheadBy As Long
    Dim SkipThisLine As Boolean
    On Error GoTo ErrHandler
    Do
        If Streaming Then
            LookAheadBy = MaxLngs(LComment, 2)
            If i + LookAheadBy > BufferUpdatedTo Then
                If useADODB Then
                    atEndOfStream = Stream.EOS
                Else
                    atEndOfStream = Stream.atEndOfStream
                End If
                If Not atEndOfStream Then
                    GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
                End If
            End If
        End If

        SkipThisLine = False
        If LComment > 0 Then
            If Mid$(Buffer, i + 1, LComment) = Comment Then
                SkipThisLine = True
            End If
        End If
        If IgnoreEmptyLines Then
            Select Case Mid$(Buffer, i + 1, 1)
                Case vbLf, vbCr
                    SkipThisLine = True
            End Select
        End If

        If SkipThisLine Then
            If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
            If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
            If PosLF < PosCR Then
                i = PosLF
            ElseIf PosLF = PosCR + 1 Then
                i = PosLF
            Else
                i = PosCR
            End If
        Else
            Exit Do
        End If
    Loop

    Exit Sub
ErrHandler:
    Throw "#SkipLines: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SearchInBuffer
' Purpose    : Sub-routine of ParseCSVContents. Returns the location in the buffer of the first-encountered string
'              amongst the elements of SearchFor, starting the search at point SearchFrom and finishing the search at
'              point BufferUpdatedTo. If none found in that region returns BufferUpdatedTo + 1. Otherwise returns the
'              location of the first found and sets the by-reference argument Which to indicate which element of
'              SearchFor was the first to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SearchInBuffer(SearchFor() As String, StartingAt As Long, Stream As Object, _
    useADODB As Boolean, Delimiter As String, QuoteChar As String, ByRef Which As Long, _
    ByRef Buffer As String, ByRef BufferUpdatedTo As Long) As Long

    Dim atEndOfStream As Boolean
    Dim InstrRes As Long
    Dim PrevBufferUpdatedTo As Long

    On Error GoTo ErrHandler

    'in this call only search as far as BufferUpdatedTo
    InstrRes = InStrMulti(SearchFor, Buffer, StartingAt, BufferUpdatedTo, Which)
    If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
        SearchInBuffer = InstrRes
        Exit Function
    Else
        If useADODB Then
            atEndOfStream = Stream.EOS
        Else
            atEndOfStream = Stream.atEndOfStream
        End If
        If atEndOfStream Then
            SearchInBuffer = BufferUpdatedTo + 1
            Exit Function
        End If
    End If

    Do
        PrevBufferUpdatedTo = BufferUpdatedTo
        GetMoreFromStream Stream, Delimiter, QuoteChar, Buffer, BufferUpdatedTo
        InstrRes = InStrMulti(SearchFor, Buffer, PrevBufferUpdatedTo + 1, BufferUpdatedTo, Which)
        If (InstrRes > 0 And InstrRes <= BufferUpdatedTo) Then
            SearchInBuffer = InstrRes
            Exit Function
        ElseIf Stream.EOS Then
            SearchInBuffer = BufferUpdatedTo + 1
            Exit Function
        End If
    Loop
    Exit Function

    Exit Function
ErrHandler:
    Throw "#SearchInBuffer: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InStrMulti
' Purpose    : Sub-routine of ParseCSVContents. Returns the first point in SearchWithin at which one of the elements of
'              SearchFor is found, search is restricted to region [StartingAt, EndingAt] and Which is updated with the
'              index identifying which was the first of the strings to be found.
' -----------------------------------------------------------------------------------------------------------------------
Private Function InStrMulti(SearchFor() As String, SearchWithin As String, StartingAt As Long, _
     EndingAt As Long, ByRef Which As Long) As Long

    Const Inf As Long = 2147483647
    Dim i As Long
    Dim InstrRes() As Long
    Dim LB As Long
    Dim Result As Long
    Dim UB As Long

    On Error GoTo ErrHandler
    LB = LBound(SearchFor): UB = UBound(SearchFor)

    Result = Inf

    ReDim InstrRes(LB To UB)
    For i = LB To UB
        InstrRes(i) = InStr(StartingAt, SearchWithin, SearchFor(i))
        If InstrRes(i) > 0 Then
            If InstrRes(i) <= EndingAt Then
                If InstrRes(i) < Result Then
                    Result = InstrRes(i)
                    Which = i
                End If
            End If
        End If
    Next
    InStrMulti = IIf(Result = Inf, 0, Result)

    Exit Function
ErrHandler:
    Throw "#InStrMulti: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetMoreFromStream, Sub-routine of ParseCSVContents
' Purpose    : Write CHUNKSIZE characters from the TextStream T into the buffer, modifying the passed-by-reference
'              arguments  Buffer, BufferUpdatedTo and Streaming.
'              Complexities:
'           a) We have to be careful not to update the buffer to a point part-way through a two-character end-of-line
'              or a multi-character delimiter, otherwise calling method SearchInBuffer might give the wrong result.
'           b) We update a few characters of the buffer beyond the BufferUpdatedTo point with the delimiter, the
'              QuoteChar and vbCrLf. This ensures that the calls to Instr that search the buffer for these strings do
'              not needlessly scan the unupdated part of the buffer.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub GetMoreFromStream(t As Variant, Delimiter As String, QuoteChar As String, _
ByRef Buffer As String, ByRef BufferUpdatedTo As Long)

    Const ChunkSize As Long = 5000  ' The number of characters to read from the stream on each call. _
                              Set to a small number for testing logic and a bigger number for _
                              performance, but not too high since a common use case is reading _
                              just the first line of a file. Suggest 5000? Note that when reading _
                              an entire file (NumRows argument to CSVRead is zero) function _
                              GetMoreFromStream is not called.
    Dim atEndOfStream As Boolean
    Dim ExpandBufferBy As Long
    Dim FirstPass As Boolean
    Dim i As Long
    Dim IsScripting As Boolean
    Dim NCharsToWriteToBuffer As Long
    Dim NewChars As String
    Dim OKToExit As Boolean

    On Error GoTo ErrHandler
    
    Select Case TypeName(t)
        Case "TextStream"
            IsScripting = True
        Case "Stream"
            IsScripting = False
        Case Else
            Throw "T must be of type Scripting.TextStream or ADODB.Stream"
    End Select
    
    FirstPass = True
    Do
        If IsScripting Then
            NewChars = t.Read(IIf(FirstPass, ChunkSize, 1))
            atEndOfStream = t.atEndOfStream
        Else
            NewChars = t.ReadText(IIf(FirstPass, ChunkSize, 1))
            atEndOfStream = t.EOS
        End If
        FirstPass = False
        If atEndOfStream Then
            'Ensure NewChars terminates with vbCrLf
            If Right$(NewChars, 1) <> vbCr And Right$(NewChars, 1) <> vbLf Then
                NewChars = NewChars & vbCrLf
            ElseIf Right$(NewChars, 1) = vbCr Then
                NewChars = NewChars & vbLf
            End If
        End If

        NCharsToWriteToBuffer = Len(NewChars) + Len(Delimiter) + 3

        If BufferUpdatedTo + NCharsToWriteToBuffer > Len(Buffer) Then
            ExpandBufferBy = MaxLngs(Len(Buffer), NCharsToWriteToBuffer)
            Buffer = Buffer & String(ExpandBufferBy, "?")
        End If
        
        Mid$(Buffer, BufferUpdatedTo + 1, Len(NewChars)) = NewChars
        BufferUpdatedTo = BufferUpdatedTo + Len(NewChars)

        OKToExit = True
        'Ensure we don't leave the buffer updated to part way through a two-character end of line marker.
        If Right$(NewChars, 1) = vbCr Then
            OKToExit = False
        End If
        'Ensure we don't leave the buffer updated to a point part-way through a multi-character delimiter
        If Len(Delimiter) > 1 Then
            For i = 1 To Len(Delimiter) - 1
                If Mid$(Buffer, BufferUpdatedTo - i + 1, i) = Left$(Delimiter, i) Then
                    OKToExit = False
                    Exit For
                End If
            Next i
            If Mid$(Buffer, BufferUpdatedTo - Len(Delimiter) + 1, Len(Delimiter)) = Delimiter Then
                OKToExit = True
            End If
        End If
        If OKToExit Then Exit Do
    Loop

    'Line below arranges that when calling Instr(Buffer,....) we don't pointlessly scan the space characters _
     we can be sure that there is space in the buffer to write the extra characters thanks to
    Mid$(Buffer, BufferUpdatedTo + 1, 2 + Len(QuoteChar) + Len(Delimiter)) = vbCrLf & QuoteChar & Delimiter

    Exit Sub
ErrHandler:
    Throw "#GetMoreFromStream: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CountQuotes
' Purpose    : Count the quotes in a string, only used when applying column-by-column type conversion, because in that
'              case it's not possible to use the count of quotes made at parsing time which is organised row-by-row.
' -----------------------------------------------------------------------------------------------------------------------
Private Function CountQuotes(Str As String, QuoteChar As String) As Long
    Dim N As Long
    Dim pos As Long

    Do
        pos = InStr(pos + 1, Str, QuoteChar)
        If pos = 0 Then
            CountQuotes = N
            Exit Function
        End If
        N = N + 1
    Loop
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConvertField
' Purpose    : Convert a field in the file into an element of the returned array.
' Parameters :
'General
'  Field                : Field, i.e. characters from the file between successive delimiters.
'  AnyConversion        : Is any type conversion to take place? i.e. processing other than trimming whitespace and
'                         unquoting.
'  FieldLength          : The length of Field.
'Whitespace and Quotes
'  TrimFields           : Should leading and trailing spaces be trimmed from fields? For quoted fields, this will not
'                         remove spaces between the quotes.
'  QuoteChar            : The quote character, typically ". No support for different opening and closing quote characters
'                         or different escape character.
'  QuoteCount           : How many quote characters does Field contain?
'  ConvertQuoted        : Should quoted fields (after quote removal) be converted according to arguments
'                         ShowNumbersAsNumbers, ShowDatesAsDates, and the contents of Sentinels.
'Numbers
'  ShowNumbersAsNumbers : If Field is a string representation of a number should the function return that number?
'  SepStandard          : Is the decimal separator the same as the system defaults? If True then the next two arguments
'                         are ignored.
'  DecimalSeparator     : The decimal separator used in Field.
'  SysDecimalSeparator  : The default decimal separator on the system.
'Dates
'  ShowDatesAsDates     : If Field is a string representation of a date should the function return that date?
'  ISO8601              : If Field is a date, does it respect (a subset of) ISO8601?
'  AcceptWithoutTimeZone: In the case of ISO8601 dates, should conversion be applied to dates-with-time that have no time
'                         zone information?
'  AcceptWithTimeZone   : In the case of ISO8601 dates, should conversion be applied to dates-with-time that have time
'                         zone information?
'  DateOrder            : If Field is a string representation what order of parts must it respect (not relevant if
'                         ISO8601 is True) 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  DateSeparator        : The date separator, must be either "-" or "/".
'  SysDateOrder         : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D.
'  SysDateSeparator     : The Windows system date separator.
'Booleans, Errors, Missings
'  AnySentinels         : Does the sentinel dictionary have any elements?
'  Sentinels            : A dictionary of Sentinels. If Sentinels.Exists(Field) Then ConvertField = Sentinels(Field)
'  MaxSentinelLength    : The maximum length of the keys of Sentinels.
'  ShowMissingsAs       : The value to which missing fields (consecutive delimiters) are converted. If CSVRead has a
'                         MissingStrings argument then values matching those strings are also converted to
'                         ShowMissingsAs, thanks to method MakeSentinels.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ConvertField(Field As String, AnyConversion As Boolean, FieldLength As Long, _
     TrimFields As Boolean, QuoteChar As String, quoteCount As Long, ConvertQuoted As Boolean, _
     ShowNumbersAsNumbers As Boolean, SepStandard As Boolean, DecimalSeparator As String, _
     SysDecimalSeparator As String, ShowDatesAsDates As Boolean, ISO8601 As Boolean, _
     AcceptWithoutTimeZone As Boolean, AcceptWithTimeZone As Boolean, DateOrder As Long, _
     DateSeparator As String, SysDateOrder As Long, SysDateSeparator As String, _
     AnySentinels As Boolean, Sentinels As Dictionary, MaxSentinelLength As Long, _
     ShowMissingsAs As Variant) As Variant

    Dim Converted As Boolean
    Dim dblResult As Double
    Dim dtResult As Date

    If TrimFields Then
        If Left$(Field, 1) = " " Then
            Field = Trim$(Field)
            FieldLength = Len(Field)
        ElseIf Right$(Field, 1) = " " Then
            Field = Trim$(Field)
            FieldLength = Len(Field)
        End If
    End If

    If FieldLength = 0 Then
        ConvertField = ShowMissingsAs
        Exit Function
    End If

    If Not AnyConversion Then
        If quoteCount = 0 Then
            ConvertField = Field
            Exit Function
        End If
    End If

    If quoteCount > 0 Then
        If Left$(Field, 1) = QuoteChar Then
            If Right$(Field, 1) = QuoteChar Then
                Field = Mid$(Field, 2, FieldLength - 2)
                If quoteCount > 2 Then
                    Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar)
                End If
                If ConvertQuoted Then
                    FieldLength = Len(Field)
                Else
                    ConvertField = Field
                    Exit Function
                End If
            End If
        End If
    End If

    If AnySentinels Then
        If FieldLength <= MaxSentinelLength Then
            If Sentinels.Exists(Field) Then
                ConvertField = Sentinels.item(Field)
                Exit Function
            End If
        End If
    End If

    If Not ConvertQuoted Then
        If quoteCount > 0 Then
            ConvertField = Field
            Exit Function
        End If
    End If

    If ShowNumbersAsNumbers Then
        CastToDouble Field, dblResult, SepStandard, DecimalSeparator, SysDecimalSeparator, Converted
        If Converted Then
            ConvertField = dblResult
            Exit Function
        End If
    End If

    If ShowDatesAsDates Then
        If ISO8601 Then
            CastISO8601 Field, dtResult, Converted, AcceptWithoutTimeZone, AcceptWithTimeZone
        Else
            CastToDate Field, dtResult, DateOrder, DateSeparator, SysDateOrder, SysDateSeparator, Converted
        End If
        If Not Converted Then
            If InStr(Field, ":") > 0 Then
                CastToTime Field, dtResult, Converted
                If Not Converted Then
                    CastToTimeB Field, dtResult, Converted
                End If
            End If
        End If
        If Converted Then
            ConvertField = dtResult
            Exit Function
        End If
    End If

    ConvertField = Field
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unquote
' Purpose    : Unquote a field.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Unquote(ByVal Field As String, QuoteChar As String, quoteCount As Long) As String

    On Error GoTo ErrHandler
    If quoteCount > 0 Then
        If Left$(Field, 1) = QuoteChar Then
            If Right$(QuoteChar, 1) = QuoteChar Then
                Field = Mid$(Field, 2, Len(Field) - 2)
                If quoteCount > 2 Then
                    Field = Replace(Field, QuoteChar & QuoteChar, QuoteChar)
                End If
            End If
        End If
    End If
    Unquote = Field

    Exit Function
ErrHandler:
    Throw "#Unquote: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDouble, sub-routine of ConvertField
' Purpose    : Casts strIn to double where strIn has specified decimals separator.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDouble(strIn As String, ByRef dblOut As Double, SepStandard As Boolean, _
     DecimalSeparator As String, SysDecimalSeparator As String, ByRef Converted As Boolean)
    
    On Error GoTo ErrHandler
    If SepStandard Then
        dblOut = CDbl(strIn)
    Else
        dblOut = CDbl(Replace(strIn, DecimalSeparator, SysDecimalSeparator))
    End If
    Converted = True
ErrHandler:
    'Do nothing - strIn was not a string representing a number.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToDate, sub-routine of ConvertField
' Purpose    : In-place conversion of a string that looks like a date into a Long or Date. No error if string cannot be
'              converted to date. Converts Dates, DateTimes and Times. Times in very simple format hh:mm:ss
'              Does not handle ISO8601 - see alternative function CastISO8601
' Parameters :
'  strIn           : String
'  dtOut           : Result of cast
'  DateOrder       : The date order respected by the contents of strIn. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  DateSeparator   : The date separator used by the input
'  SysDateOrder    : The Windows system date order. 0 = M-D-Y, 1= D-M-Y, 2 = Y-M-D
'  SysDateSeparator: The Windows system date separator
'  Converted       : Boolean flipped to TRUE if conversion takes place
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToDate(strIn As String, ByRef dtOut As Date, DateOrder As Long, _
     DateSeparator As String, SysDateOrder As Long, SysDateSeparator As String, _
    ByRef Converted As Boolean)
    
    Dim D As String
    Dim m As String
    Dim pos1 As Long 'First date separator
    Dim pos2 As Long 'Second date separator
    Dim pos3 As Long 'Space to separate date from time
    Dim pos4 As Long 'decimal point for fractions of a second
    Dim Converted2 As Boolean
    Dim HasFractionalSecond As Boolean
    Dim HasTimePart As Boolean
    Dim TimePart As String
    Dim TimePartConverted As Date
    Dim y As String
    
    On Error GoTo ErrHandler
    
    pos1 = InStr(strIn, DateSeparator)
    If pos1 = 0 Then Exit Sub
    pos2 = InStr(pos1 + 1, strIn, DateSeparator)
    If pos2 = 0 Then Exit Sub
    pos3 = InStr(pos2 + 1, strIn, " ")
    
    HasTimePart = pos3 > 0
    
    If Not HasTimePart Then
        If DateOrder = 2 Then 'Y-M-D is unambiguous as long as year given as 4 digits
            If pos1 = 5 Then
                dtOut = CDate(strIn)
                Converted = True
                Exit Sub
            End If
        ElseIf DateOrder = SysDateOrder Then
            dtOut = CDate(strIn)
            Converted = True
            Exit Sub
        End If
        If DateOrder = 0 Then 'M-D-Y
            m = Left$(strIn, pos1 - 1)
            D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
            y = Mid$(strIn, pos2 + 1)
        ElseIf DateOrder = 1 Then 'D-M-Y
            D = Left$(strIn, pos1 - 1)
            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
            y = Mid$(strIn, pos2 + 1)
        ElseIf DateOrder = 2 Then 'Y-M-D
            y = Left$(strIn, pos1 - 1)
            m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
            D = Mid$(strIn, pos2 + 1)
        Else
            Throw "DateOrder must be 0, 1, or 2"
        End If
        If SysDateOrder = 0 Then
            dtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y)
            Converted = True
        ElseIf SysDateOrder = 1 Then
            dtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y)
            Converted = True
        ElseIf SysDateOrder = 2 Then
            dtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D)
            Converted = True
        End If
        Exit Sub
    End If

    pos4 = InStr(pos3 + 1, strIn, ".")
    HasFractionalSecond = pos4 > 0

    If DateOrder = 0 Then 'M-D-Y
        m = Left$(strIn, pos1 - 1)
        D = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
        TimePart = Mid$(strIn, pos3)

    ElseIf DateOrder = 1 Then 'D-M-Y
        D = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        y = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
        TimePart = Mid$(strIn, pos3)
    ElseIf DateOrder = 2 Then 'Y-M-D
        y = Left$(strIn, pos1 - 1)
        m = Mid$(strIn, pos1 + 1, pos2 - pos1 - 1)
        D = Mid$(strIn, pos2 + 1, pos3 - pos2 - 1)
        TimePart = Mid$(strIn, pos3)
    Else
        Throw "DateOrder must be 0, 1, or 2"
    End If
    If Not HasFractionalSecond Then
        If DateOrder = 2 Then 'Y-M-D is unambiguous as long as year given as 4 digits
            If pos1 = 5 Then
                dtOut = CDate(strIn)
                Converted = True
                Exit Sub
            End If
        ElseIf DateOrder = SysDateOrder Then
            dtOut = CDate(strIn)
            Converted = True
            Exit Sub
        End If
    
        If SysDateOrder = 0 Then
            dtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y & TimePart)
            Converted = True
        ElseIf SysDateOrder = 1 Then
            dtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y & TimePart)
            Converted = True
        ElseIf SysDateOrder = 2 Then
            dtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D & TimePart)
            Converted = True
        End If
    Else 'CDate does not cope with fractional seconds, so use CastToTimeB
        CastToTimeB Mid$(TimePart, 2), TimePartConverted, Converted2
        If Converted2 Then
            If SysDateOrder = 0 Then
                dtOut = CDate(m & SysDateSeparator & D & SysDateSeparator & y) + TimePartConverted
                Converted = True
            ElseIf SysDateOrder = 1 Then
                dtOut = CDate(D & SysDateSeparator & m & SysDateSeparator & y) + TimePartConverted
                Converted = True
            ElseIf SysDateOrder = 2 Then
                dtOut = CDate(y & SysDateSeparator & m & SysDateSeparator & D) + TimePartConverted
                Converted = True
            End If
        End If
    End If

    Exit Sub
ErrHandler:
    'Do nothing - was not a string representing a date with the specified date order and date separator.
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NumDimensions(x As Variant) As Long
    Dim i As Long
    Dim y As Long
    If Not IsArray(x) Then
        NumDimensions = 0
        Exit Function
    Else
        On Error GoTo ExitPoint
        i = 1
        Do While True
            y = LBound(x, i)
            i = i + 1
        Loop
    End If
ExitPoint:
    NumDimensions = i - 1
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeSentinels
' Purpose    : Returns a Dictionary keyed on strings for which if a key to the dictionary is a field of the CSV file then
'              that field should be converted to the associated item value. Handles Booleans, Missings and Excel errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub MakeSentinels(ByRef Sentinels As Scripting.Dictionary, ByRef MaxLength As Long, _
    ByRef AnySentinels As Boolean, ShowBooleansAsBooleans As Boolean, ShowErrorsAsErrors As Boolean, _
    ByRef ShowMissingsAs As Variant, Optional TrueStrings As Variant, Optional FalseStrings As Variant, _
    Optional MissingStrings As Variant)

    Const Err_FalseStrings As String = "FalseStrings must be omitted or provided as a string or an array of " & _
        "strings that represent Boolean value False"
    Const Err_MissingStrings As String = "MissingStrings must be omitted or provided a string or an array of " & _
        "strings that represent missing values"
    Const Err_ShowMissings As String = "ShowMissingsAs has an illegal value, such as an array or an object"
    Const Err_TrueStrings As String = "TrueStrings must be omitted or provided as string or an array of " & _
        "strings that represent Boolean value True"
    Const Err_TrueStrings2 As String = "TrueStrings has been provided, but type conversion for Booleans is " & _
        "not switched on for any column"
    Const Err_FalseStrings2 As String = "FalseStrings has been provided, but type conversion for Booleans " & _
        "is not switched on for any column"

    On Error GoTo ErrHandler

    If IsMissing(ShowMissingsAs) Then ShowMissingsAs = Empty
    Select Case VarType(ShowMissingsAs)
        Case vbObject, vbArray, vbByte, vbDataObject, vbUserDefinedType, vbVariant
            Throw Err_ShowMissings
    End Select
    
    If Not IsMissing(MissingStrings) And Not IsEmpty(MissingStrings) Then
        AddKeysToDict Sentinels, MissingStrings, ShowMissingsAs, Err_MissingStrings
    End If

    If ShowBooleansAsBooleans Then
        If IsMissing(TrueStrings) Or IsEmpty(TrueStrings) Then
            AddKeysToDict Sentinels, Array("TRUE", "true", "True"), True, Err_TrueStrings
        Else
            AddKeysToDict Sentinels, TrueStrings, True, Err_TrueStrings
        End If
        If IsMissing(FalseStrings) Or IsEmpty(FalseStrings) Then
            AddKeysToDict Sentinels, Array("FALSE", "false", "False"), False, Err_FalseStrings
        Else
            AddKeysToDict Sentinels, FalseStrings, False, Err_FalseStrings
        End If
    Else
        If Not (IsMissing(TrueStrings) Or IsEmpty(TrueStrings)) Then
            Throw Err_TrueStrings2
        End If
        If Not (IsMissing(FalseStrings) Or IsEmpty(FalseStrings)) Then
            Throw Err_FalseStrings2
        End If
    End If
    
    If ShowErrorsAsErrors Then
        AddKeyToDict Sentinels, "#DIV/0!", CVErr(xlErrDiv0)
        AddKeyToDict Sentinels, "#NAME?", CVErr(xlErrName)
        AddKeyToDict Sentinels, "#REF!", CVErr(xlErrRef)
        AddKeyToDict Sentinels, "#NUM!", CVErr(xlErrNum)
        AddKeyToDict Sentinels, "#NULL!", CVErr(xlErrNull)
        AddKeyToDict Sentinels, "#N/A", CVErr(xlErrNA)
        AddKeyToDict Sentinels, "#VALUE!", CVErr(xlErrValue)
        AddKeyToDict Sentinels, "#SPILL!", CVErr(2045)
        AddKeyToDict Sentinels, "#BLOCKED!", CVErr(2047)
        AddKeyToDict Sentinels, "#CONNECT!", CVErr(2046)
        AddKeyToDict Sentinels, "#UNKNOWN!", CVErr(2048)
        AddKeyToDict Sentinels, "#GETTING_DATA!", CVErr(2043)
        AddKeyToDict Sentinels, "#FIELD!", CVErr(2049)
        AddKeyToDict Sentinels, "#CALC!", CVErr(2050)
    End If

    Dim k As Variant
    MaxLength = 0
    For Each k In Sentinels.Keys
        If Len(k) > MaxLength Then MaxLength = Len(k)
    Next
    AnySentinels = Sentinels.count > 0

    Exit Sub
ErrHandler:
    Throw "#MakeSentinels: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeysToDict, Sub-routine of MakeSentinels
' Purpose    : Broadcast AddKeyToDict over an array of keys.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeysToDict(ByRef Sentinels As Scripting.Dictionary, ByVal Keys As Variant, item As Variant, _
     FriendlyErrorString As String)

    Dim i As Long
    Dim j As Long
  
    On Error GoTo ErrHandler
  
    If TypeName(Keys) = "Range" Then
        Keys = Keys.value
    End If
    
    If VarType(Keys) = vbString Then
        If InStr(Keys, ",") > 0 Then
            Keys = VBA.Split(Keys, ",")
        End If
    End If
    
    Select Case NumDimensions(Keys)
        Case 0
            AddKeyToDict Sentinels, Keys, item, FriendlyErrorString
        Case 1
            For i = LBound(Keys) To UBound(Keys)
                AddKeyToDict Sentinels, Keys(i), item, FriendlyErrorString
            Next i
        Case 2
            For i = LBound(Keys, 1) To UBound(Keys, 1)
                For j = LBound(Keys, 2) To UBound(Keys, 2)
                    AddKeyToDict Sentinels, Keys(i, j), item, FriendlyErrorString
                Next j
            Next i
        Case Else
            Throw FriendlyErrorString
    End Select
    Exit Sub
ErrHandler:
    Throw "#AddKeysToDict: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddKeyToDict, Sub-routine of MakeSentinels
' Purpose    : Wrap .Add method to have more helpful error message if things go awry.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AddKeyToDict(ByRef Sentinels As Scripting.Dictionary, Key As Variant, item As Variant, _
    Optional FriendlyErrorString As String)

    Dim FoundRepeated As Boolean

    On Error GoTo ErrHandler

    If VarType(Key) <> vbString Then Throw FriendlyErrorString & " but '" & CStr(Key) & "' is of type " & TypeName(Key)
    If Len(Key) = 0 Then Exit Sub
    
    If Not Sentinels.Exists(Key) Then
        Sentinels.Add Key, item
    Else
        FoundRepeated = True
        If VarType(item) = VarType(Sentinels.item(Key)) Then
            If item = Sentinels.item(Key) Then
                FoundRepeated = False
            End If
        End If
    End If

    If FoundRepeated Then
        Throw "There is a conflicting definition of what the string '" & Key & _
            "' should be converted to, both the " & TypeName(item) & " value '" & CStr(item) & _
            "' and the " & TypeName(Sentinels.item(Key)) & " value '" & CStr(Sentinels.item(Key)) & _
            "' have been specified. Please check the TrueStrings, FalseStrings and MissingStrings arguments."
    End If

    Exit Sub
ErrHandler:
    Throw "#AddKeyToDict: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseISO8601
' Purpose    : Test harness for calling from spreadsheets
' -----------------------------------------------------------------------------------------------------------------------
Public Function ParseISO8601(strIn As String) As Variant
    Dim Converted As Boolean
    Dim dtOut As Date

    On Error GoTo ErrHandler
    CastISO8601 strIn, dtOut, Converted, True, True

    If Converted Then
        ParseISO8601 = dtOut
    Else
        ParseISO8601 = "#Not recognised as ISO8601 date!"
    End If
    Exit Function
ErrHandler:
    ParseISO8601 = "#ParseISO8601: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTime
' Purpose    : Cast strings that represent a time to a date, no handling of TimeZone.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTime(strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean)

    On Error GoTo ErrHandler
    
    dtOut = CDate(strIn)
    If dtOut <= 1 Then
        Converted = True
    End If
    
    Exit Sub
ErrHandler:
    'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastToTimeB
' Purpose    : CDate does not correctly cope with times such as '04:20:10.123 am' or '04:20:10.123', i.e, times with a
'              fractional second, so this method is called after CastToTime
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastToTimeB(strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean)
    Static rx As VBScript_RegExp_55.RegExp
    Dim DecPointAt As Long
    Dim FractionalSecond As Double
    Dim SpaceAt As Long
    
    On Error GoTo ErrHandler
    If rx Is Nothing Then
        Set rx = New RegExp
        With rx
            .IgnoreCase = True
            .Pattern = "^[0-2]?[0-9]:[0-5]?[0-9]:[0-5]?[0-9](\.[0-9]+)( am| pm)?$"
            .Global = False        'Find first match only
        End With
    End If

    If Not rx.Test(strIn) Then Exit Sub
    DecPointAt = InStr(strIn, ".")
    If DecPointAt = 0 Then Exit Sub ' should never happen
    SpaceAt = InStr(strIn, " ")
    If SpaceAt = 0 Then SpaceAt = Len(strIn) + 1
    FractionalSecond = CDbl(Mid$(strIn, DecPointAt, SpaceAt - DecPointAt)) / 86400
    
    dtOut = CDate(Left$(strIn, DecPointAt - 1) + Mid$(strIn, SpaceAt)) + FractionalSecond
    Converted = True
    Exit Sub
ErrHandler:
    'Do nothing, was not a valid time (e.g. h,m or s out of range)
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CastISO8601
' Purpose    : Convert ISO8601 formatted datestrings to UTC date. https://xkcd.com/1179/

'Always accepts dates without time
'Format                        Example
'yyyy-mm-dd                    2021-08-23

'If AcceptWithoutTimeZone is True:
'yyyy-mm-ddThh:mm:ss           2021-08-23T08:47:21
'yyyy-mm-ddThh:mm:ss.000       2021-08-23T08:47:20.920

'If AcceptWithTimeZone is True:
'yyyy-mm-ddThh:mm:ssZ          2021-08-23T08:47:21Z
'yyyy-mm-ddThh:mm:ss.000Z      2021-08-23T08:47:20.920Z
'yyyy-mm-ddThh:mm:ss+hh:mm     2021-08-23T08:47:21+05:00
'yyyy-mm-ddThh:mm:ss.000+hh:mm 2021-08-23T08:47:20.920+05:00

' Parameters :
'  StrIn                : The string to be converted
'  DtOut                : The date that the string converts to.
'  Converted            : Did the function convert (true) or reject as not a correctly formatted date (false)
'  AcceptWithoutTimeZone: Should the function accept datetime without time zone given?
'  AcceptWithTimeZone   : Should the function accept datetime with time zone given?

'       IMPORTANT:       WHEN TIMEZONE IS GIVEN THE FUNCTION RETURNS THE TIME IN UTC

' -----------------------------------------------------------------------------------------------------------------------
Private Sub CastISO8601(ByVal strIn As String, ByRef dtOut As Date, ByRef Converted As Boolean, _
     AcceptWithoutTimeZone As Boolean, AcceptWithTimeZone As Boolean)

    Dim L As Long
    Dim LocalTime As Double
    Dim MilliPart As Double
    Dim MinusPos As Long
    Dim PlusPos As Long
    Dim Sign As Long
    Dim ZAtEnd As Boolean
    
    Static rxNoNo As VBScript_RegExp_55.RegExp
    Static RxYesNo As VBScript_RegExp_55.RegExp
    Static RxNoYes As VBScript_RegExp_55.RegExp
    Static rxYesYes As VBScript_RegExp_55.RegExp
    Static rxExists As Boolean

    On Error GoTo ErrHandler
    
    If Not rxExists Then
        Set rxNoNo = New RegExp
        'Reject datetime
        With rxNoNo
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9]$"
            .Global = False
        End With
        
        'Accept datetime without time zone, reject datetime with timezone
        Set RxYesNo = New RegExp
        With RxYesNo
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?)?$"
            .Global = False
        End With
        
        'Reject datetime without time zone, accept datetime with timezone
        Set RxNoYes = New RegExp
        With RxNoYes
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?(Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?$"
            .Global = False
        End With
        
        'Accept datetime, both with and without timezone
        Set rxYesYes = New RegExp
        With rxYesYes
            .IgnoreCase = False
            .Pattern = "^[0-9][0-9][0-9][0-9]\-[[0-1][0-9]\-[0-3][0-9](T[0-2][0-9]:[0-5][0-9]:[0-5][0-9](\.[0-9]+)?((Z|((\+|\-)[0-2][0-9]:[0-5][0-9])))?)?$"
            .Global = False
        End With
        rxExists = True
    End If
    
    L = Len(strIn)

    If L = 10 Then
        If rxNoNo.Test(strIn) Then
            'This works irrespective of Windows regional settings
            dtOut = CDate(strIn)
            Converted = True
            Exit Sub
        End If
    ElseIf L < 10 Then
        Converted = False
        Exit Sub
    ElseIf L > 40 Then
        Converted = False
        Exit Sub
    End If

    Converted = False
    
    If AcceptWithoutTimeZone Then
        If AcceptWithTimeZone Then
            If Not rxYesYes.Test(strIn) Then Exit Sub
        Else
            If Not RxYesNo.Test(strIn) Then Exit Sub
        End If
    Else
        If AcceptWithTimeZone Then
            If Not RxNoYes.Test(strIn) Then Exit Sub
        Else
            If Not rxNoNo.Test(strIn) Then Exit Sub
        End If
    End If
    
    'Replace the "T" separator
    Mid$(strIn, 11, 1) = " "
    
    If L = 19 Then
        dtOut = CDate(strIn)
        Converted = True
        Exit Sub
    End If

    If Right$(strIn, 1) = "Z" Then
        Sign = 0
        ZAtEnd = True
    Else
        PlusPos = InStr(20, strIn, "+")
        If PlusPos > 0 Then
            Sign = 1
        Else
            MinusPos = InStr(20, strIn, "-")
            If MinusPos > 0 Then
                Sign = -1
            End If
        End If
    End If

    If Mid$(strIn, 20, 1) = "." Then 'Have fraction of a second
        Select Case Sign
            Case 0
                'Example: "2021-08-23T08:47:20.920Z"
                MilliPart = CDbl(Mid$(strIn, 20, IIf(ZAtEnd, L - 20, L - 19)))
            Case 1
                'Example: "2021-08-23T08:47:20.920+05:00"
                MilliPart = CDbl(Mid$(strIn, 20, PlusPos - 20))
            Case -1
                'Example: "2021-08-23T08:47:20.920-05:00"
                MilliPart = CDbl(Mid$(strIn, 20, MinusPos - 20))
        End Select
    End If
    
    LocalTime = CDate(Left$(strIn, 19)) + MilliPart / 86400

    Dim Adjust As Date
    Select Case Sign
        Case 0
            dtOut = LocalTime
            Converted = True
            Exit Sub
        Case 1
            If L <> PlusPos + 5 Then Exit Sub
            Adjust = CDate(Right$(strIn, 5))
            dtOut = LocalTime - Adjust
            Converted = True
        Case -1
            If L <> MinusPos + 5 Then Exit Sub
            Adjust = CDate(Right$(strIn, 5))
            dtOut = LocalTime + Adjust
            Converted = True
    End Select

    Exit Sub
ErrHandler:
    'Was not recognised as ISO8601 date
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetLocalOffsetToUTC
' Purpose    : Get the PC's offset to UTC.
'See "gogeek"'s post at _
 https://stackoverflow.com/questions/1600875/how-to-get-the-current-datetime-in-utc-from-an-excel-vba-macro
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetLocalOffsetToUTC() As Double
    Dim dt As Object
    Dim TimeNow As Date
    Dim UTC As Date
    On Error GoTo ErrHandler
    TimeNow = Now()

    Set dt = CreateObject("WbemScripting.SWbemDateTime")
    dt.SetVarDate TimeNow
    UTC = dt.GetVarDate(False)
    GetLocalOffsetToUTC = (TimeNow - UTC)

    Exit Function
ErrHandler:
    Throw "#GetLocalOffsetToUTC: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISOZFormatString
' Purpose    : Returns the format string required to save datetimes with timezone under the assumton that the offset to
'              UTC is the same as the curent offset on this PC - use with care, Daylight saving may mean that that's not
'              a correct assumption for all the dates in a set of data...
' -----------------------------------------------------------------------------------------------------------------------
Private Function ISOZFormatString() As String
    Dim RightChars As String
    Dim TimeZone As String

    On Error GoTo ErrHandler
    TimeZone = GetLocalOffsetToUTC()

    If TimeZone = 0 Then
        RightChars = "Z"
    ElseIf TimeZone > 0 Then
        RightChars = "+" & Format$(TimeZone, "hh:mm")
    Else
        RightChars = "-" & Format$(Abs(TimeZone), "hh:mm")
    End If
    ISOZFormatString = "yyyy-mm-ddT:hh:mm:ss" & RightChars

    Exit Function
ErrHandler:
    Throw "#ISOZFormatString: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseTextFile
' Purpose    : Convert a text file to a 2-dim array with one column, one line of file to one element of array, works for
'              files with any style of line endings - Windows, Mac, Unix, or a mixture of line endings.
' Parameters :
'  FileNameOrContents : FileName or CSV-style string.
'  isFile             : If True then fist argument is the name of a file, else it's a CSV-style string.
'  useADODB           : Should the file be read using ADODB.Stream rather than Scripting.TextStream? Necessary only for
'                       UTF-8 files.
'  CharSet            : Used only if useADODB is True and isFile is True.
'  TriState           : Used only if useADODB is False and isFile is True.
'  SkipToLine         : Return starts at this line of the file.
'  NumLinesToReturn   : This many lines are returned. Pass zero for all lines from SkipToLine.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseTextFile(FileNameOrContents As String, isFile As Boolean, useADODB As Boolean, _
     CharSet As String, TriState As Long, SkipToLine As Long, NumLinesToReturn As Long, _
     CallingFromWorksheet As Boolean) As Variant

    Const Err_FileEmpty As String = "File is empty"
    Dim Buffer As String
    Dim BufferUpdatedTo As Long
    Dim FoundCR As Boolean
    Dim HaveReachedSkipToLine As Boolean
    Dim i As Long 'Index to read from Buffer
    Dim j As Long 'Index to write to Starts, Lengths
    Dim Err_StringTooLong As String
    Dim lengths() As Long
    Dim MSLIA As Long
    Dim NumLinesFound As Long
    Dim PosCR As Long
    Dim PosLF As Long
    Dim ReturnArray() As String
    Dim SearchFor() As String
    Dim Starts() As Long
    Dim Stream As Object
    Dim Streaming As Boolean
    Dim Tmp As Long
    Dim Which As Long

    On Error GoTo ErrHandler
    
    If isFile Then
        If useADODB Then
            Set Stream = CreateObject("ADODB.Stream")
            Stream.CharSet = CharSet
            Stream.Open
            Stream.LoadFromFile FileNameOrContents
            If Stream.EOS Then Throw Err_FileEmpty
        Else
            If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
            Set Stream = m_FSO.GetFile(FileNameOrContents).OpenAsTextStream(ForReading, TriState)
            If Stream.atEndOfStream Then Throw Err_FileEmpty
        End If
    
        If NumLinesToReturn = 0 Then
            Buffer = ReadAllFromStream(Stream)
            Streaming = False
        Else
            GetMoreFromStream Stream, vbNullString, vbNullString, Buffer, BufferUpdatedTo
            Streaming = True
        End If
    Else
        Buffer = FileNameOrContents
        Streaming = False
    End If
       
    If Streaming Then
        ReDim SearchFor(1 To 2)
        SearchFor(1) = vbLf
        SearchFor(2) = vbCr
    End If

    ReDim Starts(1 To 8): ReDim lengths(1 To 8)
    
    If Not Streaming Then
        'Ensure Buffer terminates with vbCrLf
        If Right$(Buffer, 1) <> vbCr And Right$(Buffer, 1) <> vbLf Then
            Buffer = Buffer & vbCrLf
        ElseIf Right$(Buffer, 1) = vbCr Then
            Buffer = Buffer & vbLf
        End If
        BufferUpdatedTo = Len(Buffer)
    End If
    
    NumLinesFound = 0
    i = 0: j = 1
    
    Starts(1) = i + 1
    If SkipToLine = 1 Then HaveReachedSkipToLine = True

    Do
        If Not Streaming Then
            If PosLF <= i Then PosLF = InStr(i + 1, Buffer, vbLf): If PosLF = 0 Then PosLF = BufferUpdatedTo + 1
            If PosCR <= i Then PosCR = InStr(i + 1, Buffer, vbCr): If PosCR = 0 Then PosCR = BufferUpdatedTo + 1
            If PosCR < PosLF Then
                FoundCR = True
                i = PosCR
            Else
                FoundCR = False
                i = PosLF
            End If
        Else
            i = SearchInBuffer(SearchFor, i + 1, Stream, useADODB, vbNullString, _
                vbNullString, Which, Buffer, BufferUpdatedTo)
            FoundCR = (Which = 2)
        End If

        If i >= BufferUpdatedTo + 1 Then
            Exit Do
        End If

        If j + 1 > UBound(Starts) Then
            ReDim Preserve Starts(1 To UBound(Starts) * 2)
            ReDim Preserve lengths(1 To UBound(lengths) * 2)
        End If

        lengths(j) = i - Starts(j)
        If FoundCR Then
            If Mid$(Buffer, i + 1, 1) = vbLf Then
                'Ending is Windows rather than Mac or Unix.
                i = i + 1
            End If
        End If
                    
        Starts(j + 1) = i + 1
                    
        j = j + 1
        NumLinesFound = NumLinesFound + 1
        If Not HaveReachedSkipToLine Then
            If NumLinesFound = SkipToLine - 1 Then
                HaveReachedSkipToLine = True
                Tmp = Starts(j)
                ReDim Starts(1 To 8): ReDim lengths(1 To 8)
                j = 1: NumLinesFound = 0
                Starts(1) = Tmp
            End If
        ElseIf NumLinesToReturn > 0 Then
            If NumLinesFound = NumLinesToReturn Then
                Exit Do
            End If
        End If
    Loop
   
    If SkipToLine > NumLinesFound Then
        If NumLinesToReturn = 0 Then 'Attempting to read from SkipToLine to the end of the file, but that would _
                                      be zero or a negative number of rows. So throw an error.
                             
            Throw "SkipToLine (" & CStr(SkipToLine) & ") exceeds the number of lines in the file (" & _
                CStr(NumLinesFound) & ")"
        Else
            'Attempting to read a set number of rows, function will return an array of null strings
            NumLinesFound = 0
        End If
    End If
    If NumLinesToReturn = 0 Then NumLinesToReturn = NumLinesFound

    ReDim ReturnArray(1 To NumLinesToReturn, 1 To 1)
    MSLIA = MaxStringLengthInArray()
    For i = 1 To MinLngs(NumLinesToReturn, NumLinesFound)
        If CallingFromWorksheet Then
            If lengths(i) > MSLIA Then
                Err_StringTooLong = "Line " & Format$(i, "#,###") & " of the file is of length " + Format$(lengths(i), "###,###")
                If MSLIA >= 32767 Then
                    Err_StringTooLong = Err_StringTooLong & ". Excel cells cannot contain strings longer than " + Format$(MSLIA, "####,####")
                Else
                    Err_StringTooLong = Err_StringTooLong & _
                        ". An array containing a string longer than " + Format$(MSLIA, "###,###") + _
                        " cannot be returned from VBA to an Excel worksheet"
                End If
                Throw Err_StringTooLong
            End If
        End If
        ReturnArray(i, 1) = Mid$(Buffer, Starts(i), lengths(i))
    Next i

    ParseTextFile = ReturnArray

    Exit Function
ErrHandler:
    Throw "#ParseTextFile: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Purpose    : Test if Excel's Function Wizard is active to allow early exit in slow functions.
' https://stackoverflow.com/questions/20866484/can-i-disable-a-vba-udf-calculation-when-the-insert-function-function-arguments
' -----------------------------------------------------------------------------------------------------------------------
Private Function FunctionWizardActive() As Boolean
    
    On Error GoTo ErrHandler
    If Not Application.CommandBars.item("Standard").Controls.item(1).Enabled Then
        FunctionWizardActive = True
    End If

    Exit Function
ErrHandler:
    Throw "#FunctionWizardActive: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Throw
' Purpose    : Simple error handling.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Throw(ByVal ErrorString As String)
    Err.Raise vbObjectError + 1, , ErrorString
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
' -----------------------------------------------------------------------------------------------------------------------
Public Function ThrowIfError(Data As Variant) As Variant
    ThrowIfError = Data
    If VarType(Data) = vbString Then
        If Left$(Data, 1) = "#" Then
            If Right$(Data, 1) = "!" Then
                Throw CStr(Data)
            End If
        End If
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileExists
' Purpose    : Returns True if FileName exists on disk, False o.w.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileExists(FileName As String) As Boolean
    Dim F As Scripting.File
    On Error GoTo ErrHandler
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set F = m_FSO.GetFile(FileName)
    FileExists = True
    Exit Function
ErrHandler:
    FileExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FolderExists
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FolderExists(FolderPath As String) As Boolean
    Dim F As Scripting.Folder
    
    On Error GoTo ErrHandler
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    
    Set F = m_FSO.GetFolder(FolderPath)
    FolderExists = True
    Exit Function
ErrHandler:
    FolderExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileDelete
' Purpose    : Delete a file, returns True or error string.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub FileDelete(FileName As String)
    Dim F As Scripting.File
    On Error GoTo ErrHandler

    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set F = m_FSO.GetFile(FileName)
    F.Delete

    Exit Sub
ErrHandler:
    Throw "#FileDelete: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful throws an error.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub CreatePath(ByVal FolderPath As String)

    Dim F As Scripting.Folder
    Dim i As Long
    Dim ParentFolderName As String
    Dim ThisFolderName As String

    On Error GoTo ErrHandler

    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject

    If Left$(FolderPath, 2) = "\\" Then
    ElseIf Mid$(FolderPath, 2, 2) <> ":\" Or _
        Asc(UCase$(Left$(FolderPath, 1))) < 65 Or _
        Asc(UCase$(Left$(FolderPath, 1))) > 90 Then
        Throw "First three characters of FolderPath must give drive letter followed by "":\"" or else be""\\"" for " & _
            "UNC folder name"
    End If

    FolderPath = Replace(FolderPath, "/", "\")

    If Right$(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If

    If FolderExists(FolderPath) Then
        GoTo EarlyExit
    End If

    'Go back until we find parent folder that does exist
    For i = Len(FolderPath) - 1 To 3 Step -1
        If Mid$(FolderPath, i, 1) = "\" Then
            If FolderExists(Left$(FolderPath, i)) Then
                Set F = m_FSO.GetFolder(Left$(FolderPath, i))
                ParentFolderName = Left$(FolderPath, i)
                Exit For
            End If
        End If
    Next i

    If F Is Nothing Then Throw "Cannot create folder " & Left$(FolderPath, 3)

    'now add folders one level at a time
    For i = Len(ParentFolderName) + 1 To Len(FolderPath)
        If Mid$(FolderPath, i, 1) = "\" Then
            
            ThisFolderName = Mid$(FolderPath, InStrRev(FolderPath, "\", i - 1) + 1, _
                i - 1 - InStrRev(FolderPath, "\", i - 1))
            F.SubFolders.Add ThisFolderName
            Set F = m_FSO.GetFolder(Left$(FolderPath, i))
        End If
    Next i

EarlyExit:
    Set F = m_FSO.GetFolder(FolderPath)
    Set F = Nothing

    Exit Sub
ErrHandler:
    Throw "#CreatePath: " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromPath
' Purpose    : Split file-with-path to file name (if ReturnFileName is True) or path otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FileFromPath(FullFileName As String, Optional ReturnFileName As Boolean = True) As Variant
    Dim SlashPos As Long
    Dim SlashPos2 As Long

    On Error GoTo ErrHandler

    SlashPos = InStrRev(FullFileName, "\")
    SlashPos2 = InStrRev(FullFileName, "/")
    If SlashPos2 > SlashPos Then SlashPos = SlashPos2
    If SlashPos = 0 Then Throw "Neither '\' nor '/' found"

    If ReturnFileName Then
        FileFromPath = Mid$(FullFileName, SlashPos + 1)
    Else
        FileFromPath = Left$(FullFileName, SlashPos - 1)
    End If

    Exit Function
ErrHandler:
    Throw "#FileFromPath: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsNumber
' Purpose   : Is a singleton a number?
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsNumber(x As Variant) As Boolean
    Select Case VarType(x)
        Case vbDouble, vbInteger, vbSingle, vbLong ', vbCurrency, vbDecimal
            IsNumber = True
    End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NCols
' Purpose   : Number of columns in an array. Missing has zero rows, 1-dimensional arrays
'             have one row and the number of columns returned by this function.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NCols(Optional TheArray As Variant) As Long
    If TypeName(TheArray) = "Range" Then
        NCols = TheArray.Columns.count
    ElseIf IsMissing(TheArray) Then
        NCols = 0
    ElseIf VarType(TheArray) < vbArray Then
        NCols = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                NCols = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
            Case Else
                NCols = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
        End Select
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NRows
' Purpose   : Number of rows in an array. Missing has zero rows, 1-dimensional arrays have one row.
' -----------------------------------------------------------------------------------------------------------------------
Private Function NRows(Optional TheArray As Variant) As Long
    If TypeName(TheArray) = "Range" Then
        NRows = TheArray.Rows.count
    ElseIf IsMissing(TheArray) Then
        NRows = 0
    ElseIf VarType(TheArray) < vbArray Then
        NRows = 1
    Else
        Select Case NumDimensions(TheArray)
            Case 1
                NRows = 1
            Case Else
                NRows = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
        End Select
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Transpose
' Purpose   : Returns the transpose of an array.
' Arguments
' TheArray  : An array of arbitrary values.
'             Return is always 1-based, even when input is zero-based.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Transpose(TheArray As Variant) As Variant
    Dim Co As Long
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim N As Long
    Dim Result As Variant
    Dim Ro As Long
    On Error GoTo ErrHandler
    Force2DArrayR TheArray, N, m
    Ro = LBound(TheArray, 1) - 1
    Co = LBound(TheArray, 2) - 1
    ReDim Result(1 To m, 1 To N)
    For i = 1 To N
        For j = 1 To m
            Result(j, i) = TheArray(i + Ro, j + Co)
        Next j
    Next i
    Transpose = Result
    Exit Function
ErrHandler:
    Throw "#Transpose: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArrayR
' Purpose   : When writing functions to be called from sheets, we often don't want to process
'             the inputs as Range objects, but instead as Arrays. This method converts the
'             input into a 2-dimensional 1-based array (even if it's a single cell or single row of cells)
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Force2DArrayR(ByRef RangeOrArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
    If TypeName(RangeOrArray) = "Range" Then RangeOrArray = RangeOrArray.Value2
    Force2DArray RangeOrArray, NR, NC
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Force2DArray
' Purpose   : In-place amendment of singletons and one-dimensional arrays to two dimensions.
'             singletons and 1-d arrays are returned as 2-d 1-based arrays. Leaves two
'             two dimensional arrays untouched (i.e. a zero-based 2-d array will be left as zero-based).
'             See also Force2DArrayR that also handles Range objects.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Force2DArray(ByRef TheArray As Variant, Optional ByRef NR As Long, Optional ByRef NC As Long)
    Dim TwoDArray As Variant

    On Error GoTo ErrHandler

    Select Case NumDimensions(TheArray)
        Case 0
            ReDim TwoDArray(1 To 1, 1 To 1)
            TwoDArray(1, 1) = TheArray
            TheArray = TwoDArray
            NR = 1: NC = 1
        Case 1
            Dim i As Long
            Dim LB As Long
            LB = LBound(TheArray, 1)
            NR = 1: NC = UBound(TheArray, 1) - LB + 1
            ReDim TwoDArray(1 To 1, 1 To NC)
            For i = 1 To UBound(TheArray, 1) - LBound(TheArray) + 1
                TwoDArray(1, i) = TheArray(LB + i - 1)
            Next i
            TheArray = TwoDArray
        Case 2
            NR = UBound(TheArray, 1) - LBound(TheArray, 1) + 1
            NC = UBound(TheArray, 2) - LBound(TheArray, 2) + 1
            'Nothing to do
        Case Else
            Throw "Cannot convert array of dimension greater than two"
    End Select

    Exit Sub
ErrHandler:
    Throw "#Force2DArray: " & Err.Description & "!"
End Sub

Private Function MaxLngs(x As Long, y As Long) As Long
    If x > y Then
        MaxLngs = x
    Else
        MaxLngs = y
    End If
End Function

Private Function MinLngs(x As Long, y As Long) As Long
    If x > y Then
        MinLngs = y
    Else
        MinLngs = x
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterCSVWrite
' Purpose    : Register the function CSVWrite with the Excel function wizard. Suggest this function is called from a
'              WorkBook_Open event.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RegisterCSVWrite()
    Const Description As String = "Creates a comma-separated file on disk containing Data. Any existing file of " & _
                                  "the same name is overwritten. If successful, the function returns FileName, " & _
                                  "otherwise an ""error string"" (starts with `#`, ends with `!`) describing what " & _
                                  "went wrong."
    Dim ArgDescs() As String

    On Error GoTo ErrHandler

    ReDim ArgDescs(1 To 8)
    ArgDescs(1) = "An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty, " & _
                  "Excel errors or null values. Data typically has two dimensions, but if Data has only one " & _
                  "dimension then the output file has a single column, one field per row."
    ArgDescs(2) = "The full name of the file, including the path. Alternatively, if FileName is omitted, then the " & _
                  "function returns Data converted CSV-style to a string."
    ArgDescs(3) = "If TRUE (the default) then all strings in Data are quoted before being written to file. If " & _
                  "FALSE only strings containing Delimiter, line feed, carriage return or double quote are quoted. " & _
                  "Double quotes are always escaped by another double quote."
    ArgDescs(4) = "A format string that determines how dates, including cells formatted as dates, appear in the " & _
                  "file. If omitted, defaults to `yyyy-mm-dd`."
    ArgDescs(5) = "Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use `ISOZ` for " & _
                  "ISO8601 format with time zone the same as the PC's clock. Use with care, daylight saving may be " & _
                  "inconsistent across the datetimes in data."
    ArgDescs(6) = "The delimiter string, if omitted defaults to a comma. Delimiter may have more than one " & _
                  "character."
    ArgDescs(7) = "Allowed entries are `ANSI` (the default), `UTF-8` and `UTF-16`. An error will result if this " & _
                  "argument is `ANSI` but Data contains characters that cannot be written to an ANSI file. `UTF-8` " & _
                  "and `UTF-16` files are written with a byte option mark."
    ArgDescs(8) = "Controls the line endings of the file written. Enter `Windows` (the default), `Unix` or `Mac`. " & _
                  "Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13) or " & _
                  "the strings `CRLF`, `LF` or `CR`."
    Application.MacroOptions "CSVWrite", Description, , , , , , , , , ArgDescs
    Exit Sub

ErrHandler:
    Debug.Print "Warning: Registration of function CSVWrite failed with error: " + Err.Description
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CSVWrite
' Purpose   : Creates a comma-separated file on disk containing Data. Any existing file of the same
'             name is overwritten. If successful, the function returns FileName, otherwise an "error
'             string" (starts with `#`, ends with `!`) describing what went wrong.
' Arguments
' Data      : An array of data, or an Excel range. Elements may be strings, numbers, dates, Booleans, empty,
'             Excel errors or null values. Data typically has two dimensions, but if Data has only one
'             dimension then the output file has a single column, one field per row.
' FileName  : The full name of the file, including the path. Alternatively, if FileName is omitted, then the
'             function returns Data converted CSV-style to a string.
' QuoteAllStrings: If `TRUE` (the default) then elements of Data that are strings are quoted before being
'             written to file, other elements (Numbers, Booleans, Errors) are not quoted. If `FALSE` then
'             the only elements of Data that are quoted are strings containing Delimiter, line feed,
'             carriage return or double quote. In all cases, double quotes are escaped by another double
'             quote.
' DateFormat: A format string that determines how dates, including cells formatted as dates, appear in the
'             file. If omitted, defaults to `yyyy-mm-dd`.
' DateTimeFormat: Format for datetimes. Defaults to `ISO` which abbreviates `yyyy-mm-ddThh:mm:ss`. Use
'             `ISOZ` for ISO8601 format with time zone the same as the PC's clock. Use with care, daylight
'             saving may be inconsistent across the datetimes in data.
' Delimiter : The delimiter string, if omitted defaults to a comma. Delimiter may have more than one
'             character.
' Encoding  : Allowed entries are `ANSI` (the default), `UTF-8` and `UTF-16`. An error will result if this
'             argument is `ANSI` but Data contains characters that cannot be written to an ANSI file.
'             `UTF-8` and `UTF-16` files are written with a byte option mark.
' EOL       : Controls the line endings of the file written. Enter `Windows` (the default), `Unix` or `Mac`.
'             Also supports the line-ending characters themselves (ascii 13 + ascii 10, ascii 10, ascii 13)
'             or the strings `CRLF`, `LF` or `CR`. The last line of the file is written with a line ending.
'
' Notes     : See also companion function CSVRead.
'
'             For discussion of the CSV format see
'             https://tools.ietf.org/html/rfc4180#section-2
' -----------------------------------------------------------------------------------------------------------------------
Public Function CSVWrite(ByVal Data As Variant, Optional ByVal FileName As String, _
    Optional ByVal QuoteAllStrings As Boolean = True, Optional ByVal DateFormat As String = "YYYY-MM-DD", _
    Optional ByVal DateTimeFormat As String = "ISO", Optional ByVal Delimiter As String = ",", _
    Optional ByVal Encoding As String = "ANSI", Optional ByVal EOL As String = vbNullString) As String

    Const DQ As String = """"
    Const Err_Delimiter As String = "Delimiter must have at least one character and cannot start with a " & _
        "double  quote, line feed or carriage return"
    Const Err_Dimensions As String = "Data has more than two dimensions, which is not supported"
    Const Err_Encoding As String = "Encoding must be ""ANSI"" (the default) or ""UTF-8"" or ""UTF-16"""
    
    Dim EOLIsWindows As Boolean
    Dim ErrRet As String
    Dim i As Long
    Dim j As Long
    Dim Lines() As String
    Dim OneLine() As String
    Dim OneLineJoined As String
    Dim Stream As Object
    Dim Unicode As Boolean
    Dim WriteToFile As Boolean

    On Error GoTo ErrHandler
    
    Select Case UCase$(Encoding)
        Case "ANSI", "UTF-8", "UTF-16"
        Case Else
            Throw Err_Encoding
    End Select

    WriteToFile = Len(FileName) > 0

    If EOL = vbNullString Then
        If WriteToFile Then
            EOL = vbCrLf
        Else
            EOL = vbLf
        End If
    End If

    EOL = OStoEOL(EOL, "EOL")
    EOLIsWindows = EOL = vbCrLf

    If Len(Delimiter) = 0 Or Left$(Delimiter, 1) = DQ Or Left$(Delimiter, 1) = vbLf Or Left$(Delimiter, 1) = vbCr Then
        Throw Err_Delimiter
    End If
    
    Select Case UCase$(DateTimeFormat)
        Case "ISO"
            DateTimeFormat = "yyyy-mm-ddThh:mm:ss"
        Case "ISOZ"
            DateTimeFormat = ISOZFormatString()
    End Select

    If TypeName(Data) = "Range" Then
        'Preserve elements of type Date by using .Value, not .Value2
        Data = Data.value
    End If
    Select Case NumDimensions(Data)
        Case 0
            Dim Tmp() As Variant
            ReDim Tmp(1 To 1, 1 To 1)
            Tmp(1, 1) = Data
            Data = Tmp
        Case 1
            ReDim Tmp(LBound(Data) To UBound(Data), 1 To 1)
            For i = LBound(Data) To UBound(Data)
                Tmp(i, 1) = Data(i)
            Next i
            Data = Tmp
        Case Is > 2
            Throw Err_Dimensions
    End Select
    
    ReDim OneLine(LBound(Data, 2) To UBound(Data, 2))
    
    If WriteToFile Then
        If UCase$(Encoding) = "UTF-8" Then
            Set Stream = CreateObject("ADODB.Stream")
            Stream.Open
            Stream.Type = 2 'Text
            Stream.CharSet = "utf-8"
    
            For i = LBound(Data) To UBound(Data)
                For j = LBound(Data, 2) To UBound(Data, 2)
                    OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",")
                Next j
                OneLineJoined = VBA.Join(OneLine, Delimiter) & EOL
                Stream.WriteText OneLineJoined
            Next i
            Stream.SaveToFile FileName, 2 'adSaveCreateOverWrite

            CSVWrite = FileName
        Else
            Unicode = UCase$(Encoding) = "UTF-16"
            If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
            Set Stream = m_FSO.CreateTextFile(FileName, True, Unicode)
  
            For i = LBound(Data) To UBound(Data)
                For j = LBound(Data, 2) To UBound(Data, 2)
                    OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",")
                Next j
                OneLineJoined = VBA.Join(OneLine, Delimiter)
                WriteLineWrap Stream, OneLineJoined, EOLIsWindows, EOL, Unicode
            Next i

            Stream.Close: Set Stream = Nothing
            CSVWrite = FileName
        End If
    Else

        ReDim Lines(LBound(Data) To UBound(Data) + 1) 'add one to ensure that result has a terminating EOL
  
        For i = LBound(Data) To UBound(Data)
            For j = LBound(Data, 2) To UBound(Data, 2)
                OneLine(j) = Encode(Data(i, j), QuoteAllStrings, DateFormat, DateTimeFormat, ",")
            Next j
            Lines(i) = VBA.Join(OneLine, Delimiter)
        Next i
        CSVWrite = VBA.Join(Lines, EOL)
        If Len(CSVWrite) > 32767 Then
            If TypeName(Application.Caller) = "Range" Then
                Throw "Cannot return string of length " & Format$(CStr(Len(CSVWrite)), "#,###") & _
                    " to a cell of an Excel worksheet"
            End If
        End If
    End If
    
    Exit Function
ErrHandler:
    ErrRet = "#CSVWrite: " & Err.Description & "!"
    If Not Stream Is Nothing Then
        Stream.Close
        Set Stream = Nothing
    End If
    If m_ErrorStyle = es_ReturnString Then
        CSVWrite = ErrRet
    Else
        Throw ErrRet
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OStoEOL
' Purpose    : Convert text describing an operating system to the end-of-line marker employed. Note that "Mac" converts
'              to vbCr but Apple operating systems since OSX use vbLf, matching Unix.
' -----------------------------------------------------------------------------------------------------------------------
Private Function OStoEOL(OS As String, ArgName As String) As String

    Const Err_Invalid As String = " must be one of ""Windows"", ""Unix"" or ""Mac"", or the associated end of line characters."

    On Error GoTo ErrHandler
    Select Case LCase$(OS)
        Case "windows", vbCrLf, "crlf"
            OStoEOL = vbCrLf
        Case "unix", "linux", vbLf, "lf"
            OStoEOL = vbLf
        Case "mac", vbCr, "cr"
            OStoEOL = vbCr
        Case Else
            Throw ArgName & Err_Invalid
    End Select

    Exit Function
ErrHandler:
    Throw "#OStoEOL: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Purpose    : Encode arbitrary value as a string, sub-routine of CSVWrite.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Encode(x As Variant, QuoteAllStrings As Boolean, DateFormat As String, DateTimeFormat As String, _
     Delim As String) As String
    
    Const DQ As String = """"
    Const DQ2 As String = """"""

    On Error GoTo ErrHandler
    Select Case VarType(x)

        Case vbString
            If InStr(x, DQ) > 0 Then
                Encode = DQ & Replace(x, DQ, DQ2) & DQ
            ElseIf QuoteAllStrings Then
                Encode = DQ & x & DQ
            ElseIf InStr(x, vbCr) > 0 Then
                Encode = DQ & x & DQ
            ElseIf InStr(x, vbLf) > 0 Then
                Encode = DQ & x & DQ
            ElseIf InStr(x, Delim) > 0 Then
                Encode = DQ & x & DQ
            Else
                Encode = x
            End If
        Case vbBoolean, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbEmpty 'vbLongLong - not available on 16 bit.
            Encode = CStr(x)
        Case vbDate
            If CLng(x) = CDbl(x) Then
                Encode = Format$(x, DateFormat)
            Else
                Encode = Format$(x, DateTimeFormat)
            End If
        Case vbNull
            Encode = "NULL"
        Case vbError
            Select Case CStr(x) 'Editing this case statement? Edit also its inverse, see method MakeSentinels
                Case "Error 2000"
                    Encode = "#NULL!"
                Case "Error 2007"
                    Encode = "#DIV/0!"
                Case "Error 2015"
                    Encode = "#VALUE!"
                Case "Error 2023"
                    Encode = "#REF!"
                Case "Error 2029"
                    Encode = "#NAME?"
                Case "Error 2036"
                    Encode = "#NUM!"
                Case "Error 2042"
                    Encode = "#N/A"
                Case "Error 2043"
                    Encode = "#GETTING_DATA!"
                Case "Error 2045"
                    Encode = "#SPILL!"
                Case "Error 2046"
                    Encode = "#CONNECT!"
                Case "Error 2047"
                    Encode = "#BLOCKED!"
                Case "Error 2048"
                    Encode = "#UNKNOWN!"
                Case "Error 2049"
                    Encode = "#FIELD!"
                Case "Error 2050"
                    Encode = "#CALC!"
                Case Else
                    Encode = CStr(x)        'should never hit this line...
            End Select
        Case Else
            Throw "Cannot convert variant of type " & TypeName(x) & " to String"
    End Select
    Exit Function
ErrHandler:
    Throw "#Encode: " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WriteLineWrap
' Purpose    : Wrapper to TextStream.Write[Line] to give more informative error message than "invalid procedure call or
'              argument" if the error is caused by attempting to write illegal characters to a stream opened with
'              TriStateFalse.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub WriteLineWrap(t As TextStream, text As String, EOLIsWindows As Boolean, EOL As String, Unicode As Boolean)

    Dim ErrDesc As String
    Dim ErrNum As Long
    Dim i As Long

    On Error GoTo ErrHandler
    If EOLIsWindows Then
        t.WriteLine text
    Else
        t.Write text
        t.Write EOL
    End If

    Exit Sub

ErrHandler:
    ErrNum = Err.Number
    ErrDesc = Err.Description
    If Not Unicode Then
        If ErrNum = 5 Then
            For i = 1 To Len(text)
                If Not CanWriteCharToAscii(Mid$(text, i, 1)) Then
                    ErrDesc = "Data contains characters that cannot be written to an ascii file (first found is '" & _
                        Mid$(text, i, 1) & "' with unicode character code " & AscW(Mid$(text, i, 1)) & _
                        "). Try calling CSVWrite with argument Encoding as ""UTF-8"" or ""UTF-16"""
                    Exit For
                End If
            Next i
        End If
    End If
    Throw "#WriteLineWrap: " & ErrDesc & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CanWriteCharToAscii
' Purpose    : Not all characters for which AscW(c) < 255 can be written to an ascii file. If AscW(c) is in the following
'              list then they cannot:
'             128,130,131,132,133,134,135,136,137,138,139,140,142,145,146,147,148,149,150,151,152,153,154,155,156,158,159
' -----------------------------------------------------------------------------------------------------------------------
Private Function CanWriteCharToAscii(c As String) As Boolean
    Dim code As Long
    code = AscW(c)
    If code > 255 Or code < 0 Then
        CanWriteCharToAscii = False
    Else
        CanWriteCharToAscii = Chr$(AscW(c)) = c
    End If
End Function

