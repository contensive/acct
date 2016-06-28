VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RFPExportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const ContentNamePeople = "People"
Const ContentNameSessionType = "Session Types"
Const ContentNameSessionSubmissions = "Session Submissions"
Const ContentNameSessionSubmissionsRules = "Session Submissions Rules"
Const ContentNameContactRules = "Session Submissions Contact Rules"
Const ContentNameSessionTracks = "Session Tracks"
'
Private Main As Object
'
Public Sub Init(MainObject As Variant)
    On Error GoTo ErrorTrap
    '
    Set Main = MainObject
    '
    Call Main.SetStreamTimeout(1800)
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleError("RFPExportClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim FileLink As String
    Dim Link As String
    '
    If Main.GetStreamText("button") <> "" Then
Call Main.TestPoint("GetContent - 100")
'        Call Main.ClearStream
'        Call Main.SetStreamType("text/comma-separated-values")
'        Call Main.SetStreamHeader("Content-Disposition", "attachment; filename=RFPExport.csv")
'        Call Main.WriteStream(GetExport())
'        Call Main.CloseStream
        Link = "RFPExports\RFPExport" & Rnd() & ".csv"
Call Main.TestPoint("GetContent - 200")
        Call Main.SaveVirtualFile(Link, GetExport())
Call Main.TestPoint("GetContent - 300")
        FileLink = "<a href=""" & Main.ServerFilePath & Link & """>Click here to download the export</a>"
Call Main.TestPoint("GetContent - 400")
    End If
    '
    GetContent = Main.GetFormStart() _
                & "<table cellpadding=""10"" cellspacing=""0"" border=""0"" width=""100%"">" _
                & "<tr>" _
                & "<td>" _
                & FileLink & "<br /><br />" _
                & Main.GetFormButton("  Export Completed RFPs  ") _
                & "</td>" _
                & "</tr>" _
                & "</table>" _
                & Main.GetFormEnd
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPExportClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetExport() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CS As Long
    Dim CSCount As Long
    Dim SQL As String
    Dim ParticipantCount As Long
    Dim ContactCount As Long
    Dim Pointer As Long
    Dim Header As String
    Dim File As String
    Dim SubmissionID As Long
    '
Call Main.TestPoint("GetExport - 100")
    CS = Main.OpenCSContent(ContentNameSessionSubmissions, "Completed=1")
    If Main.CSOK(CS) Then
        '
Call Main.TestPoint("GetExport - 200")
        '
'        SQL = "SELECT Count(SessionSubmissions.ID) AS ContactCount"
'        SQL = SQL & " FROM SessionSubmissions LEFT JOIN SessionSubmissionsRules ON SessionSubmissions.ID = SessionSubmissionsRules.SessionSubmissionsID"
'        SQL = SQL & " Where (((SessionSubmissions.Completed) = -1))"
'        SQL = SQL & " GROUP BY SessionSubmissionsRules.ContentControlID"
'        SQL = SQL & " HAVING (((SessionSubmissionsRules.ContentControlID)=" & Main.GetRecordID("Content", ContentNameContactRules) & "))"
'        SQL = SQL & " ORDER BY Count(SessionSubmissions.ID) DESC;"
        
        SQL = "SELECT Count(SessionSubmissionsRules.MemberID) AS ContactCount, SessionSubmissionsRules.SessionSubmissionsID"
        SQL = SQL & " From SessionSubmissionsRules"
        SQL = SQL & " GROUP BY SessionSubmissionsRules.ContentControlID, SessionSubmissionsRules.SessionSubmissionsID"
        SQL = SQL & " Having (((SessionSubmissionsRules.ContentControlID) = " & Main.GetRecordID("Content", ContentNameContactRules) & "))"
        SQL = SQL & " ORDER BY Count(SessionSubmissionsRules.MemberID) DESC;"
        
        '
        CSCount = Main.OpenCSSQL("Default", SQL)
        If Main.CSOK(CSCount) Then
            ContactCount = Main.GetCSInteger(CSCount, "ContactCount")
        End If
        Call Main.CloseCS(CSCount)
Call Main.TestPoint("GetExport - 300")
        '
'        SQL = "SELECT Count(SessionSubmissions.ID) AS ParticipantCount"
'        SQL = SQL & " FROM SessionSubmissions LEFT JOIN SessionSubmissionsRules ON SessionSubmissions.ID = SessionSubmissionsRules.SessionSubmissionsID"
'        SQL = SQL & " Where (((SessionSubmissions.Completed) = -1))"
'        SQL = SQL & " GROUP BY SessionSubmissionsRules.ContentControlID"
'        SQL = SQL & " HAVING (((SessionSubmissionsRules.ContentControlID)=" & Main.GetRecordID("Content", ContentNameSessionSubmissionsRules) & "))"
'        SQL = SQL & " ORDER BY Count(SessionSubmissions.ID) DESC;"
        
'        SQL = "SELECT Count(SessionSubmissionsRules.MemberID) AS ParticipantCount, SessionSubmissionsRules.SessionSubmissionsID"
'        SQL = SQL & " From SessionSubmissionsRules"
'        SQL = SQL & " GROUP BY SessionSubmissionsRules.ContentControlID, SessionSubmissionsRules.SessionSubmissionsID"
'        SQL = SQL & " Having (((SessionSubmissionsRules.ContentControlID) = " & Main.GetRecordID("Content", ContentNameSessionSubmissionsRules) & "))"
'        SQL = SQL & " ORDER BY Count(SessionSubmissionsRules.MemberID) DESC;"
        
        SQL = "SELECT Count(SessionSubmissionsRules.MemberID) AS ParticipantCount"
        SQL = SQL & " FROM SessionSubmissionsRules LEFT JOIN SessionSubmissions ON SessionSubmissionsRules.SessionSubmissionsID = SessionSubmissions.ID"
        SQL = SQL & " GROUP BY SessionSubmissionsRules.SessionSubmissionsID, SessionSubmissionsRules.ContentControlID, SessionSubmissions.Completed"
        SQL = SQL & " Having (((SessionSubmissionsRules.ContentControlID) = " & Main.GetRecordID("Content", ContentNameSessionSubmissionsRules) & ") And ((SessionSubmissions.Completed) = -1))"
        SQL = SQL & " ORDER BY Count(SessionSubmissionsRules.MemberID) DESC;"
        
        CSCount = Main.OpenCSSQL("Default", SQL)
        If Main.CSOK(CSCount) Then
            ParticipantCount = Main.GetCSInteger(CSCount, "ParticipantCount")
        End If
        Call Main.CloseCS(CSCount)
Call Main.TestPoint("GetExport - 400")
        '
        Header = "id,Date Added,Name,Description,Session Type,Track,Outcome,Support Letter,PowerPoint,Video,Internet"
        '
        For Pointer = 1 To ContactCount
            Header = Header & ",Contact " & Pointer & " First Name" _
                            & ",Contact " & Pointer & " Last Name" _
                            & ",Contact " & Pointer & " Title" _
                            & ",Contact " & Pointer & " Organization" _
                            & ",Contact " & Pointer & " Address" _
                            & ",Contact " & Pointer & " Address 2" _
                            & ",Contact " & Pointer & " City" _
                            & ",Contact " & Pointer & " State" _
                            & ",Contact " & Pointer & " Zip" _
                            & ",Contact " & Pointer & " Phone" _
                            & ",Contact " & Pointer & " Fax" _
                            & ",Contact " & Pointer & " Email"
        Next
Call Main.TestPoint("GetExport - 500")
        For Pointer = 1 To ParticipantCount
            Header = Header & ",Participant " & Pointer & " Presenter" _
                            & ",Participant " & Pointer & " First Name" _
                            & ",Participant " & Pointer & " Last Name" _
                            & ",Participant " & Pointer & " Title" _
                            & ",Participant " & Pointer & " Organization" _
                            & ",Participant " & Pointer & " Address" _
                            & ",Participant " & Pointer & " Address 2" _
                            & ",Participant " & Pointer & " City" _
                            & ",Participant " & Pointer & " State" _
                            & ",Participant " & Pointer & " Zip" _
                            & ",Participant " & Pointer & " Phone" _
                            & ",Participant " & Pointer & " Fax" _
                            & ",Participant " & Pointer & " Email"
        Next
Call Main.TestPoint("GetExport - 600")
        Stream = Header & vbCrLf
        '
        Do While Main.CSOK(CS)
Call Main.TestPoint("GetExport - 610")
            SubmissionID = Main.GetCSInteger(CS, "ID")
            Stream = Stream & FormatValue(CStr(SubmissionID)) & ","
            Stream = Stream & FormatValue(Main.GetCSText(CS, "DateAdded")) & ","
            Stream = Stream & FormatValue(Main.GetCSText(CS, "Name")) & ","
            Stream = Stream & FormatValue(Main.GetCSText(CS, "Description")) & ","
            Stream = Stream & FormatValue(Main.GetRecordName(ContentNameSessionType, Main.GetCSInteger(CS, "TypeID"))) & ","
            Stream = Stream & FormatValue(Main.GetRecordName(ContentNameSessionTracks, Main.GetCSInteger(CS, "TrackID"))) & ","
            Stream = Stream & FormatValue(Main.GetCSText(CS, "Outcome")) & ","
            File = Main.GetCSText(CS, "BoardLetter")
            If File <> "" Then
                Stream = Stream & FormatValue("http://" & Main.ServerHost & Main.ServerFilePath & File) & ","
            Else
                Stream = Stream & ","
            End If
Call Main.TestPoint("GetExport - 615")
            Stream = Stream & kmaGetYesNo(Main.GetCSBoolean(CS, "PowerPoint")) & ","
            Stream = Stream & kmaGetYesNo(Main.GetCSBoolean(CS, "Video")) & ","
            Stream = Stream & kmaGetYesNo(Main.GetCSBoolean(CS, "Connection"))
Call Main.TestPoint("GetExport - 616")
            '
            Stream = Stream & Replace(GetAddedFields(SubmissionID, ContentNameContactRules, ContactCount), vbCrLf, "", , , vbTextCompare)
            Stream = Stream & Replace(GetAddedFields(SubmissionID, ContentNameSessionSubmissionsRules, ParticipantCount), vbCrLf, "", , , vbTextCompare)
            '
            Stream = Stream & vbCrLf
Call Main.TestPoint("GetExport - 620")
            Call Main.NextCSRecord(CS)
        Loop
Call Main.TestPoint("GetExport - 700")
    Else
Call Main.TestPoint("GetExport - 800")
        Stream = ""
    End If
    Call Main.CloseCS(CS)
Call Main.TestPoint("GetExport - 900")
    '
    GetExport = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPExportClass", "GetExport", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function FormatValue(Value As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Replace(Value, """", """""", , , vbTextCompare)
    Stream = Replace(Stream, vbCrLf, "", , , vbTextCompare)
    Stream = """" & Stream & """"
    '
    FormatValue = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPExportClass", "GetCharityExport", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetAddedFields(SubmissionID As Long, ContentName As String, UpperCount As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CS As Long
    Dim SQL As String
    Dim FieldCount As Long
    Dim Pointer As Long
    '
Call Main.TestPoint("GetAddedFields - 100")
    SQL = "SELECT ccMembers.Presenter, ccMembers.FirstName, ccMembers.LastName, ccMembers.Company, ccMembers.Title, ccMembers.Address, ccMembers.Address2, ccMembers.City, ccMembers.State, ccMembers.Zip, ccMembers.Email, ccMembers.Phone, ccMembers.Fax"
    SQL = SQL & " FROM SessionSubmissionsRules LEFT JOIN ccMembers ON SessionSubmissionsRules.MemberID = ccMembers.ID"
    SQL = SQL & " WHERE (SessionSubmissionsRules.SessionSubmissionsID=" & SubmissionID & ")"
    SQL = SQL & " AND (SessionSubmissionsRules.ContentControlID=" & Main.GetRecordID("Content", ContentName) & ")"
    SQL = SQL & " AND (SessionSubmissionsRules.Active<>0)"
    '
    CS = Main.OpenCSSQL("Default", SQL)
    Do While Main.CSOK(CS)
Call Main.TestPoint("GetAddedFields - 110")
        If ContentName = ContentNameSessionSubmissionsRules Then
            Stream = Stream & "," & FormatValue(kmaGetYesNo(Main.GetCSBoolean(CS, "Presenter")))
        End If
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "FirstName"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "LastName"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Title"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Company"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Address"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Address2"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "City"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "State"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Zip"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Phone"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Fax"))
        Stream = Stream & "," & FormatValue(Main.GetCSText(CS, "Email"))
        Call Main.NextCSRecord(CS)
        FieldCount = FieldCount + 1
Call Main.TestPoint("GetAddedFields - 120")
    Loop
    Call Main.CloseCS(CS)
    '
Call Main.TestPoint("GetAddedFields - 130")
    If FieldCount < UpperCount Then
        For Pointer = FieldCount To UpperCount - 1
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
            Stream = Stream & ","
        Next
    End If
Call Main.TestPoint("GetAddedFields - 140")
    '
    GetAddedFields = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPExportClass", "GetAddedFields", Err.Number, Err.Source, Err.Description, True, False)
    End Function
