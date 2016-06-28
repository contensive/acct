VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameButtonValue = "button"
'
Const ContentNamePeople = "People"
Const ContentNameSessionType = "Session Types"
Const ContentNameSessionSubmissions = "Session Submissions"
Const ContentNameSessionSubmissionsRules = "Session Submissions Rules"
'
Const FormButtonExport = "  Export Results  "
'
Const FormSpacer = "<img src=""/ccLib/images/spacer.gif"" width=25 height=1>"
Const FormSpacer2 = "<img src=""/ccLib/images/spacer.gif"" width=25 height=1>"
Const ButtonSpacer = "&nbsp;&nbsp;&nbsp;"
'
Const PageNameInstructions = "RFP Export Instructions"
'
Const RequestNameSessionTitle = "sessiontitle"
Const RequestNameDescription = "description"
Const RequestNameEditedDescription = "editeddescription"
Const RequestNameTypeID = "typeid"
Const RequestNameOutcome = "outcome"
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameTitle = "title"
Const RequestNameCompany = "company"
Const RequestNameAddress = "address"
Const RequestNameAddress2 = "address2"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameZip = "zip"
Const RequestNamePhone = "phone"
Const RequestNameFax = "fax"
Const RequestNameEmail = "email"
Const RequestNameAidDescription = "aiddescription"
Const RequestNameAudiovisualNotes = "aidenotes"
Const RequestNameAccepted = "accepted"
'
Const RequestNamePresenterFirstName = "presenterfirstname"
Const RequestNamePresenterLastName = "presenterlastname"
Const RequestNamePresenterTitle = "presentertitle"
Const RequestNamePresenterCompany = "presentercompany"
Const RequestNamePresenterAddress = "presenteraddress"
Const RequestNamePresenterAddress2 = "presenteraddress2"
Const RequestNamePresenterCity = "presentercity"
Const RequestNamePresenterState = "presenterstate"
Const RequestNamePresenterZip = "presenterzip"
Const RequestNamePresenterPhone = "presenterphone"
Const RequestNamePresenterFax = "presenterfax"
Const RequestNamePresenterEmail = "presenteremail"
'
Private ButtonValue As String
Private WorkingQueryString As String
'
Private ErrorString As String
'
Private Main As Object
'
Public Sub Init(MainObject As Variant)
    '
    Set Main = MainObject
    '
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString <> "" Then
        WorkingQueryString = "?" & WorkingQueryString & "&"
    Else
        WorkingQueryString = "?"
        End If
    '
    Call ProcessForm
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetFormExport()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    '
    If ButtonValue = FormButtonExport Then
        Call Main.SetStreamTimeout(300)
        Call Main.SetStreamType("text/plain")
        Call Main.ClearStream
        Call Main.SetStreamBuffer(False)
        Call Main.SetStreamHeader("Content-Disposition", "attachment; filename=RFP_Export.csv")
        Call Main.WriteStream(GetExport())
        Call Main.CloseStream
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormExport() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ButtonString As String
    '
    ButtonString = ButtonString & Main.GetFormButton(FormButtonExport)
    '
    Stream = Stream & Main.GetFormStart
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & GetPageHead(PageNameInstructions)
    '
    Stream = Stream & GetFormRow2("<b>Session Information</b>")
    '
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameSessionTitle), GetProperty("Suggested Session Title"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameDescription), GetProperty("Session Description"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameEditedDescription), GetProperty("Edited Description"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameTypeID), GetProperty("Session Type"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameOutcome), GetProperty("Anticipated Outcomes"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameAidDescription), GetProperty("Audiovisual Aids"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameAudiovisualNotes), GetProperty("Audiovisual Notes"))
    '
    Stream = Stream & GetFormRow2("<b>Primary Presenter</b>")
    '
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameFirstName), GetProperty("First Name"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameLastName), GetProperty("Last Name"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameTitle), GetProperty("Title"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameCompany), GetProperty("College/Organization"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameAddress), GetProperty("Address"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameAddress2), GetProperty("Address2"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameCity), GetProperty("City"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameState), GetProperty("State"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameZip), GetProperty("Zip"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePhone), GetProperty("Phone"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameFax), GetProperty("Fax"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameEmail), GetProperty("Email"))
    '
    Stream = Stream & GetFormRow2("<b>Additional Presenter(s)</b>")
    '
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterFirstName), GetProperty("First Name"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterLastName), GetProperty("Last Name"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterTitle), GetProperty("Title"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterCompany), GetProperty("College/Organization"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterAddress), GetProperty("Address"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterAddress2), GetProperty("Address2"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterCity), GetProperty("City"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterState), GetProperty("State"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterZip), GetProperty("Zip"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterPhone), GetProperty("Phone"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterFax), GetProperty("Fax"))
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenterEmail), GetProperty("Email"))
    '
    Stream = Stream & GetFormRow2("<b>Filters</b>")
    '
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameAccepted), "Accepted Only")
    '
    Stream = Stream & GetFormRow3(FormSpacer, ButtonString)
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd
    '
    GetFormExport = Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetFormExport", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetPageHead(PageName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=10 width=""100%"">"
    '
    If ErrorString <> "" Then
        Stream = Stream & "<span class=""ccError"">" & ErrorString & "</span>"
    Else
        Stream = Stream & Main.GetContentCopy(PageName)
        End If
    '
    Stream = Stream & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetPageHead = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetPageHead", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow(FieldCaption As String, Innards As String, Optional AlignLeft As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim AlignmentString As String
    '
    If Not AlignLeft Then
        AlignmentString = " align=right"
    Else
        AlignmentString = " align=left"
        End If
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD " & AlignmentString & ">" & FieldCaption & "</TD>"
    Stream = Stream & "<TD width=""100%"">" & Innards & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetFormRow = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow3(FieldCaption As String, Innards As String, Optional AlignLeft As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim AlignmentString As String
    '
    If Not AlignLeft Then
        AlignmentString = " align=right"
    Else
        AlignmentString = " align=left"
        End If
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD " & AlignmentString & ">" & FieldCaption & "</TD>"
    Stream = Stream & "<TD width=""100%"">" & Innards & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetFormRow3 = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetFormRow3", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow2(Innards As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=""100%"">" & Innards & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetFormRow2 = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetFormRow2", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetProperty(PropertyName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Main.GetSiteProperty(PropertyName, PropertyName, True)
    '
    GetProperty = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetProperty", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetExport() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim SQL As String
    '
    SQL = GetQuery()
    '
    CSPointer = Main.OpenCSSQL("default", SQL)
    If Main.CSOK(CSPointer) Then
        Do While Main.CSOK(CSPointer)
            
        Loop
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetExport = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetExport", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetQuery() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim TableList As String
    Dim FieldList As String
    Dim WhereClause As String
    Dim PeopleFlag As Boolean
    Dim RulesFlag As Boolean
    Dim TypeFlag As Boolean
    '
    Dim TablePeople As String
    Dim TableRules As String
    Dim TableType As String
    Dim TableRFP As String
    '
    TableRFP = Main.GetContentTablename(ContentNameSessionSubmissions)
    TableRules = Main.GetContentTablename(ContentNameSessionSubmissionsRules)
    TableType = Main.GetContentTablename(ContentNameSessionType)
    TablePeople = Main.GetContentTablename(ContentNamePeople)
    '
    TableList = TableRFP
    FieldList = ","
    '
    If Main.GetStreamBoolean(RequestNameSessionTitle) Then
        FieldList = FieldList & TableRFP & ".SessionTitle AS Title"
        End If
    If Main.GetStreamBoolean(RequestNameDescription) Then
        FieldList = FieldList & "," & TableRFP & ".Description AS Description"
        End If
    If Main.GetStreamBoolean(RequestNameEditedDescription) Then
        FieldList = FieldList & "," & TableRFP & ".EditedDescription AS EditedDescription"
        End If
    If Main.GetStreamBoolean(RequestNameTypeID) Then
        FieldList = FieldList & "," & TableType & ".Name AS SessionType"
        TypeFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameOutcome) Then
        FieldList = FieldList & "," & TableRFP & ".Outcome AS Outcome"
        End If
    If Main.GetStreamBoolean(RequestNameAidDescription) Then
        FieldList = FieldList & "," & TableRFP & ".AidDescription AS AidDescription"
        End If
    If Main.GetStreamBoolean(RequestNameAudiovisualNotes) Then
        FieldList = FieldList & "," & TableRFP & ".AudiovisualNotes AS AudiovisualNotes"
        End If
    If Main.GetStreamBoolean(RequestNameFirstName) Then
        FieldList = FieldList & "," & TablePeople & ".FirstName AS PrimaryFirstName"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameLastName) Then
        FieldList = FieldList & "," & TablePeople & ".LastName AS PrimaryLastName"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameTitle) Then
        FieldList = FieldList & "," & TablePeople & ".Title AS PrimaryTitle"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameCompany) Then
        FieldList = FieldList & "," & TablePeople & ".Company AS PrimaryCompany"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameAddress) Then
        FieldList = FieldList & "," & TablePeople & ".Address AS PrimaryAddress"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameAddress2) Then
        FieldList = FieldList & "," & TablePeople & ".Address2 AS PrimaryAddress2"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameCity) Then
        FieldList = FieldList & "," & TablePeople & ".City AS PrimaryCity"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameState) Then
        FieldList = FieldList & "," & TablePeople & ".State AS PrimaryState"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameZip) Then
        FieldList = FieldList & "," & TablePeople & ".Zip AS PrimaryZip"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNamePhone) Then
        FieldList = FieldList & "," & TablePeople & ".Phone AS PrimaryPhone"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameFax) Then
        FieldList = FieldList & "," & TablePeople & ".Fax AS PrimaryFax"
        PeopleFlag = True
        End If
    If Main.GetStreamBoolean(RequestNameEmail) Then
        FieldList = FieldList & "," & TablePeople & ".eMail AS PrimaryEmail"
        PeopleFlag = True
        End If
    '
    Stream = "SELECT * "
    Stream = Stream & " FROM " & TableList
    If Main.GetStreamBoolean(RequestNameAccepted) Then
        WhereClause = WhereClause & " and (" & TableRFP & ".Accepted)"
        End If
    Stream = Stream & " WHERE " & WhereClause
    '
    GetQuery = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetQuery", Err.Number, Err.Source, Err.Description, True, False)
    End Function
