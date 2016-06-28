VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RFPSubmissionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameFormID = "formid"
Const RequestNameSourceFormID = "sourceformid"
Const RequestNameButtonValue = "button"
Const RequestNameSessionTrackID = "TrackID"
Const RequestNameSerial = "frmsrl"
Const RequestNamePresenter = "psntr"
'
Const ContentNamePeople = "People"
Const ContentNameSessionType = "Session Types"
Const ContentNameSessionSubmissions = "Session Submissions"
Const ContentNameAVRules = "AV Submissions Rules"
Const ContentNameSessionSubmissionsRules = "Session Submissions Rules"
Const ContentNameContactRules = "Session Submissions Contact Rules"
Const ContentNameSessionTracks = "Session Tracks"
Const RequestNameEmailFlag = "emlflg"
'
Const FormRequest = 100
Const FormParticipants = 200
Const FormThankYou = 300
Const FormContacts = 400
Const FormLogin = 500
Const FormReview = 600
'
Const FormButtonLogin = "  Continue with Login "
Const FormButtonContinue = "  Continue  "
Const FormButtonComplete = "  Complete  "
Const FormButtonAddParticipant = "  Add Participant  "
Const FormButtonAddContact = "  Add Contact  "
Const FormButtonPrevious = "  Previous Step  "
Const FormButtonPassword = "  Send Password  "
'
Const FormSpacer = "<img src=""/ccLib/images/spacer.gif"" width=150 height=1>"
Const ButtonSpacer = "&nbsp;&nbsp;&nbsp;"
Const Required = "<span class=""ccError"">*</span> "
Const ContactTable = "cntct"
'
Const TableIDParticipants = "participants"
'
Const GroupNameSessionSubmitters = "Session Submissions"
'
Const SystemEmailAutoResponder = "Session Submission Auto Responder"
Const SystemEmailNotification = "Session Submission Notification"
'
Const PageNameFormThankYou = "RFP Submission Thank You"
Const PageNameFormRequest = "RFP Submission Instructions"
Const PageNameFormParticipants = "RFP Participant Instructions"
Const PageNameFormContacts = "RFP Contact Instructions"
Const PageNameFormReview = "RFP Review Instructions"
Const PageNamePrimaryPresenter = "Primary Presenter Instructions"
'
Const RequestNameSessionTitle = "sessiontitle"
Const RequestNameDescription = "description"
Const RequestNameTypeID = "typeid"
Const RequestNameOutcome = "outcome"
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameTitle = "title"
Const RequestNameCompany = "company"
Const RequestNameAddress = "address_"
Const RequestNameAddress2 = "address2_"
Const RequestNameCity = "city"
Const RequestNameState = "state"
Const RequestNameZip = "zip"
Const RequestNamePhone = "phone"
Const RequestNameFax = "fax"
Const RequestNameEmail = "email"

Const RequestNameRepFirstName = "repfirstname"
Const RequestNameRepLastName = "replastname"
Const RequestNameRepTitle = "reptitle"
Const RequestNameRepCompany = "vcompany"
Const RequestNameRepAddress = "repaddress_"
Const RequestNameRepAddress2 = "repaddress2_"
Const RequestNameRepCity = "repcity"
Const RequestNameRepState = "repstate"
Const RequestNameRepZip = "repzip"
Const RequestNameRepPhone = "repphone"
Const RequestNameRepFax = "repfax"
Const RequestNameRepEmail = "repemail"

Const RequestNameAidDescription = "aiddescription"
Const RequestNamePowerPoint = "ppaid"
Const RequestNameVideo = "vidaid"
Const RequestNameConnection = "connaid"
Const RequestNameAV = "reqav_"
Const RequestNameBoardLetter = "brdlettr"
Const RequestNameUsername = "uname"
Const RequestNamePassword = "pword"
Const RequestNameContactID = "cntid"
Const RequestNameRowCount = "rwcnt"
Const RequestNameDeleteFile = "dltfl"
Const RequestNameExistingFile = "exstfl"
Const RequestNameDeleteRow = "dltrw"
'
Const RequestNameParticipantCount = "participantcount"
Const RequestNameContactCount = "contactcount"
Const RequestNameSubmissionID = "submissionid"
'
Private FormID As Long
Private SourceFormID As Long
Private ButtonValue As String
'
Private SessionTitle As String
Private Description As String
Private TypeID As Long
Private Outcome As String
Private FirstName As String
Private LastName As String
Private Title As String
Private Company As String
Private Address As String
Private Address2 As String
Private City As String
Private State As String
Private Zip As String
Private Phone As String
Private Fax As String
Private eMail As String
'
Private RepFirstName As String
Private RepLastName As String
Private RepTitle As String
Private RepCompany As String
Private RepAddress As String
Private RepAddress2 As String
Private RepCity As String
Private RepState As String
Private RepZip As String
Private RepPhone As String
Private RepFax As String
Private RepeMail As String
'
Private AidDescription As String
Private SubmissionID As Long
Private ParticipantCount As Long
Private ContactCount As Long
Private TrackID As Long
Private PowerPoint As Boolean
Private Video As Boolean
Private Connection As Boolean
Private BoardLetter As String
Private Username As String
Private Password As String
'
Private FormSerialStream As String
Private FormSerial As String
'
Private RowCount As Long
Private EmailFlag As Boolean
'
Private ErrorString As String
Private WorkingQueryString As String
'
Private afterFormStuff As String
'
Private Main As Object
'
Public Sub Init(MainObject As Variant)
    '
    Set Main = MainObject
    '
    Call LoadForm
    Call ProcessForm
    '
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString <> "" Then
        WorkingQueryString = "?" & WorkingQueryString & "&"
    Else
        WorkingQueryString = "?"
        End If
    EmailFlag = Main.GetStreamBoolean(RequestNameEmailFlag)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadForm()
    On Error GoTo ErrorTrap
    '
    FormID = Main.GetStreamInteger(RequestNameFormID)
    SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    ButtonValue = Main.GetStreamText(RequestNameButtonValue)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "LoadForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    '
    FormSerialStream = Main.GetStreamText(RequestNameSerial)
    FormSerial = Main.GetVisitProperty("RFPAppFormSerial")
    '
    If FormSerial = FormSerialStream Then
        Select Case SourceFormID
            Case FormReview
                If ButtonValue = FormButtonPrevious Then
                    FormID = FormParticipants
                Else
                    Call LoadFormReview
                    Call ProcessFormReview
                End If
            Case FormRequest
                If ButtonValue = FormButtonPrevious Then
                    FormID = FormLogin
                Else
                    Call LoadFormRequest
                    Call ProcessFormRequest
                End If
            Case FormContacts
                If ButtonValue = FormButtonPrevious Then
                    FormID = FormRequest
                Else
                    Call LoadFormContacts
                    Call ProcessFormContacts
                End If
            Case FormParticipants
                If ButtonValue = FormButtonPrevious Then
                    FormID = FormContacts
                Else
                    Call LoadFormParticipants
                    Call ProcessFormParticipants
                End If
            Case FormLogin
                Call ProcessFormLogin
        End Select
    Else
        Select Case SourceFormID
            Case FormReview
                FormID = FormThankYou
            Case FormRequest
                FormID = FormContacts
            Case FormContacts
                FormID = FormParticipants
            Case FormParticipants
                FormID = FormReview
            Case FormLogin
                Call ProcessFormLogin
        End Select
    End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetForm() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Select Case FormID
        Case FormReview
            Call LoadFormReview
            Stream = Stream & GetFormReview
        Case FormThankYou
            Stream = Stream & Main.GetContentCopy(PageNameFormThankYou)
        Case FormParticipants
            Call LoadFormParticipants
            Stream = Stream & GetFormParticipants
        Case FormContacts
            Call LoadFormContacts
            Stream = Stream & GetFormContacts
        Case FormRequest
            Call LoadFormRequest
            Stream = Stream & GetFormRequest
        Case Else
            FormID = FormLogin
            Stream = Stream & GetFormLogin
    End Select
    '
    GetForm = Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetPageHead(PageName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD colspan=2 width=""100%"">"
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
    Call HandleError("RFPSubmissionClass", "GetPageHead", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRow(FieldCaption As String, Innards As String, Optional AlignLeft As Boolean, Optional alignTop As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim AlignmentString As String
    Dim vAlignmentString As String
    '
    If Not AlignLeft Then
        AlignmentString = " align=right"
    Else
        AlignmentString = " align=left"
    End If
    '
    If alignTop Then
        vAlignmentString = " valign=top"
    End If
    '
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD " & vAlignmentString & AlignmentString & ">" & FieldCaption & "</TD>"
    Stream = Stream & "<TD " & vAlignmentString & " width=""100%"">" & Innards & "</TD>"
    Stream = Stream & "</TR>"
    '
    GetFormRow = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFormRow", Err.Number, Err.Source, Err.Description, True, False)
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
    Call HandleError("RFPSubmissionClass", "GetFormRow2", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormRequest() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ButtonString As String
    '
    FormSerial = (Rnd() * 10000) * Main.VisitID
    Call Main.SetVisitProperty("RFPAppFormSerial", FormSerial)
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSerial, FormSerial)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonContinue)
    ButtonString = ButtonString & ButtonSpacer & Main.GetFormButton(FormButtonPrevious)
    '
    Stream = Stream & Main.GetUploadFormStart()
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & GetPageHead(PageNameFormRequest)
    Stream = Stream & GetFormRow(GetProperty("Suggested Session Title"), GetField(RequestNameSessionTitle, SessionTitle))
    Stream = Stream & GetFormRow(GetProperty("Session Description"), GetField(RequestNameDescription, Description, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Session Type"), Main.GetFormInputSelect(RequestNameTypeID, TypeID, ContentNameSessionType))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy("Tracks Note"))
    Stream = Stream & GetFormRow(GetProperty("Track"), Main.GetFormInputSelect(RequestNameSessionTrackID, TrackID, ContentNameSessionTracks))
    '
    Stream = Stream & GetFormRow(GetProperty("Anticipated Outcomes"), GetField(RequestNameOutcome, Outcome, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Board Support Letter"), GetFileRow(RequestNameBoardLetter, BoardLetter))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy("Audio Visual Aids Description"))
    Stream = Stream & GetFormRow(GetProperty("Audiovisual Aids"), GetAVFields())
    '
    Stream = Stream & GetFormRow(FormSpacer, ButtonString)
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy("ACCT Proposal Footer"))
    '
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    '
    GetFormRequest = Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFormRequest", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadFormRequest()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    If SourceFormID = FormRequest Then
        SessionTitle = Main.GetStreamText(RequestNameSessionTitle)
        Description = Main.GetStreamText(RequestNameDescription)
        TypeID = Main.GetStreamInteger(RequestNameTypeID)
        Outcome = Main.GetStreamText(RequestNameOutcome)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Title = Main.GetStreamText(RequestNameTitle)
        Company = Main.GetStreamText(RequestNameCompany)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        Fax = Main.GetStreamText(RequestNameFax)
        eMail = Main.GetStreamText(RequestNameEmail)
        AidDescription = Main.GetStreamText(RequestNameAidDescription)
        TrackID = Main.GetStreamInteger(RequestNameSessionTrackID)
        PowerPoint = Main.GetStreamBoolean(RequestNamePowerPoint)
        Video = Main.GetStreamBoolean(RequestNameVideo)
        Connection = Main.GetStreamBoolean(RequestNameConnection)
        BoardLetter = Main.GetStreamText(RequestNameBoardLetter)
    Else
        '
        SubmissionID = GetSubmissionID
        '
        CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "(ID=" & Main.EncodeSQLNumber(SubmissionID) & ") AND (MemberID=" & Main.MemberID & ")")
        If Main.CSOK(CSPointer) Then
            SessionTitle = Main.GetCSText(CSPointer, "SessionTitle")
            Description = Main.GetCSText(CSPointer, "Description")
            TypeID = Main.GetCSInteger(CSPointer, "TypeID")
            Outcome = Main.GetCSText(CSPointer, "Outcome")
            AidDescription = Main.GetCSText(CSPointer, "AidDescription")
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Title = Main.GetCSText(CSPointer, "Title")
            Company = Main.GetCSText(CSPointer, "Company")
            Address = Main.GetCSText(CSPointer, "Address")
            Address2 = Main.GetCSText(CSPointer, "Address2")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            Phone = Main.GetCSText(CSPointer, "Phone")
            Fax = Main.GetCSText(CSPointer, "Fax")
            eMail = Main.GetCSText(CSPointer, "Email")
            TrackID = Main.GetCSInteger(CSPointer, "TrackID")
            PowerPoint = Main.GetCSBoolean(CSPointer, "PowerPoint")
            Video = Main.GetCSBoolean(CSPointer, "Video")
            Connection = Main.GetCSBoolean(CSPointer, "Connection")
            BoardLetter = Main.GetCSText(CSPointer, "BoardLetter")
            End If
        Call Main.CloseCS(CSPointer)
        '
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "LoadFormRequest", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub LoadFormParticipants()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    If SourceFormID = FormParticipants Then
        SubmissionID = Main.GetStreamInteger(RequestNameSubmissionID)
        If SubmissionID = 0 Then
            SubmissionID = GetSubmissionID()
        End If
        RowCount = Main.GetStreamInteger(RequestNameRowCount)
        ParticipantCount = Main.GetStreamInteger(RequestNameParticipantCount)
    End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "LoadFormParticipants", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub LoadFormContacts()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    If SourceFormID = FormContacts Then
        SubmissionID = Main.GetStreamInteger(RequestNameSubmissionID)
        If SubmissionID = 0 Then
            SubmissionID = GetSubmissionID()
        End If
        RowCount = Main.GetStreamInteger(RequestNameRowCount)
        ContactCount = Main.GetStreamInteger(RequestNameContactCount)
        '
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Title = Main.GetStreamText(RequestNameTitle)
        Company = Main.GetStreamText(RequestNameCompany)
        Address = Main.GetStreamText(RequestNameAddress)
        Address2 = Main.GetStreamText(RequestNameAddress2)
        City = Main.GetStreamText(RequestNameCity)
        State = Main.GetStreamText(RequestNameState)
        Zip = Main.GetStreamText(RequestNameZip)
        Phone = Main.GetStreamText(RequestNamePhone)
        Fax = Main.GetStreamText(RequestNameFax)
        eMail = Main.GetStreamText(RequestNameEmail)
        Username = Main.GetStreamText(RequestNameUsername)
        Password = Main.GetStreamText(RequestNamePassword)
        '
    Else
        CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(Main.MemberID))
        If Main.CSOK(CSPointer) Then
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Title = Main.GetCSText(CSPointer, "Title")
            Company = Main.GetCSText(CSPointer, "Company")
            Address = Main.GetCSText(CSPointer, "Address")
            Address2 = Main.GetCSText(CSPointer, "Address2")
            City = Main.GetCSText(CSPointer, "City")
            State = Main.GetCSText(CSPointer, "State")
            Zip = Main.GetCSText(CSPointer, "Zip")
            Phone = Main.GetCSText(CSPointer, "Phone")
            Fax = Main.GetCSText(CSPointer, "Fax")
            eMail = Main.GetCSText(CSPointer, "Email")
            Username = Main.GetCSText(CSPointer, "Username")
            Password = Main.GetCSText(CSPointer, "Password")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "LoadFormContacts", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetProperty(PropertyName As String)
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
    Call HandleError("RFPSubmissionClass", "GetProperty", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetField(RequestName As String, Value As String, Optional Height As String, Optional Width As String, Optional MaxLenghth As String, Optional Password As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    If Height = "" Then
        Height = 1
        End If
    If Width = "" Then
        Width = 25
        End If
    '
    Stream = Main.GetFormInputText(RequestName, Value, Height, Width, , Password)
    Stream = Replace(Stream, "<INPUT ", "<INPUT maxlength=" & MaxLenghth & " ", 1, 99, 1)
    '
    GetField = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetField", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetSubmissionID() As Long
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Criteria As String
    Dim LocalSubmissionID As Long
    Dim CS As Long
    '
    LocalSubmissionID = Main.GetStreamInteger(RequestNameSubmissionID)
    '
    If LocalSubmissionID <> 0 Then
        GetSubmissionID = LocalSubmissionID
    Else
        '
        Criteria = "(MemberID=" & Main.EncodeSQLNumber(Main.MemberID) & ")"
        If Not Main.IsAuthenticated() Then
            Criteria = Criteria & " AND (VisitID=" & Main.EncodeSQLNumber(Main.VisitID) & ")"
            Criteria = Criteria & "AND (Completed=0)"
        Else
            Criteria = Criteria & "AND (Completed=1)"
        End If
        '
        CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, Criteria, , , , , "ID")
        If Main.CSOK(CSPointer) Then
            LocalSubmissionID = Main.GetCSInteger(CSPointer, "ID")
            GetSubmissionID = LocalSubmissionID
        Else
            Call Main.CloseCS(CSPointer)
            CSPointer = Main.InsertCSContent(ContentNameSessionSubmissions)
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "MemberID", Main.MemberID)
                Call Main.SetCS(CSPointer, "VisitID", Main.VisitID)
                Call Main.SetCS(CSPointer, "Completed", 0)
                Call Main.SaveCSRecord(CSPointer)
                LocalSubmissionID = Main.GetCSInteger(CSPointer, "ID")
                GetSubmissionID = LocalSubmissionID
                End If
            End If
        Call Main.CloseCS(CSPointer)
        '
    End If
    '
Call Main.TestPoint("GetSubmissionID: " & LocalSubmissionID)
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetField", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormRequest()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim CS As Long
    '
    If ButtonValue <> "" Then
        '
        SubmissionID = GetSubmissionID
        '
        If SessionTitle = "" Then
            ErrorString = ErrorString & "Please enter a session title in order to continue.<br />"
            End If
        If Description = "" Then
            ErrorString = ErrorString & "Please enter a session description in order to continue.<br />"
            End If
        If TypeID = 0 Then
            ErrorString = ErrorString & "Please enter a session type in order to continue.<br />"
            End If
        If Outcome = "" Then
            ErrorString = ErrorString & "Please enter an anticipated outcome in order to continue.<br />"
            End If
        If TrackID = 0 Then
            ErrorString = ErrorString & "Please enter the choice of track in order to continue. <br>"
        End If
        '
        If ErrorString = "" Then
            '
            CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "ID=" & Main.EncodeSQLNumber(SubmissionID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "Name", SessionTitle)
                Call Main.SetCS(CSPointer, "SessionTitle", SessionTitle)
                Call Main.SetCS(CSPointer, "Description", Description)
                Call Main.SetCS(CSPointer, "TypeID", TypeID)
                Call Main.SetCS(CSPointer, "Outcome", Outcome)
                Call Main.SetCS(CSPointer, "AidDescription", AidDescription)
                'Call Main.SetCS(CSPointer, "PowerPoint", PowerPoint)
                'Call Main.SetCS(CSPointer, "Video", Video)
                'Call Main.SetCS(CSPointer, "Connection", Connection)
                Call Main.SetCS(CSPointer, "TrackID", TrackID)
                If BoardLetter <> "" Then
                    Call Main.SetCSFormInput(CSPointer, "BoardLetter", RequestNameBoardLetter)
                    Call Main.SetCS(CSPointer, "BoardLetterUploaded", 1)
                Else
                    If Main.GetStreamBoolean(RequestNameDeleteFile & "_" & RequestNameBoardLetter) Then
                        Call Main.SetCS(CSPointer, "BoardLetter", "")
                        Call Main.SetCS(CSPointer, "BoardLetterUploaded", 0)
                    End If
                End If
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
            Call Main.DeleteContentRecords("AV Submission Rules", "SessionSubmissionsID=" & SubmissionID)
            '
            CSPointer = Main.OpenCSContent("AV Types", , , , , , "ID")
            Do While Main.CSOK(CSPointer)
                If Main.GetStreamBoolean(RequestNameAV & Main.GetCSInteger(CSPointer, "ID")) Then
                    CS = Main.InsertCSContent("AV Submission Rules")
                    If Main.CSOK(CS) Then
                        Call Main.SetCS(CS, "SessionSubmissionsID", SubmissionID)
                        Call Main.SetCS(CS, "avTypeID", Main.GetCSInteger(CSPointer, "ID"))
                    End If
                    Call Main.CloseCS(CS)
                End If
                Call Main.NextCSRecord(CSPointer)
            Loop
            Call Main.CloseCS(CSPointer)
            
            FormID = FormContacts
        Else
            FormID = FormRequest
            End If
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "ProcessFormRequest", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormParticipants() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CS As Long
    Dim LocalRowCount As Long
    Dim Counter As Long
    Dim ButtonString As String
    Dim Criteria As String
    Dim Flag As Boolean
    Dim LocalAddedCount As Long
    '
    Criteria = GetSubmissionCriteria(True)
    '
    Stream = Main.GetFormStart
    Stream = Stream & "<table name=""" & ContactTable & """ id=""" & ContactTable & """ border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">"
    Stream = Stream & GetPageHead(PageNameFormParticipants)
    '
    If SourceFormID = FormParticipants Then
        LocalRowCount = 1
        CS = Main.OpenCSContent(ContentNamePeople, Criteria)
        If Main.CSOK(CS) Then
            Flag = True
            Do While Main.CSOK(CS)
                Stream = Stream & GetFormRow2(Main.GetFormInputHidden(RequestNameContactID & LocalRowCount, Main.GetCSInteger(CS, "ID")) & "<b>Participant</b>")
                Stream = Stream & GetParticipantStreamFields(LocalRowCount)
                Call Main.NextCSRecord(CS)
                If Main.CSOK(CS) Then
                    LocalRowCount = LocalRowCount + 1
                End If
            Loop
            Call Main.CloseCS(CS)
        End If
        '
        LocalAddedCount = LocalRowCount
        '
        For Counter = 1 To ParticipantCount
            Stream = Stream & GetFormRow2("<b>Participant</b>")
            Stream = Stream & GetParticipantStreamFields(LocalAddedCount + Counter)
            LocalRowCount = LocalRowCount + 1
        Next
        '
    Else
        CS = Main.OpenCSContent(ContentNamePeople, Criteria)
        If Main.CSOK(CS) Then
            Flag = True
            LocalRowCount = 1
            Do While Main.CSOK(CS)
                Stream = Stream & GetFormRow2(Main.GetFormInputHidden(RequestNameContactID & LocalRowCount, Main.GetCSInteger(CS, "ID")) & "<b>Participant</b>")
                Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameDeleteRow & LocalRowCount), "Delete")
                Stream = Stream & GetFormRow2(FormSpacer)
                Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenter & LocalRowCount, Main.GetCSBoolean(CS, "Presenter")), "I am the moderator/facilitator for this session")
                Stream = Stream & GetFormRow(GetProperty("First Name"), GetField(RequestNameRepFirstName & LocalRowCount, Main.GetCSText(CS, "FirstName")))
                Stream = Stream & GetFormRow(GetProperty("Last Name"), GetField(RequestNameRepLastName & LocalRowCount, Main.GetCSText(CS, "LastName")))
                Stream = Stream & GetFormRow(GetProperty("Title"), GetField(RequestNameRepTitle & LocalRowCount, Main.GetCSText(CS, "Title")))
                Stream = Stream & GetFormRow(GetProperty("College/Organization"), GetField(RequestNameRepCompany & LocalRowCount, Main.GetCSText(CS, "Company")))
                Stream = Stream & GetFormRow(GetProperty("Address"), GetField(RequestNameRepAddress & LocalRowCount, Main.GetCSText(CS, "Address")))
                Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameRepAddress2 & LocalRowCount, Main.GetCSText(CS, "Address2")))
                Stream = Stream & GetFormRow(GetProperty("City"), GetField(RequestNameRepCity & LocalRowCount, Main.GetCSText(CS, "City")))
                Stream = Stream & GetFormRow(GetProperty("State"), GetField(RequestNameRepState & LocalRowCount, Main.GetCSText(CS, "State")))
                Stream = Stream & GetFormRow(GetProperty("Zip"), GetField(RequestNameRepZip & LocalRowCount, Main.GetCSText(CS, "Zip")))
                Stream = Stream & GetFormRow(GetProperty("Phone"), GetField(RequestNameRepPhone & LocalRowCount, Main.GetCSText(CS, "Phone")))
                Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameRepFax & LocalRowCount, Main.GetCSText(CS, "Fax")))
                Stream = Stream & GetFormRow(GetProperty("Email"), GetField(RequestNameRepEmail & LocalRowCount, Main.GetCSText(CS, "Email")))
                Call Main.NextCSRecord(CS)
                If Main.CSOK(CS) Then
                    LocalRowCount = LocalRowCount + 1
                End If
            Loop
        End If
        Call Main.CloseCS(CS)
    End If
    '
    If Not Flag Then
        Stream = Stream & GetFormRow2("<b>Participant</b>")
        Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenter & "1", ""), "I am the moderator/facilitator for this session")
        Stream = Stream & GetFormRow(GetProperty("First Name"), GetField(RequestNameRepFirstName & "1", RepFirstName))
        Stream = Stream & GetFormRow(GetProperty("Last Name"), GetField(RequestNameRepLastName & "1", RepLastName))
        Stream = Stream & GetFormRow(GetProperty("Title"), GetField(RequestNameRepTitle & "1", RepTitle))
        Stream = Stream & GetFormRow(GetProperty("College/Organization"), GetField(RequestNameRepCompany & "1", RepCompany))
        Stream = Stream & GetFormRow(GetProperty("Address"), GetField(RequestNameRepAddress & "1", RepAddress))
        Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameRepAddress2 & "1", RepAddress2))
        Stream = Stream & GetFormRow(GetProperty("City"), GetField(RequestNameRepCity & "1", RepCity))
        Stream = Stream & GetFormRow(GetProperty("State"), GetField(RequestNameRepState & "1", RepState))
        Stream = Stream & GetFormRow(GetProperty("Zip"), GetField(RequestNameRepZip & "1", RepZip))
        Stream = Stream & GetFormRow(GetProperty("Phone"), GetField(RequestNameRepPhone & "1", RepPhone))
        Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameRepFax & "1", RepFax))
        Stream = Stream & GetFormRow(GetProperty("Email"), GetField(RequestNameRepEmail & "1", RepeMail))
        LocalRowCount = 1
    End If
    '
    RowCount = LocalRowCount
    '
    FormSerial = (Rnd() * 10000) * Main.VisitID
    Call Main.SetVisitProperty("RFPAppFormSerial", FormSerial)
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSerial, FormSerial)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameRowCount, RowCount, RequestNameRowCount)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameParticipantCount, ParticipantCount, RequestNameParticipantCount)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSubmissionID, SubmissionID)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonContinue)
    ButtonString = ButtonString & "&nbsp;<input value=""" & FormButtonAddParticipant & """ onClick=""InsertParticipantRows(); return false;"" type=""button"">"
    ButtonString = ButtonString & ButtonSpacer & Main.GetFormButton(FormButtonPrevious)
    '
    Stream = Stream & GetFormRow(FormSpacer, ButtonString)
    Stream = Stream & "</table>"
    Stream = Stream & Main.GetFormEnd
    '
    GetFormParticipants = GetScripts & Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFormParticipants", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetParticipantScript() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<script language=""JavaScript1.2"">" & vbCrLf
    Stream = Stream & "function InsertParticipantRows() {" & vbCrLf
    Stream = Stream & "var RLTable = document.getElementById('" & TableIDParticipants & "' );" & vbCrLf
    Stream = Stream & "var CountElement = document.getElementById('" & RequestNameParticipantCount & "' );" & vbCrLf
    Stream = Stream & "var tCols,tRows,NewRowNumber;" & vbCrLf
    Stream = Stream & "var NewRow,NewCell1,NewCell2,NewCell3,NewCell4,NewCell5,NewCell6,NewCell7,NewCell8,NewCell9,NewCell10,NewCell11,NewCell12,NewCell13;" & vbCrLf
    Stream = Stream & "if ( RLTable ) {" & vbCrLf
    Stream = Stream & "NewRowNumber = parseInt( CountElement.value )+1;" & vbCrLf
    Stream = Stream & "CountElement.value = NewRowNumber;" & vbCrLf
    Stream = Stream & "tRows = RLTable.getElementsByTagName(""TR"");" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell1 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell1.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell1.colSpan = '2';" & vbCrLf
    Stream = Stream & "NewCell1.innerHTML= '<b>" & GetProperty("Panel Participant") & "&nbsp;'+Number(NewRowNumber+1);'</b>'" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell2 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell2.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell2.innerHTML= 'First Name:';" & vbCrLf
    Stream = Stream & "NewCell2 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell2.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell2.innerHTML= '" & Main.GetFormInputText(RequestNameFirstName & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameFirstName & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell3 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell3.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell3.innerHTML= 'Last Name:';" & vbCrLf
    Stream = Stream & "NewCell3 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell3.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell3.innerHTML= '" & Main.GetFormInputText(RequestNameLastName & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameLastName & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell4 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell4.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell4.innerHTML= 'Title:';" & vbCrLf
    Stream = Stream & "NewCell4 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell4.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell4.innerHTML= '" & Main.GetFormInputText(RequestNameTitle & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameTitle & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell5 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell5.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell5.innerHTML= 'College/Organization:';" & vbCrLf
    Stream = Stream & "NewCell5 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell5.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell5.innerHTML= '" & Main.GetFormInputText(RequestNameCompany & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameCompany & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell6 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell6.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell6.innerHTML= 'Address:';" & vbCrLf
    Stream = Stream & "NewCell6 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell6.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell6.innerHTML= '" & Main.GetFormInputText(RequestNameAddress & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameAddress & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell7 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell7.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell7.innerHTML= '';" & vbCrLf
    Stream = Stream & "NewCell7 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell7.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell7.innerHTML= '" & Main.GetFormInputText(RequestNameAddress2 & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameAddress2 & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell8 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell8.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell8.innerHTML= 'City:';" & vbCrLf
    Stream = Stream & "NewCell8 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell8.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell8.innerHTML= '" & Main.GetFormInputText(RequestNameCity & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameCity & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell9 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell9.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell9.innerHTML= 'State:';" & vbCrLf
    Stream = Stream & "NewCell9 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell9.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell9.innerHTML= '" & Main.GetFormInputText(RequestNameState & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameState & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell10 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell10.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell10.innerHTML= 'Zip:';" & vbCrLf
    Stream = Stream & "NewCell10 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell10.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell10.innerHTML= '" & Main.GetFormInputText(RequestNameZip & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameZip & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell11 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell11.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell11.innerHTML= 'Phone:';" & vbCrLf
    Stream = Stream & "NewCell11 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell11.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell11.innerHTML= '" & Main.GetFormInputText(RequestNamePhone & "_'+NewRowNumber+'", Main.GetStreamText(RequestNamePhone & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell12 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell12.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell12.innerHTML= 'Fax:';" & vbCrLf
    Stream = Stream & "NewCell12 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell12.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell12.innerHTML= '" & Main.GetFormInputText(RequestNameFax & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameFax & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "NewRow = RLTable.insertRow();" & vbCrLf
    Stream = Stream & "NewCell13 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell13.align = 'right';" & vbCrLf
    Stream = Stream & "NewCell13.innerHTML= 'Email:';" & vbCrLf
    Stream = Stream & "NewCell13 = NewRow.insertCell();" & vbCrLf
    Stream = Stream & "NewCell13.align = 'left';" & vbCrLf
    Stream = Stream & "NewCell13.innerHTML= '" & Main.GetFormInputText(RequestNameEmail & "_'+NewRowNumber+'", Main.GetStreamText(RequestNameEmail & "_'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "}" & vbCrLf
    Stream = Stream & "}" & vbCrLf
    Stream = Stream & "</script>" & vbCrLf
    '
    GetParticipantScript = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetParticipantScript", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormParticipantsOLD()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim ParticipantPointer As Long
    Dim LocalMemberID As Long
    Dim LocalEmail As String
    Dim LocalFirstName As String
    Dim LocalLastName As String
    Dim Copy As String
    Dim LocalQueryString As String
    '
    If ButtonValue <> "" Then
        '
        For ParticipantPointer = 0 To ParticipantCount
            '
            LocalEmail = Main.GetStreamText(RequestNameFirstName & "_" & ParticipantPointer)
            LocalFirstName = Main.GetStreamText(RequestNameFirstName & "_" & ParticipantPointer)
            LocalLastName = Main.GetStreamText(RequestNameLastName & "_" & ParticipantPointer)
            '
            If (LocalEmail <> "") And (LocalFirstName <> "") And (LocalLastName <> "") Then
                '
                If LocalEmail <> "" Then
                    CSPointer = Main.OpenCSContent(ContentNamePeople, "Email=" & Main.EncodeSQLText(LocalEmail), , , "ID")
                    If Main.CSOK(CSPointer) Then
                        LocalMemberID = Main.GetCSInteger(CSPointer, "ID")
                        End If
                    Call Main.CloseCS(CSPointer)
                    End If
                '
                If LocalMemberID = 0 Then
                    CSPointer = Main.InsertCSContent(ContentNamePeople)
                    If Main.CSOK(CSPointer) Then
                        Call Main.SetCS(CSPointer, "Name", LocalFirstName & " " & LocalLastName)
                        Call Main.SetCS(CSPointer, "FirstName", LocalFirstName)
                        Call Main.SetCS(CSPointer, "LastName", LocalLastName)
                        Call Main.SetCS(CSPointer, "Title", Main.GetStreamText(RequestNameTitle & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "Company", Main.GetStreamText(RequestNameCompany & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "Address", Main.GetStreamText(RequestNameAddress & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "Address2", Main.GetStreamText(RequestNameAddress2 & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "City", Main.GetStreamText(RequestNameCity & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "State", Main.GetStreamText(RequestNameState & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "Zip", Main.GetStreamText(RequestNameZip & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "Phone", Main.GetStreamText(RequestNamePhone & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "Fax", Main.GetStreamText(RequestNameFax & "_" & ParticipantPointer))
                        Call Main.SetCS(CSPointer, "Email", Main.GetStreamText(RequestNameEmail & "_" & ParticipantPointer))
                        Call Main.SaveCSRecord(CSPointer)
                        LocalMemberID = Main.GetCSInteger(CSPointer, "ID")
                        End If
                    Call Main.CloseCS(CSPointer)
                    End If
                '
                CSPointer = Main.InsertCSContent(ContentNameSessionSubmissionsRules)
                If Main.CSOK(CSPointer) Then
                    Call Main.SetCS(CSPointer, "SessionSubmissionsID", SubmissionID)
                    Call Main.SetCS(CSPointer, "MemberID", LocalMemberID)
                    Call Main.SaveCSRecord(CSPointer)
                    End If
                Call Main.CloseCS(CSPointer)
                '
                LocalMemberID = 0
                LocalEmail = ""
                LocalFirstName = ""
                LocalLastName = ""
                '
                End If
            '
        Next
        '
        CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "ID=" & Main.EncodeSQLNumber(SubmissionID))
        If Main.CSOK(CSPointer) Then
            Call Main.SetCS(CSPointer, "Completed", 1)
            Call Main.SaveCSRecord(CSPointer)
            End If
        Call Main.CloseCS(CSPointer)
        '
        LocalQueryString = "cid=" & Main.GetContentID(ContentNameSessionSubmissions)
        LocalQueryString = LocalQueryString & "&id=" & SubmissionID
        LocalQueryString = LocalQueryString & "&af=4"
        '
        Copy = "<br /><br /><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?" & LocalQueryString & """>Click here to view the submission</a>"
        '
        Call Main.AddGroupMember(GroupNameSessionSubmitters)
        '
        Call Main.SendSystemEmail(SystemEmailAutoResponder, , Main.MemberID)
        Call Main.SendSystemEmail(SystemEmailNotification, Copy)
        '
        FormID = FormThankYou
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "ProcessFormParticipants", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormRequestOLD() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ButtonString As String
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonContinue)
    '
    Stream = Stream & Main.GetFormStart
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & GetPageHead(PageNameFormRequest)
    Stream = Stream & GetFormRow(GetProperty("Suggested Session Title"), GetField(RequestNameSessionTitle, SessionTitle))
    Stream = Stream & GetFormRow(GetProperty("Session Description"), GetField(RequestNameDescription, Description, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Session Type"), Main.GetFormInputSelect(RequestNameTypeID, TypeID, ContentNameSessionType))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy("Tracks Note"))
    Stream = Stream & GetFormRow(GetProperty("Track"), Main.GetFormInputSelect(RequestNameSessionTrackID, TrackID, ContentNameSessionTracks))
    '
    Stream = Stream & GetFormRow(GetProperty("Anticipated Outcomes"), GetField(RequestNameOutcome, Outcome, 10, 40))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy(PageNamePrimaryPresenter))
    '
    Stream = Stream & GetFormRow(GetProperty("First Name"), GetField(RequestNameFirstName, FirstName))
    Stream = Stream & GetFormRow(GetProperty("Last Name"), GetField(RequestNameLastName, LastName))
    Stream = Stream & GetFormRow(GetProperty("Title"), GetField(RequestNameTitle, Title))
    Stream = Stream & GetFormRow(GetProperty("College/Organization"), GetField(RequestNameCompany, Company))
    Stream = Stream & GetFormRow(GetProperty("Address"), GetField(RequestNameAddress, Address))
    Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameAddress2, Address2))
    Stream = Stream & GetFormRow(GetProperty("City"), GetField(RequestNameCity, City))
    Stream = Stream & GetFormRow(GetProperty("State"), GetField(RequestNameState, State))
    Stream = Stream & GetFormRow(GetProperty("Zip"), GetField(RequestNameZip, Zip))
    Stream = Stream & GetFormRow(GetProperty("Phone"), GetField(RequestNamePhone, Phone))
    Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameFax, Fax))
    Stream = Stream & GetFormRow(GetProperty("Email"), GetField(RequestNameEmail, eMail))
    Stream = Stream & GetFormRow(GetProperty("Audiovisual Aids"), GetField(RequestNameAidDescription, AidDescription, 10, 40))
    '
    Stream = Stream & GetFormRow(FormSpacer, ButtonString)
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy("ACCT Proposal Footer"))
    '
        If FirstName = "" Then
            ErrorString = ErrorString & "Please enter your first name in order to continue.<br />"
            End If
        If LastName = "" Then
            ErrorString = ErrorString & "Please enter your last name in order to continue.<br />"
            End If
        If Company = "" Then
            ErrorString = ErrorString & "Please enter your College/Organization in order to continue.<br />"
            End If
        If Address = "" Then
            ErrorString = ErrorString & "Please enter your address in order to continue.<br />"
            End If
        If City = "" Then
            ErrorString = ErrorString & "Please enter your city in order to continue.<br />"
            End If
        If State = "" Then
            ErrorString = ErrorString & "Please enter your state in order to continue.<br />"
            End If
        If Zip = "" Then
            ErrorString = ErrorString & "Please enter your zip in order to continue.<br />"
            End If
        If Phone = "" Then
            ErrorString = ErrorString & "Please enter your phone number in order to continue.<br />"
            End If
        If eMail = "" Then
            ErrorString = ErrorString & "Please enter your email address in order to continue.<br />"
            End If
                'Call Main.SetCS(CSPointer, "FirstName", FirstName)
'                Call Main.SetCS(CSPointer, "LastName", LastName)
'                Call Main.SetCS(CSPointer, "Title", Title)
'                Call Main.SetCS(CSPointer, "Company", Company)
'                Call Main.SetCS(CSPointer, "Address", Address)
'                Call Main.SetCS(CSPointer, "Address2", Address2)
'                Call Main.SetCS(CSPointer, "City", City)
'                Call Main.SetCS(CSPointer, "State", State)
'                Call Main.SetCS(CSPointer, "Zip", Zip)
'                Call Main.SetCS(CSPointer, "Phone", Phone)
'                Call Main.SetCS(CSPointer, "Fax", Fax)
'                Call Main.SetCS(CSPointer, "Email", eMail)
'
'        CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(Main.MemberID))
'        If Main.CSOK(CSPointer) Then
'            If FirstName = "" Then
'                FirstName = Main.GetCSText(CSPointer, "FirstName")
'                End If
'            If LastName = "" Then
'                LastName = Main.GetCSText(CSPointer, "LastName")
'                End If
'            If LastName = "" Then
'                Title = Main.GetCSText(CSPointer, "Title")
'                End If
'            If Company = "" Then
'                Company = Main.GetCSText(CSPointer, "Company")
'                End If
'            If Address = "" Then
'                Address = Main.GetCSText(CSPointer, "Address")
'                End If
'            If Address2 = "" Then
'                Address2 = Main.GetCSText(CSPointer, "Address2")
'                End If
'            If City = "" Then
'                City = Main.GetCSText(CSPointer, "City")
'                End If
'            If State = "" Then
'            State = Main.GetCSText(CSPointer, "State")
'                End If
'            If Zip = "" Then
'                Zip = Main.GetCSText(CSPointer, "Zip")
'                End If
'            If Phone = "" Then
'                Phone = Main.GetCSText(CSPointer, "Phone")
'                End If
'            If Fax = "" Then
'                Fax = Main.GetCSText(CSPointer, "Fax")
'                End If
'            If eMail = "" Then
'                eMail = Main.GetCSText(CSPointer, "Email")
'                End If
'            End If
'        Call Main.CloseCS(CSPointer)
    '
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd
    '
    GetFormRequestOLD = Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFormRequestOLD", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetAVFields() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim innerString As String
    Dim CS As Long
    Dim boxSelected As Boolean
    '
    CSPointer = Main.OpenCSContent("AV Types")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<table width=""100%"" cellpadding=""3"" cellspacing=""0"" border=""0"">"
        Do While Main.CSOK(CSPointer)
            innerString = "<b>" & Main.GetCS(CSPointer, "Name") & " - " & FormatCurrency(Main.GetCSNumber(CSPointer, "price"), 2) & "</b><br />"
            innerString = innerString & Main.GetCS(CSPointer, "description")
            CS = Main.OpenCSContent("AV Submission Rules", "(SessionSubmissionsID=" & SubmissionID & ") and (avtypeID=" & Main.GetCSInteger(CSPointer, "ID") & ")")
            boxSelected = Main.CSOK(CS)
            Call Main.CloseCS(CS)
            Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameAV & Main.GetCS(CSPointer, "ID"), boxSelected), innerString, , True)
            Main.NextCSRecord (CSPointer)
        Loop
        Stream = Stream & "</table>"
    End If
    Call Main.CloseCS(CSPointer)
    '
    GetAVFields = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetAVFields", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormContacts2()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim ContactPointer As Long
    Dim LocalMemberID As Long
    Dim LocalEmail As String
    Dim LocalFirstName As String
    Dim LocalLastName As String
    Dim Copy As String
    Dim LocalQueryString As String
    '
    If ButtonValue <> "" Then
        '
        For ContactPointer = 0 To ContactCount
            '
            LocalEmail = Main.GetStreamText(RequestNameFirstName & "_" & ContactPointer)
            LocalFirstName = Main.GetStreamText(RequestNameFirstName & "_" & ContactPointer)
            LocalLastName = Main.GetStreamText(RequestNameLastName & "_" & ContactPointer)
            '
            If (LocalEmail <> "") And (LocalFirstName <> "") And (LocalLastName <> "") Then
                '
                If LocalEmail <> "" Then
                    CSPointer = Main.OpenCSContent(ContentNamePeople, "Email=" & Main.EncodeSQLText(LocalEmail), , , "ID")
                    If Main.CSOK(CSPointer) Then
                        LocalMemberID = Main.GetCSInteger(CSPointer, "ID")
                        End If
                    Call Main.CloseCS(CSPointer)
                    End If
                '
                If LocalMemberID = 0 Then
                    CSPointer = Main.InsertCSContent(ContentNamePeople)
                    If Main.CSOK(CSPointer) Then
                        Call Main.SetCS(CSPointer, "Name", LocalFirstName & " " & LocalLastName)
                        Call Main.SetCS(CSPointer, "FirstName", LocalFirstName)
                        Call Main.SetCS(CSPointer, "LastName", LocalLastName)
                        Call Main.SetCS(CSPointer, "Title", Main.GetStreamText(RequestNameTitle & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "Company", Main.GetStreamText(RequestNameCompany & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "Address", Main.GetStreamText(RequestNameAddress & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "Address2", Main.GetStreamText(RequestNameAddress2 & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "City", Main.GetStreamText(RequestNameCity & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "State", Main.GetStreamText(RequestNameState & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "Zip", Main.GetStreamText(RequestNameZip & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "Phone", Main.GetStreamText(RequestNamePhone & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "Fax", Main.GetStreamText(RequestNameFax & "_" & ContactPointer))
                        Call Main.SetCS(CSPointer, "Email", Main.GetStreamText(RequestNameEmail & "_" & ContactPointer))
                        Call Main.SaveCSRecord(CSPointer)
                        LocalMemberID = Main.GetCSInteger(CSPointer, "ID")
                        End If
                    Call Main.CloseCS(CSPointer)
                    End If
                '
                CSPointer = Main.InsertCSContent(ContentNameContactRules)
                If Main.CSOK(CSPointer) Then
                    Call Main.SetCS(CSPointer, "SessionSubmissionsID", SubmissionID)
                    Call Main.SetCS(CSPointer, "MemberID", LocalMemberID)
                    Call Main.SaveCSRecord(CSPointer)
                    End If
                Call Main.CloseCS(CSPointer)
                '
                LocalMemberID = 0
                LocalEmail = ""
                LocalFirstName = ""
                LocalLastName = ""
                '
                End If
            '
        Next
        '
        CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "ID=" & Main.EncodeSQLNumber(SubmissionID))
        If Main.CSOK(CSPointer) Then
            Call Main.SetCS(CSPointer, "Completed", 1)
            Call Main.SaveCSRecord(CSPointer)
            End If
        Call Main.CloseCS(CSPointer)
        '
        LocalQueryString = "cid=" & Main.GetContentID(ContentNameSessionSubmissions)
        LocalQueryString = LocalQueryString & "&id=" & SubmissionID
        LocalQueryString = LocalQueryString & "&af=4"
        '
        Copy = "<br /><br /><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?" & LocalQueryString & """>Click here to view the submission</a>"
        '
        Call Main.AddGroupMember(GroupNameSessionSubmitters)
        '
        Call Main.SendSystemEmail(SystemEmailAutoResponder, , Main.MemberID)
        Call Main.SendSystemEmail(SystemEmailNotification, Copy)
        '
        FormID = FormThankYou
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "ProcessFormContacts", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormContacts() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CS As Long
    Dim LocalRowCount As Long
    Dim Counter As Long
    Dim ButtonString As String
    Dim Criteria As String
    Dim LocalAddedCount As Long
    '
    Criteria = GetSubmissionCriteria()
    '
    Stream = Main.GetFormStart
    Stream = Stream & "<table name=""" & ContactTable & """ id=""" & ContactTable & """ border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">"
    Stream = Stream & GetPageHead(PageNameFormContacts)
    Stream = Stream & GetFormRow2("<b>Primary Contact</b>")
    Stream = Stream & GetFormRow(Required & GetProperty("First Name"), GetField(RequestNameFirstName, FirstName))
    Stream = Stream & GetFormRow(Required & GetProperty("Last Name"), GetField(RequestNameLastName, LastName))
    Stream = Stream & GetFormRow(Required & GetProperty("Title"), GetField(RequestNameTitle, Title))
    Stream = Stream & GetFormRow(Required & GetProperty("College/Organization"), GetField(RequestNameCompany, Company))
    Stream = Stream & GetFormRow(Required & GetProperty("Address"), GetField(RequestNameAddress, Address))
    Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameAddress2, Address2))
    Stream = Stream & GetFormRow(Required & GetProperty("City"), GetField(RequestNameCity, City))
    Stream = Stream & GetFormRow(Required & GetProperty("State"), GetField(RequestNameState, State))
    Stream = Stream & GetFormRow(Required & GetProperty("Zip"), GetField(RequestNameZip, Zip))
    Stream = Stream & GetFormRow(Required & GetProperty("Phone"), GetField(RequestNamePhone, Phone))
    Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameFax, Fax))
    Stream = Stream & GetFormRow(Required & GetProperty("Email"), GetField(RequestNameEmail, eMail))
    Stream = Stream & GetFormRow(Required & "Username:", GetField(RequestNameUsername, Username))
    Stream = Stream & GetFormRow(Required & "Password:", GetField(RequestNamePassword, Password))
    '
    If SourceFormID = FormContacts Then
        CS = Main.OpenCSContent(ContentNamePeople, "(ID<>" & Main.MemberID & ") AND (" & Criteria & ")")
        If Main.CSOK(CS) Then
            LocalRowCount = 1
            Do While Main.CSOK(CS)
                Stream = Stream & GetFormRow2(Main.GetFormInputHidden(RequestNameContactID & LocalRowCount, Main.GetCSInteger(CS, "ID")) & "<b>Alternate Contact</b>")
                Stream = Stream & GetContactStreamFields(LocalRowCount)
                Call Main.NextCSRecord(CS)
                If Main.CSOK(CS) Then
                    LocalRowCount = LocalRowCount + 1
                End If
            Loop
            Call Main.CloseCS(CS)
        End If
        '
        LocalAddedCount = LocalRowCount
        '
        For Counter = 1 To ContactCount
            Stream = Stream & GetFormRow2("<b>Alternate Contact</b>")
            Stream = Stream & GetContactStreamFields(LocalAddedCount + Counter)
            LocalRowCount = LocalRowCount + 1
        Next
        '
    Else
        CS = Main.OpenCSContent(ContentNamePeople, "(ID<>" & Main.MemberID & ") AND " & Criteria)
        If Main.CSOK(CS) Then
            LocalRowCount = 1
            Do While Main.CSOK(CS)
                Stream = Stream & GetFormRow2(Main.GetFormInputHidden(RequestNameContactID & LocalRowCount, Main.GetCSInteger(CS, "ID")) & "<b>Alternate Contact</b>")
                Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNameDeleteRow & LocalRowCount), "Delete")
                Stream = Stream & GetFormRow(Required & GetProperty("First Name"), GetField(RequestNameFirstName & LocalRowCount, Main.GetCSText(CS, "FirstName")))
                Stream = Stream & GetFormRow(Required & GetProperty("Last Name"), GetField(RequestNameLastName & LocalRowCount, Main.GetCSText(CS, "LastName")))
                Stream = Stream & GetFormRow(Required & GetProperty("Title"), GetField(RequestNameTitle & LocalRowCount, Main.GetCSText(CS, "Title")))
                Stream = Stream & GetFormRow(Required & GetProperty("College/Organization"), GetField(RequestNameCompany & LocalRowCount, Main.GetCSText(CS, "Company")))
                Stream = Stream & GetFormRow(Required & GetProperty("Address"), GetField(RequestNameAddress & LocalRowCount, Main.GetCSText(CS, "Address")))
                Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameAddress2 & LocalRowCount, Main.GetCSText(CS, "Address2")))
                Stream = Stream & GetFormRow(Required & GetProperty("City"), GetField(RequestNameCity & LocalRowCount, Main.GetCSText(CS, "City")))
                Stream = Stream & GetFormRow(Required & GetProperty("State"), GetField(RequestNameState & LocalRowCount, Main.GetCSText(CS, "State")))
                Stream = Stream & GetFormRow(Required & GetProperty("Zip"), GetField(RequestNameZip & LocalRowCount, Main.GetCSText(CS, "Zip")))
                Stream = Stream & GetFormRow(Required & GetProperty("Phone"), GetField(RequestNamePhone & LocalRowCount, Main.GetCSText(CS, "Phone")))
                Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameFax & LocalRowCount, Main.GetCSText(CS, "Fax")))
                Stream = Stream & GetFormRow(Required & GetProperty("Email"), GetField(RequestNameEmail & LocalRowCount, Main.GetCSText(CS, "Email")))
                Stream = Stream & GetFormRow(Required & "Username:", GetField(RequestNameUsername & LocalRowCount, Main.GetCSText(CS, "Username")))
                Stream = Stream & GetFormRow(Required & "Password:", GetField(RequestNamePassword & LocalRowCount, Main.GetCSText(CS, "Password")))
                Call Main.NextCSRecord(CS)
                If Main.CSOK(CS) Then
                    LocalRowCount = LocalRowCount + 1
                End If
            Loop
        End If
        Call Main.CloseCS(CS)
    End If
    '
    RowCount = LocalRowCount
    '
    FormSerial = (Rnd() * 10000) * Main.VisitID
    Call Main.SetVisitProperty("RFPAppFormSerial", FormSerial)
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSerial, FormSerial)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameRowCount, RowCount, RequestNameRowCount)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameContactCount, ContactCount, RequestNameContactCount)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSubmissionID, SubmissionID)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonContinue)
    ButtonString = ButtonString & "&nbsp;<input value=""" & FormButtonAddContact & """ onClick=""InsertContactRows(); return false;"" type=""button"">"
    ButtonString = ButtonString & ButtonSpacer & Main.GetFormButton(FormButtonPrevious)
    '
    Stream = Stream & GetFormRow(FormSpacer, ButtonString)
    Stream = Stream & "</table>"
    Stream = Stream & Main.GetFormEnd
    '
    GetFormContacts = GetScripts & Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFormContacts", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'
Private Function GetScripts() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<script language=""JavaScript1.2"">" & vbCrLf
    '
    Stream = Stream & "function InsertContactRows() {"
    Stream = Stream & "    var RLTable = document.getElementById('" & ContactTable & "');"
    Stream = Stream & "    var CountElement = document.getElementById('" & RequestNameRowCount & "');"
    Stream = Stream & "    var tCols,tRows,NewRowNumber;"
    Stream = Stream & "    var insrtcnt = document.getElementById('" & RequestNameContactCount & "');;"
    Stream = Stream & "    var NewRow0, NewRow, NewRow2, NewRow3, NewRow4, NewRow5, NewRow6, NewRow7;"
    Stream = Stream & "    var NewRow8, NewRow9, NewRow10, NewRow11, NewRow12, NewRow13, NewRow14;"
    Stream = Stream & "    var NewCell1,NewCell2,NewCell3,NewCell4,NewCell5,NewCell6;"
    Stream = Stream & "    var NewCell7,NewCell8,NewCell9,NewCell10,NewCell11,NewCell12;"
    Stream = Stream & "    var NewCell3,NewCell4,NewCell4,NewCell4,NewCell4,NewCell8;"
    Stream = Stream & "    if ( RLTable ) {"
    'Stream = Stream & "alert('before: ' + CountElement.value);"
    Stream = Stream & "        insrtcnt.value = parseInt( insrtcnt.value )+1;"
    Stream = Stream & "        NewRowNumber = parseInt( CountElement.value )+1;"
    Stream = Stream & "        CountElement.value = NewRowNumber;"
    'Stream = Stream & "alert('admin: ' + CountElement.value);"
    Stream = Stream & "        tRows = RLTable.getElementsByTagName(""TR"");"
    '
    Stream = Stream & "        NewRow = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow.insertCell(0);"
    Stream = Stream & "        NewCell1.colSpan = '2';"
    Stream = Stream & "        NewCell1.innerHTML='<b>Alternate Contact</b>';"
    '
    Stream = Stream & "        NewRow0 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow0.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Main.GetFormInputCheckBox(RequestNameDeleteRow & "'+NewRowNumber+'") & "';"
    Stream = Stream & "        NewCell2 = NewRow0.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= 'Delete';"
    '
    Stream = Stream & "        NewRow2 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow2.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & "First Name:';"
    Stream = Stream & "        NewCell2 = NewRow2.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameFirstName & "'+NewRowNumber+'", Main.GetStreamText(RequestNameFirstName & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow3 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow3.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & "Last Name:';"
    Stream = Stream & "        NewCell2 = NewRow3.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameLastName & "'+NewRowNumber+'", Main.GetStreamText(RequestNameLastName & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow4 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow4.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & "Title:';"
    Stream = Stream & "        NewCell2 = NewRow4.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameTitle & "'+NewRowNumber+'", Main.GetStreamText(RequestNameTitle & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow5 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow5.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & GetProperty("College/Organization") & "';"
    Stream = Stream & "        NewCell2 = NewRow5.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameCompany & "'+NewRowNumber+'", Main.GetStreamText(RequestNameCompany & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow6 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow6.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & GetProperty("Address") & "';"
    Stream = Stream & "        NewCell2 = NewRow6.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameAddress & "'+NewRowNumber+'", Main.GetStreamText(RequestNameAddress & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow6 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow6.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("Address2") & "';"
    Stream = Stream & "        NewCell2 = NewRow6.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameAddress2 & "'+NewRowNumber+'", Main.GetStreamText(RequestNameAddress2 & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow7 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow7.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & GetProperty("City") & "';"
    Stream = Stream & "        NewCell2 = NewRow7.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameCity & "'+NewRowNumber+'", Main.GetStreamText(RequestNameCity & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow8 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow8.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & GetProperty("State") & "';"
    Stream = Stream & "        NewCell2 = NewRow8.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameState & "'+NewRowNumber+'", Main.GetStreamText(RequestNameState & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow9 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow9.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & GetProperty("Zip") & "';"
    Stream = Stream & "        NewCell2 = NewRow9.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameZip & "'+NewRowNumber+'", Main.GetStreamText(RequestNameZip & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow10 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow10.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & GetProperty("Phone") & "';"
    Stream = Stream & "        NewCell2 = NewRow10.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNamePhone & "'+NewRowNumber+'", Main.GetStreamText(RequestNamePhone & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow11 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow11.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("Fax") & "';"
    Stream = Stream & "        NewCell2 = NewRow11.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameFax & "'+NewRowNumber+'", Main.GetStreamText(RequestNameFax & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow12 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow12.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & "Email:';"
    Stream = Stream & "        NewCell2 = NewRow12.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & Main.GetFormInputText(RequestNameEmail & "'+NewRowNumber+'", Main.GetStreamText(RequestNameEmail & "'+NewRowNumber+'"), , 25) & "';"
    '
    Stream = Stream & "        NewRow13 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow13.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Required & "Username:';"
    Stream = Stream & "        NewCell2 = NewRow13.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & Main.GetFormInputText(RequestNameUsername & "'+NewRowNumber+'", Main.GetStreamText(RequestNameUsername & "'+NewRowNumber+'"), , 25) & "';"
    '
    Stream = Stream & "        NewRow14 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);" & vbCrLf
    Stream = Stream & "        NewCell1 = NewRow14.insertCell(0);" & vbCrLf
    Stream = Stream & "        NewCell1.align = 'right';" & vbCrLf
    Stream = Stream & "        NewCell1.innerHTML='" & Required & "Password:';" & vbCrLf
    Stream = Stream & "        NewCell2 = NewRow14.insertCell(1);" & vbCrLf
    Stream = Stream & "        NewCell2.innerHTML= '" & Main.GetFormInputText(RequestNamePassword & "'+NewRowNumber+'", Main.GetStreamText(RequestNamePassword & "'+NewRowNumber+'"), , 25) & "';" & vbCrLf
    '
    Stream = Stream & "    }" & vbCrLf
    Stream = Stream & "}" & vbCrLf
    '
    Stream = Stream & "function InsertParticipantRows() {"
    Stream = Stream & "    var RLTable = document.getElementById('" & ContactTable & "');"
    Stream = Stream & "    var CountElement = document.getElementById('" & RequestNameRowCount & "');"
    Stream = Stream & "    var tCols,tRows,NewRowNumber;"
    Stream = Stream & "    var insrtcnt = document.getElementById('" & RequestNameParticipantCount & "');;"
    Stream = Stream & "    var NewRow0, NewRow, NewRow2, NewRow3, NewRow4, NewRow5, NewRow6, NewRow7;"
    Stream = Stream & "    var NewRow8, NewRow9, NewRow10, NewRow11, NewRow12, NewRow13, NewRow14, NewRow15;"
    Stream = Stream & "    var NewCell1,NewCell2,NewCell3,NewCell4,NewCell5,NewCell6;"
    Stream = Stream & "    var NewCell7,NewCell8,NewCell9,NewCell10,NewCell11,NewCell12;"
    Stream = Stream & "    var NewCell3,NewCell4,NewCell4,NewCell4,NewCell4,NewCell8;"
    Stream = Stream & "    if ( RLTable ) {"
    'Stream = Stream & "alert('before: ' + CountElement.value);"
    Stream = Stream & "        NewRowNumber = parseInt( CountElement.value )+1;"
    Stream = Stream & "        insrtcnt.value = parseInt( insrtcnt.value )+1;"
    Stream = Stream & "        CountElement.value = NewRowNumber;"
    'Stream = Stream & "alert('admin: ' + CountElement.value);"
    Stream = Stream & "        tRows = RLTable.getElementsByTagName(""TR"");"
    '
    Stream = Stream & "        NewRow = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow.insertCell(0);"
    Stream = Stream & "        NewCell1.colSpan = '2';"
    Stream = Stream & "        NewCell1.innerHTML='<b>Participant</b>';"
    '
    Stream = Stream & "        NewRow0 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow0.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Main.GetFormInputCheckBox(RequestNameDeleteRow & "'+NewRowNumber+'") & "';"
    Stream = Stream & "        NewCell2 = NewRow0.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= 'Delete';"
    '
    Stream = Stream & "        NewRow15 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow15.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & Main.GetFormInputCheckBox(RequestNamePresenter & "'+NewRowNumber+'") & "';"
    Stream = Stream & "        NewCell2 = NewRow15.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= 'I am the moderator/facilitator for this session';"
    '
    Stream = Stream & "        NewRow2 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow2.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='First Name:';"
    Stream = Stream & "        NewCell2 = NewRow2.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepFirstName & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepFirstName & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow3 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow3.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='Last Name:';"
    Stream = Stream & "        NewCell2 = NewRow3.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepLastName & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepLastName & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow4 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow4.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='Title:';"
    Stream = Stream & "        NewCell2 = NewRow4.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepTitle & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepTitle & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow5 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow5.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("College/Organization") & "';"
    Stream = Stream & "        NewCell2 = NewRow5.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepCompany & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepCompany & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow6 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow6.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("Address") & "';"
    Stream = Stream & "        NewCell2 = NewRow6.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepAddress & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepAddress & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow6 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow6.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("Address2") & "';"
    Stream = Stream & "        NewCell2 = NewRow6.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepAddress2 & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepAddress2 & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow7 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow7.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("City") & "';"
    Stream = Stream & "        NewCell2 = NewRow7.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepCity & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepCity & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow8 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow8.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("State") & "';"
    Stream = Stream & "        NewCell2 = NewRow8.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepState & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepState & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow9 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow9.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("Zip") & "';"
    Stream = Stream & "        NewCell2 = NewRow9.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepZip & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepZip & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow10 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow10.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("Phone") & "';"
    Stream = Stream & "        NewCell2 = NewRow10.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepPhone & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepPhone & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow11 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow11.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='" & GetProperty("Fax") & "';"
    Stream = Stream & "        NewCell2 = NewRow11.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & GetField(RequestNameRepFax & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepFax & "'+NewRowNumber+'")) & "';"
    '
    Stream = Stream & "        NewRow12 = RLTable.insertRow(parseInt( RLTable.rows.length )-1);"
    Stream = Stream & "        NewCell1 = NewRow12.insertCell(0);"
    Stream = Stream & "        NewCell1.align = 'right';"
    Stream = Stream & "        NewCell1.innerHTML='Email:';"
    Stream = Stream & "        NewCell2 = NewRow12.insertCell(1);"
    Stream = Stream & "        NewCell2.innerHTML= '" & Main.GetFormInputText(RequestNameRepEmail & "'+NewRowNumber+'", Main.GetStreamText(RequestNameRepEmail & "'+NewRowNumber+'"), , 25) & "';"
    '
    Stream = Stream & "    }" & vbCrLf
    Stream = Stream & "}" & vbCrLf
    
    '
    Stream = Stream & "</script>" & vbCrLf
    '
    GetScripts = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetScripts", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetSubmissionCriteria(Optional UseParticipants As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim CS As Long
    Dim Stream As String
    Dim ContentName As String
    '
    If UseParticipants Then
        ContentName = ContentNameSessionSubmissionsRules
    Else
        ContentName = ContentNameContactRules
    End If
    '
    CS = Main.OpenCSContent(ContentName, "(SessionSubmissionsID=" & GetSubmissionID() & ") AND (MemberID<>" & Main.MemberID & ") AND (MemberID<>0)", , , , , "MemberID")
    Do While Main.CSOK(CS)
        If Stream <> "" Then
            Stream = Stream & " OR "
        End If
        Stream = Stream & "(ID=" & Main.GetCSInteger(CS, "MemberID") & ")"
        Call Main.NextCSRecord(CS)
    Loop
    Call Main.CloseCS(CS)
    '
    If Stream = "" Then
        Stream = "(id=0)"
    End If
    '
    GetSubmissionCriteria = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetSubmissionCriteria", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormLogin() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ButtonString As String
    Dim Hiddenstring As String
    '
    FormSerial = (Rnd() * 10000) * Main.VisitID
    Call Main.SetVisitProperty("RFPAppFormSerial", FormSerial)
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSerial, FormSerial)
    If Not EmailFlag Then
        If Not Main.IsAuthenticated() Then
            ButtonString = ButtonString & Main.GetFormButton(FormButtonLogin)
            ButtonString = ButtonString & ButtonSpacer
        End If
        ButtonString = ButtonString & Main.GetFormButton(FormButtonContinue)
    Else
        '
        Hiddenstring = "<INPUT TYPE=""HIDDEN"" NAME=""ccFormSN"" VALUE=""160990832"">"
        Hiddenstring = Hiddenstring & "<INPUT TYPE=""HIDDEN"" NAME=""160990832Type"" VALUE=""lk0q56am09"">"
        '
        ButtonString = ButtonString & Hiddenstring & "<INPUT TYPE=""SUBMIT"" NAME=""Button"" VALUE=""  Send Password  "">"
    End If
    '
    Stream = Main.GetFormStart()
    Stream = Stream & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"">"
    Stream = Stream & GetPageHead("RFP Login")
    Stream = Stream & GetFormRow("", GetAutenticationFields())
    Stream = Stream & GetFormRow(FormSpacer, ButtonString)
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    '
    GetFormLogin = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFormLogin", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Public Function GetAutenticationFields() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=3>"
    If Not EmailFlag Then
        If Not Main.IsAuthenticated() Then
            Stream = Stream & GetFormRow("Username: ", GetField(RequestNameUsername, ""))
            Stream = Stream & GetFormRow("Password: ", Main.GetFormInputText(RequestNamePassword, "", , 25, , True))
            Stream = Stream & GetFormRow("", "<a href=""" & Main.ServerPage & WorkingQueryString & RequestNameEmailFlag & "=1"">Forgot your username or password?</a>")
        Else
            Stream = Stream & GetFormRow2("<b>You are currently logged in as " & Main.MemberName & ". If this is not you, please <a href=""" & Main.ServerPage & WorkingQueryString & "method=logout"">click here</a>.</b>")
            End If
    Else
        Stream = Stream & GetFormRow("Email: ", "<INPUT NAME=""160990832email"" VALUE="""" SIZE=""18"" MAXLENGTH=""50"">")
        End If
    Stream = Stream & "</TABLE>"
    '
    GetAutenticationFields = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetAutenticationFields", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormLogin()
    On Error GoTo ErrorTrap
    '
    Dim MembershipOption As String
    Dim Username As String
    Dim Password As String
    '
    Username = Main.GetStreamText(RequestNameUsername)
    Password = Main.GetStreamText(RequestNamePassword)
    '
    If (ButtonValue = FormButtonLogin) Then
        If Username = "" Then
            ErrorString = ErrorString & "Please enter your username in order to continue.<br />"
        End If
        If Password = "" Then
            ErrorString = ErrorString & "Please enter your password in order to continue.<br />"
        End If
        If ErrorString = "" Then
            Call Main.LoginMember(Username, Password)
            FormID = FormRequest
        Else
            FormID = FormLogin
        End If
    Else
        If ButtonValue = FormButtonPassword Then
            FormID = FormLogin
        Else
            '
            '   22 Feb 11 - commented out
            '
            'Call Main.LogoutMember
            '
            FormID = FormRequest
        End If
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "ProcessFormLogin", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFileRow(RequestName As String, FilePath As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim SlashPosition As Long
    Dim Filename As String
    '
    If FilePath <> "" Then
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=2>"
        SlashPosition = InStrRev(FilePath, "/", Len(FilePath), vbTextCompare)
        Filename = Right(FilePath, Len(FilePath) - SlashPosition)
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD><nobr><a target=_blank href=""http://" & Main.ServerHost & Main.ServerFilePath & FilePath & """>" & Filename & "</a></nobr></TD>"
        Stream = Stream & "<TD>Delete: " & Main.GetFormInputCheckBox(RequestNameDeleteFile & "_" & RequestName) & "</TD>"
        Stream = Stream & "<TD width=""100%"">Change: " & Main.GetFormInputFile(RequestName) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=""100"" height=""1""></TD>"
        Stream = Stream & "<TD><img src=""/ccLib/images/spacer.gif"" width=""75"" height=""1""></TD>"
        Stream = Stream & "<TD width=""100%""><img src=""/ccLib/images/spacer.gif"" width=""100"" height=""1""></TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "</TABLE>"
    Else
        Stream = Main.GetFormInputFile(RequestName)
        End If
    '
    GetFileRow = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFileRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetParticipantStreamFields(GivenRowCount As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = GetFormRow(Main.GetFormInputCheckBox(RequestNameDeleteRow & GivenRowCount), "Delete")
    Stream = Stream & GetFormRow2(FormSpacer)
    Stream = Stream & GetFormRow(Main.GetFormInputCheckBox(RequestNamePresenter & GivenRowCount, Main.GetStreamBoolean(RequestNamePresenter & GivenRowCount)), "I am the moderator/facilitator for this session")
    Stream = Stream & GetFormRow(Required & GetProperty("First Name"), GetField(RequestNameRepFirstName & GivenRowCount, Main.GetStreamText(RequestNameRepFirstName & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("Last Name"), GetField(RequestNameRepLastName & GivenRowCount, Main.GetStreamText(RequestNameRepLastName & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("Title"), GetField(RequestNameRepTitle & GivenRowCount, Main.GetStreamText(RequestNameRepTitle & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("College/Organization"), GetField(RequestNameRepCompany & GivenRowCount, Main.GetStreamText(RequestNameRepCompany & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("Address"), GetField(RequestNameRepAddress & GivenRowCount, Main.GetStreamText(RequestNameRepAddress & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameRepAddress2 & GivenRowCount, Main.GetStreamText(RequestNameRepAddress2 & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("City"), GetField(RequestNameRepCity & GivenRowCount, Main.GetStreamText(RequestNameRepCity & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("State"), GetField(RequestNameRepState & GivenRowCount, Main.GetStreamText(RequestNameRepState & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("Zip"), GetField(RequestNameRepZip & GivenRowCount, Main.GetStreamText(RequestNameRepZip & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("Phone"), GetField(RequestNameRepPhone & GivenRowCount, Main.GetStreamText(RequestNameRepPhone & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameRepFax & GivenRowCount, Main.GetStreamText(RequestNameRepFax & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & GetProperty("Email"), GetField(RequestNameRepEmail & GivenRowCount, Main.GetStreamText(RequestNameRepEmail & GivenRowCount)))
    '
    GetParticipantStreamFields = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetParticipantStreamFields", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetContactStreamFields(GivenRowCount As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = GetFormRow(Main.GetFormInputCheckBox(RequestNameDeleteRow & GivenRowCount), "Delete")
    Stream = Stream & GetFormRow(GetProperty("First Name"), GetField(RequestNameFirstName & GivenRowCount, Main.GetStreamText(RequestNameFirstName & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Last Name"), GetField(RequestNameLastName & GivenRowCount, Main.GetStreamText(RequestNameLastName & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Title"), GetField(RequestNameTitle & GivenRowCount, Main.GetStreamText(RequestNameTitle & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("College/Organization"), GetField(RequestNameCompany & GivenRowCount, Main.GetStreamText(RequestNameCompany & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Address"), GetField(RequestNameAddress & GivenRowCount, Main.GetStreamText(RequestNameAddress & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameAddress2 & GivenRowCount, Main.GetStreamText(RequestNameAddress2 & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("City"), GetField(RequestNameCity & GivenRowCount, Main.GetStreamText(RequestNameCity & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("State"), GetField(RequestNameState & GivenRowCount, Main.GetStreamText(RequestNameState & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Zip"), GetField(RequestNameZip & GivenRowCount, Main.GetStreamText(RequestNameZip & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Phone"), GetField(RequestNamePhone & GivenRowCount, Main.GetStreamText(RequestNamePhone & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameFax & GivenRowCount, Main.GetStreamText(RequestNameFax & GivenRowCount)))
    Stream = Stream & GetFormRow(GetProperty("Email"), GetField(RequestNameEmail & GivenRowCount, Main.GetStreamText(RequestNameEmail & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & "Username:", GetField(RequestNameUsername & GivenRowCount, Main.GetStreamText(RequestNameUsername & GivenRowCount)))
    Stream = Stream & GetFormRow(Required & "Password:", GetField(RequestNamePassword & GivenRowCount, Main.GetStreamText(RequestNamePassword & GivenRowCount)))
    '
    GetContactStreamFields = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetContactStreamFields", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormContacts()
    On Error GoTo ErrorTrap
    '
    Dim CS As Long
    Dim Counter As Long
    Dim ErrorFlag As Boolean
    Dim UsernameFlag As Boolean
    Dim LocalMemberID As Long
    Dim CSPointer As Long
    '
    If FirstName = "" Then
        ErrorString = ErrorString & "Please enter your first name.<br />"
    End If
    If LastName = "" Then
        ErrorString = ErrorString & "Please enter your last name.<br />"
    End If
    If Title = "" Then
        ErrorString = ErrorString & "Please enter your title.<br />"
    End If
    If Company = "" Then
        ErrorString = ErrorString & "Please enter your college.<br />"
    End If
    If Address = "" Then
        ErrorString = ErrorString & "Please enter your address.<br />"
    End If
    If City = "" Then
        ErrorString = ErrorString & "Please enter your city.<br />"
    End If
    If State = "" Then
        ErrorString = ErrorString & "Please enter your state.<br />"
    End If
    If Zip = "" Then
        ErrorString = ErrorString & "Please enter your zip code.<br />"
    End If
    If Phone = "" Then
        ErrorString = ErrorString & "Please enter your phone number.<br />"
    End If
    If eMail = "" Then
        ErrorString = ErrorString & "Please enter your email address.<br />"
    End If
    If Username = "" Then
        ErrorString = ErrorString & "Please enter a username.<br />"
    Else
        CS = Main.OpenCSContent(ContentNamePeople, "(ID<>" & Main.MemberID & ") and (username=" & Main.EncodeSQLText(Username) & ")", , False)
        If Main.CSOK(CS) Then
            ErrorString = ErrorString & "The username you have entered is already in use, please select an alternate username.<br />"
        End If
        Call Main.CloseCS(CS)
    End If
    '
    If Password = "" Then
        ErrorString = ErrorString & "Please enter your a password.<br />"
        End If
    '
    For Counter = 1 To RowCount
        If Not kmaEncodeBoolean(Main.GetStreamBoolean(RequestNameDeleteRow & Counter)) Then
            '
            If Main.GetStreamText(RequestNameFirstName & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameLastName & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameTitle & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameCompany & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameAddress & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameCity & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameState & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameZip & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNamePhone & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameEmail & Counter) = "" Then
                ErrorFlag = True
            End If
            If Main.GetStreamText(RequestNameUsername & Counter) = "" Then
                ErrorFlag = True
            Else
                CS = Main.OpenCSContent(ContentNamePeople, "(ID<>" & Main.GetStreamInteger(RequestNameContactID & Counter) & ") and (username=" & Main.EncodeSQLText(Main.GetStreamText(RequestNameUsername & Counter)) & ")", , False)
                If Main.CSOK(CS) Then
                    UsernameFlag = True
                End If
                Call Main.CloseCS(CS)
            End If
            If Main.GetStreamText(RequestNamePassword & Counter) = "" Then
                ErrorFlag = True
            End If
            If ErrorFlag Then
                ErrorString = ErrorString & "Please complete information for Alternate Contact " & Counter & ".<br />"
            End If
            If UsernameFlag Then
                ErrorString = ErrorString & "The username you have entered for Alternate Contact " & Counter & " is already in use, please select an alternate username.<br />"
            End If
            ErrorFlag = False
            UsernameFlag = False
            '
        End If
    Next
    '
    Counter = 0
    '
    If ErrorString = "" Then
        CS = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.MemberID)
        If Main.CSOK(CS) Then
            Call Main.SetCS(CS, "Name", FirstName & " " & LastName)
            Call Main.SetCS(CS, "FirstName", FirstName)
            Call Main.SetCS(CS, "LastName", LastName)
            Call Main.SetCS(CS, "Title", Title)
            Call Main.SetCS(CS, "Company", Company)
            Call Main.SetCS(CS, "Address", Address)
            Call Main.SetCS(CS, "Address2", Address2)
            Call Main.SetCS(CS, "City", City)
            Call Main.SetCS(CS, "State", State)
            Call Main.SetCS(CS, "Zip", Zip)
            Call Main.SetCS(CS, "Phone", Phone)
            Call Main.SetCS(CS, "Fax", Fax)
            Call Main.SetCS(CS, "Email", eMail)
            Call Main.SetCS(CS, "Username", Username)
            Call Main.SetCS(CS, "Password", Password)
            Call Main.SaveCSRecord(CS)
            '
            Call Main.DeleteContentRecords(ContentNameContactRules, "(SessionSubmissionsID=" & SubmissionID & ") AND (MemberID=" & Main.MemberID & ")")
            '
            CSPointer = Main.InsertCSContent(ContentNameContactRules)
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "SessionSubmissionsID", SubmissionID)
                Call Main.SetCS(CSPointer, "MemberID", Main.MemberID)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
        End If
        Call Main.CloseCS(CS)
        '
        For Counter = 1 To RowCount
            LocalMemberID = Main.GetStreamInteger(RequestNameContactID & Counter)
            If Not kmaEncodeBoolean(Main.GetStreamBoolean(RequestNameDeleteRow & Counter)) Then
                '
                If LocalMemberID = 0 Then
                    CS = Main.InsertCSContent(ContentNamePeople)
                Else
                    CS = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.GetStreamInteger(RequestNameContactID & Counter))
                End If
                If Main.CSOK(CS) Then
                    LocalMemberID = Main.GetCSInteger(CS, "ID")
                    Call Main.SetCS(CS, "Name", Main.GetStreamText(RequestNameFirstName & Counter) & " " & Main.GetStreamText(RequestNameLastName & Counter))
                    Call Main.SetCS(CS, "FirstName", Main.GetStreamText(RequestNameFirstName & Counter))
                    Call Main.SetCS(CS, "LastName", Main.GetStreamText(RequestNameLastName & Counter))
                    Call Main.SetCS(CS, "Title", Main.GetStreamText(RequestNameTitle & Counter))
                    Call Main.SetCS(CS, "Company", Main.GetStreamText(RequestNameCompany & Counter))
                    Call Main.SetCS(CS, "Address", Main.GetStreamText(RequestNameAddress & Counter))
                    Call Main.SetCS(CS, "Address2", Main.GetStreamText(RequestNameAddress2 & Counter))
                    Call Main.SetCS(CS, "City", Main.GetStreamText(RequestNameCity & Counter))
                    Call Main.SetCS(CS, "State", Main.GetStreamText(RequestNameState & Counter))
                    Call Main.SetCS(CS, "Zip", Main.GetStreamText(RequestNameZip & Counter))
                    Call Main.SetCS(CS, "Phone", Main.GetStreamText(RequestNamePhone & Counter))
                    Call Main.SetCS(CS, "Fax", Main.GetStreamText(RequestNameFax & Counter))
                    Call Main.SetCS(CS, "Email", Main.GetStreamText(RequestNameEmail & Counter))
                    Call Main.SetCS(CS, "Username", Main.GetStreamText(RequestNameUsername & Counter))
                    Call Main.SetCS(CS, "Password", Main.GetStreamText(RequestNamePassword & Counter))
                    Call Main.SaveCSRecord(CS)
                    '
                    Call Main.DeleteContentRecords(ContentNameContactRules, "(SessionSubmissionsID=" & SubmissionID & ") AND (MemberID=" & LocalMemberID & ")")
                    '
                    CSPointer = Main.InsertCSContent(ContentNameContactRules)
                    If Main.CSOK(CSPointer) Then
                        Call Main.SetCS(CSPointer, "SessionSubmissionsID", SubmissionID)
                        Call Main.SetCS(CSPointer, "MemberID", LocalMemberID)
                        Call Main.SaveCSRecord(CSPointer)
                        End If
                    Call Main.CloseCS(CSPointer)
                    '
                End If
                Call Main.CloseCS(CS)
                '
            Else
                Call Main.DeleteContentRecords(ContentNameContactRules, "(SessionSubmissionsID=" & SubmissionID & ") AND (MemberID=" & LocalMemberID & ")")
            End If
        Next
        '
        FormID = FormParticipants
    Else
        FormID = FormContacts
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MembershipClass2", "ProcessFormContacts", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormParticipants()
    On Error GoTo ErrorTrap
    '
    Dim CS As Long
    Dim Counter As Long
    Dim ErrorFlag As Boolean
    Dim LocalMemberID As Long
    Dim CSPointer As Long
    '
'    For Counter = 1 To RowCount
'        If Not kmaEncodeBoolean(Main.GetStreamBoolean(RequestNameDeleteRow & Counter)) Then
'            '
'            If Main.GetStreamText(RequestNameFirstName & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameLastName & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameTitle & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameCompany & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameAddress & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameCity & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameState & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameZip & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNamePhone & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If Main.GetStreamText(RequestNameEmail & Counter) = "" Then
'                ErrorFlag = True
'            End If
'            If ErrorFlag Then
'                ErrorString = ErrorString & "Please complete information for participant " & Counter & ".<br />"
'            End If
'            ErrorFlag = False
'            '
'        End If
'    Next
    '
    Counter = 0
    '
    If ErrorString = "" Then
        '
'        LocalMemberID = Main.GetStreamInteger(RequestNameContactID)
'        If (LocalMemberID = 0) And (RepFirstName <> "") And (RepLastName <> "") Then
'            If LocalMemberID = 0 Then
'                CS = Main.InsertCSContent(ContentNamePeople)
'            Else
'                CS = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.MemberID)
'            End If
'            If Main.CSOK(CS) Then
'                LocalMemberID = Main.GetCSInteger(CS, "ID")
'                Call Main.SetCS(CS, "Name", RepFirstName & " " & RepLastName)
'                Call Main.SetCS(CS, "FirstName", RepFirstName)
'                Call Main.SetCS(CS, "LastName", RepLastName)
'                Call Main.SetCS(CS, "Title", RepTitle)
'                Call Main.SetCS(CS, "Company", RepCompany)
'                Call Main.SetCS(CS, "Address", RepAddress)
'                Call Main.SetCS(CS, "Address2", RepAddress2)
'                Call Main.SetCS(CS, "City", RepCity)
'                Call Main.SetCS(CS, "State", RepState)
'                Call Main.SetCS(CS, "Zip", RepZip)
'                Call Main.SetCS(CS, "Phone", RepPhone)
'                Call Main.SetCS(CS, "Fax", RepFax)
'                Call Main.SetCS(CS, "Email", RepeMail)
'                Call Main.SaveCSRecord(CS)
'                '
'                Call Main.DeleteContentRecords(ContentNameSessionSubmissionsRules, "(SessionSubmissionsID=" & SubmissionID & ") AND (MemberID=" & LocalMemberID & ")")
'                '
'                CSPointer = Main.InsertCSContent(ContentNameSessionSubmissionsRules)
'                If Main.CSOK(CSPointer) Then
'                    Call Main.SetCS(CSPointer, "SessionSubmissionsID", SubmissionID)
'                    Call Main.SetCS(CSPointer, "MemberID", LocalMemberID)
'                    Call Main.SaveCSRecord(CSPointer)
'                    End If
'                Call Main.CloseCS(CSPointer)
'                '
'            End If
'            Call Main.CloseCS(CS)
'        End If
        '
        For Counter = 1 To RowCount
            LocalMemberID = Main.GetStreamInteger(RequestNameContactID & Counter)
            If (Not kmaEncodeBoolean(Main.GetStreamBoolean(RequestNameDeleteRow & Counter))) And ((Main.GetStreamText(RequestNameRepFirstName & Counter) <> "") And (Main.GetStreamText(RequestNameRepLastName & Counter) <> "")) Then
                '
                If LocalMemberID = 0 Then
                    CS = Main.InsertCSContent(ContentNamePeople)
                Else
                    CS = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.GetStreamInteger(RequestNameContactID & Counter))
                End If
                If Main.CSOK(CS) Then
                    LocalMemberID = Main.GetCSInteger(CS, "ID")
                    Call Main.SetCS(CS, "Presenter", Main.GetStreamBoolean(RequestNamePresenter & Counter))
                    Call Main.SetCS(CS, "Name", Main.GetStreamText(RequestNameRepFirstName & Counter) & " " & Main.GetStreamText(RequestNameRepLastName & Counter))
                    Call Main.SetCS(CS, "FirstName", Main.GetStreamText(RequestNameRepFirstName & Counter))
                    Call Main.SetCS(CS, "LastName", Main.GetStreamText(RequestNameRepLastName & Counter))
                    Call Main.SetCS(CS, "Title", Main.GetStreamText(RequestNameRepTitle & Counter))
                    Call Main.SetCS(CS, "Company", Main.GetStreamText(RequestNameRepCompany & Counter))
                    Call Main.SetCS(CS, "Address", Main.GetStreamText(RequestNameRepAddress & Counter))
                    Call Main.SetCS(CS, "Address2", Main.GetStreamText(RequestNameRepAddress2 & Counter))
                    Call Main.SetCS(CS, "City", Main.GetStreamText(RequestNameRepCity & Counter))
                    Call Main.SetCS(CS, "State", Main.GetStreamText(RequestNameRepState & Counter))
                    Call Main.SetCS(CS, "Zip", Main.GetStreamText(RequestNameRepZip & Counter))
                    Call Main.SetCS(CS, "Phone", Main.GetStreamText(RequestNameRepPhone & Counter))
                    Call Main.SetCS(CS, "Fax", Main.GetStreamText(RequestNameRepFax & Counter))
                    Call Main.SetCS(CS, "Email", Main.GetStreamText(RequestNameRepEmail & Counter))
                    Call Main.SaveCSRecord(CS)
                    '
                    Call Main.DeleteContentRecords(ContentNameSessionSubmissionsRules, "(SessionSubmissionsID=" & SubmissionID & ") AND (MemberID=" & LocalMemberID & ")")
                    '
                    CSPointer = Main.InsertCSContent(ContentNameSessionSubmissionsRules)
                    If Main.CSOK(CSPointer) Then
                        Call Main.SetCS(CSPointer, "SessionSubmissionsID", SubmissionID)
                        Call Main.SetCS(CSPointer, "MemberID", LocalMemberID)
                        Call Main.SaveCSRecord(CSPointer)
                        End If
                    Call Main.CloseCS(CSPointer)
                    '
                End If
                Call Main.CloseCS(CS)
                '
            Else
                Call Main.DeleteContentRecords(ContentNameSessionSubmissionsRules, "(SessionSubmissionsID=" & SubmissionID & ") AND (MemberID=" & LocalMemberID & ")")
            End If
        Next
        '
        'Call CompleteApp
        '
        FormID = FormReview
    Else
        FormID = FormParticipants
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("MembershipClass2", "ProcessFormParticipants", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub CompleteApp()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim LocalQueryString As String
    Dim Copy As String
    '
    CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "ID=" & Main.EncodeSQLNumber(SubmissionID))
    If Main.CSOK(CSPointer) Then
        Call Main.SetCS(CSPointer, "Completed", 1)
        Call Main.SaveCSRecord(CSPointer)
        End If
    Call Main.CloseCS(CSPointer)
    '
    LocalQueryString = "cid=" & Main.GetContentID(ContentNameSessionSubmissions)
    LocalQueryString = LocalQueryString & "&id=" & SubmissionID
    LocalQueryString = LocalQueryString & "&af=4"
    '
    Copy = "<br /><br /><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?" & LocalQueryString & """>Click here to view the submission</a>"
    '
    Call Main.AddGroupMember(GroupNameSessionSubmitters)
    '
    Call Main.SendSystemEmail(SystemEmailAutoResponder, , Main.MemberID)
    Call Main.SendSystemEmail(SystemEmailNotification, Copy)
    '
    Exit Sub
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "CompleteApp", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormReview() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ButtonString As String
    Dim CS As Long
    Dim addedString As String
    Dim paymentFlag As Boolean
    Dim totalAmount As Double
    Dim linePrice As Double
    Dim csType As Long
    Dim innerStream As String
    '
    FormSerial = (Rnd() * 10000) * Main.VisitID
    Call Main.SetVisitProperty("RFPAppFormSerial", FormSerial)
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSerial, FormSerial)
    ButtonString = ButtonString & "<div id=""buttonContainer"">" & Main.GetFormButton(FormButtonComplete) & "</div>"
    ButtonString = ButtonString & ButtonSpacer & Main.GetFormButton(FormButtonPrevious)
    '
    Stream = Stream & Main.GetUploadFormStart()
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & GetPageHead(PageNameFormReview)
    Stream = Stream & GetFormRow("<b>Session Title:</b>", SessionTitle)
    Stream = Stream & GetFormRow("<b>Session Description:</b>", Description)
    Stream = Stream & GetFormRow("<b>Session Type:</b>", Main.GetRecordName(ContentNameSessionType, TypeID))
    '
    Stream = Stream & GetFormRow("<b>Track:</b>", Main.GetRecordName(ContentNameSessionTracks, TrackID))
    '
    Stream = Stream & GetFormRow("<b>Anticipated Outcomes:</b>", Outcome)
    Stream = Stream & GetFormRow("<b>Board Support Letter:</b>", "<a target=""_blank"" href=""" & Main.ServerFilePath & BoardLetter & """>Click to View</a>")
    '
    '   AV types
    '
    CS = Main.OpenCSContent("AV Submission Rules", "SessionSubmissionsID=" & SubmissionID, , , , , "SessionSubmissionsID,avTypeID")
    If Main.CSOK(CS) Then
        Do While Main.CSOK(CS)
            csType = Main.OpenCSContent("AV Types", "ID=" & Main.GetCSInteger(CS, "avTypeID"), , , , , "Name,Price")
            If Main.CSOK(csType) Then
                linePrice = Main.GetCSNumber(csType, "Price")
                totalAmount = totalAmount + linePrice
                addedString = addedString & Main.GetCS(csType, "Name") & " - " & FormatCurrency(linePrice, 2) & "<br />"
            End If
            Call Main.CloseCS(csType)
            Main.NextCSRecord (CS)
        Loop
        Stream = Stream & GetFormRow("<b>Audiovisual Aids:</b>", addedString, , True)
        paymentFlag = True
    End If
    Call Main.CloseCS(CS)
    addedString = ""
    '
    '   Contatcts
    '
    CS = Main.OpenCSContent(ContentNameContactRules, "SessionSubmissionsID=" & SubmissionID, , , , , "SessionSubmissionsID,MemberID")
    If Main.CSOK(CS) Then
        Do While Main.CSOK(CS)
            addedString = addedString & Main.GetRecordName(ContentNamePeople, Main.GetCSInteger(CS, "MemberID")) & "<br />"
            Main.NextCSRecord (CS)
        Loop
        Stream = Stream & GetFormRow("<b>Contacts:</b>", addedString, , True)
    End If
    Call Main.CloseCS(CS)
    addedString = ""
    '
    '   Participants
    '
    CS = Main.OpenCSContent(ContentNameSessionSubmissionsRules, "SessionSubmissionsID=" & SubmissionID, , , , , "SessionSubmissionsID,MemberID")
    If Main.CSOK(CS) Then
        Do While Main.CSOK(CS)
            addedString = addedString & Main.GetRecordName(ContentNamePeople, Main.GetCSInteger(CS, "MemberID")) & "<br />"
            Main.NextCSRecord (CS)
        Loop
        Stream = Stream & GetFormRow("<b>Participants:</b>", addedString, , True)
    End If
    Call Main.CloseCS(CS)
    '
    '   payment
    '
    If paymentFlag Then
        Stream = Stream & GetFormRow("<b>Total Due:</b>", FormatCurrency(totalAmount, 2), , True)
        innerStream = "<table padding=""5"" cellspacing=""0"">"
        innerStream = innerStream & "<tr>"
        innerStream = innerStream & "<td><input type=""radio"" name=""payment"" id=""paymentPP""></td>"
        innerStream = innerStream & "<td>Pay with PayPal</td>"
        innerStream = innerStream & "</tr>"
        innerStream = innerStream & "<tr>"
        innerStream = innerStream & "<td><input type=""radio"" name=""payment"" id=""paymentBill""></td>"
        innerStream = innerStream & "<td>Bill Me</td>"
        innerStream = innerStream & "</tr>"
        innerStream = innerStream & "</table>"
        Stream = Stream & GetFormRow("<b>Payment Type:</b>", innerStream, , True)
        'Stream = Stream & GetFormRow(FormSpacer, getPayPalButton(CStr(submissionID) & "," & CStr(totalAmount), CStr(totalAmount), "AV Needs"))
    End If
    '
    Stream = Stream & GetFormRow(FormSpacer, Main.GetFormInputHidden("payFlag", paymentFlag) & ButtonString)
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy("ACCT Proposal Footer"))
    '
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd()
    '
    GetFormReview = Stream & getPayPalForm(CStr(SubmissionID) & "," & CStr(totalAmount), CStr(totalAmount), "AV Needs")
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "GetFormRequest", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadFormReview()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    SubmissionID = GetSubmissionID
    '
    CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "(ID=" & Main.EncodeSQLNumber(SubmissionID) & ") AND (MemberID=" & Main.MemberID & ")")
    If Main.CSOK(CSPointer) Then
        SessionTitle = Main.GetCSText(CSPointer, "SessionTitle")
        Description = Main.GetCSText(CSPointer, "Description")
        TypeID = Main.GetCSInteger(CSPointer, "TypeID")
        Outcome = Main.GetCSText(CSPointer, "Outcome")
        AidDescription = Main.GetCSText(CSPointer, "AidDescription")
        FirstName = Main.GetCSText(CSPointer, "FirstName")
        LastName = Main.GetCSText(CSPointer, "LastName")
        Title = Main.GetCSText(CSPointer, "Title")
        Company = Main.GetCSText(CSPointer, "Company")
        Address = Main.GetCSText(CSPointer, "Address")
        Address2 = Main.GetCSText(CSPointer, "Address2")
        City = Main.GetCSText(CSPointer, "City")
        State = Main.GetCSText(CSPointer, "State")
        Zip = Main.GetCSText(CSPointer, "Zip")
        Phone = Main.GetCSText(CSPointer, "Phone")
        Fax = Main.GetCSText(CSPointer, "Fax")
        eMail = Main.GetCSText(CSPointer, "Email")
        TrackID = Main.GetCSInteger(CSPointer, "TrackID")
        PowerPoint = Main.GetCSBoolean(CSPointer, "PowerPoint")
        Video = Main.GetCSBoolean(CSPointer, "Video")
        Connection = Main.GetCSBoolean(CSPointer, "Connection")
        BoardLetter = Main.GetCSText(CSPointer, "BoardLetter")
    End If
    Call Main.CloseCS(CSPointer)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "LoadFormReview", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function getPayPalButton(echoString As String, amount As String, payPalItemName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim o As String
    '
    If Main.ServerPageSecure Then
        o = o & "<a href=""javascript:document.getElementById('PayPal').submit();""><img border=0 src=""https://www.paypal.com/en_US/i/btn/x-click-but01.gif"" alt=""Make payments with PayPal - it's fast, free and secure!"">"
    Else
        o = o & "<a href=""javascript:document.getElementById('PayPal').submit();""><img border=0 src=""http://www.paypal.com/en_US/i/btn/x-click-but01.gif"" alt=""Make payments with PayPal - it's fast, free and secure!"">"
    End If
    '
'    O = O & kmaEndTableRow
'    O = O & kmaStartTableRow
'    O = O & kmaEndTableCell & "<td width=""100%"" align=""center"" valign=""middle"">"
'    O = O & "To use PayPal, simply click on this icon. You will be forwarded to the PayPal system to complete the transaction."
'    O = O & kmaEndTableCell & kmaEndTableRow & kmaEndTable
    '
    getPayPalButton = o
    '
    afterFormStuff = getPayPalForm(echoString, amount, payPalItemName)
    '
    Exit Function
    '
ErrorTrap:
    Call HandleError("acct.RFPSubmissionClass", "getPayPalButton", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
Private Function getPayPalForm(echoString As String, amount As String, payPalItemName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Pos As Long
    Dim o As String
    Dim thankYouPageLink As String
    Dim paymentConfirmLink As String
    Dim paymentCancelLink As String
    Dim qs As String
    Dim ServerLink As String
    Dim rqs As String
    Dim currency_code As String
    '
    rqs = Main.RefreshQueryString
    ServerLink = Main.ServerLink
    Pos = InStr(1, ServerLink, "?")
    If Pos > 0 Then
        ServerLink = Left(ServerLink, Pos - 1)
    End If
    '
    qs = ModifyQueryString(rqs, RequestNameFormID, FormThankYou, True)
    '
    thankYouPageLink = ServerLink & "?" & qs
    paymentConfirmLink = Main.ServerProtocol & Main.ServerHost & "/RFPPayPalConfirm?echo=" & echoString
    paymentCancelLink = Main.ServerProtocol & Main.ServerHost
    '
    currency_code = "USD"
    '
    o = o & "<form id=""PayPal"" name=""PayPal"" action=""" & Main.GetSiteProperty("RFP PayPalPostURL", "https://www.sandbox.paypal.com/cgi-bin/webscr") & """ method=""post"" style=""display:inline;"">"
    o = o & "<input type=""hidden"" name=""cmd"" value=""_xclick"">"
    o = o & "<input type=""hidden"" name=""business"" value=""" & Main.GetSiteProperty("RFP PayPalBusinessName", "josh_1297110464_biz@contensive.com") & """>"
    o = o & "<input type=""hidden"" name=""item_name"" value=""" & payPalItemName & """>"
    o = o & "<input type=""hidden"" name=""currency_code"" value=""" & Main.GetSiteProperty("MeetingSmart PayPalCurrencyCode", "USD") & """>"
    o = o & "<input type=""hidden"" name=""amount"" value=""" & amount & """>"
    o = o & "<input type=""hidden"" name=""return"" value=""" & thankYouPageLink & """>"
    o = o & "<input type=""hidden"" name=""notify_url"" value=""" & paymentConfirmLink & """>"
    o = o & "<input type=""hidden"" name=""shopping_url"" value=""" & paymentCancelLink & """>"
    o = o & "</form>"
    '
    getPayPalForm = o
    '
    Exit Function
    '
ErrorTrap:
    Call HandleError("acct.RFPSubmissionClass", "getPayPalForm", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
'
Private Sub ProcessFormReview()
    On Error GoTo ErrorTrap
    '
    If ButtonValue <> "" Then
        If (Main.GetStreamText("payment") = "") And (Main.GetStreamBoolean("payFlag")) Then
            ErrorString = ErrorString & "Please select a payment type in order to continue.<br />"
        End If
        '
        If ErrorString = "" Then
            Call CompleteApp
            FormID = FormThankYou
        Else
            FormID = FormReview
        End If
        '
    End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "ProcessFormReview", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
