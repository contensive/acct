VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RFPReviewClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private logfilename As String
'
Const RequestNameFormID = "formid"
Const RequestNameSourceFormID = "sourceformid"
Const RequestNameButtonValue = "button"
'
Const ContentNamePeople = "People"
Const ContentNameSessionType = "Session Types"
Const ContentNameSessionSubmissions = "Session Submissions"
Const ContentNameSessionSubmissionsRules = "Session Submissions Rules"
Const ContentNameSessionTracks = "Session Tracks"
Const ContentNameContactRules = "Session Submissions Contact Rules"
'
Const FormList = 100
Const FormEdit = 200
Const FormPeopleEdit = 300
Const FormExport = 400
'
Const FormButtonUpdate = "  Update  "
Const FormButtonCancel = "  Cancel  "
Const FormButtonExportApproved = "  Merge Accepted Submissions to Word "
Const FormButtonExportAll = "  Merge All Submissions to Word "
Const FormButtonDelete = "  Delete "
'
Const FormSpacer = "<img src=""/ccLib/images/spacer.gif"" width=300 height=1>"
Const FormSpacer2 = "<img src=""/ccLib/images/spacer.gif"" width=150 height=1>"
Const ButtonSpacer = "&nbsp;&nbsp;&nbsp;"
'
Const PostParticipantCountString = "1-25,26-50,51-100,101-150,More Than 150"
Const SessionRatingString = "1,2,3,4,5"
'
Const TableIDParticipants = "participants"
'
Const GroupNameSessionSubmitters = "Session Submissions"
'
Const ClassPanelEven = "PanelRowEven"
Const ClassPanelOdd = "PanelRowOdd"
Const ClassPanelHeader = "PanelRowHeader"
'
Const SystemEmailAutoResponder = "Session Submission Auto Responder"
Const SystemEmailNotification = "Session Submission Notification"
'
Const PageNameInstructions = "RFP Review Instructions"
Const PageNamePrimaryPresenter = "RFP Review Primary Presenter"
Const PageNameAdditionalPresenter = "RFP Review Aditional Presenters"
Const PageNameFormParticipantEdit = "RFP Review Aditional Presenter Instructions"
Const PageNameProposalTracking = "RFP Review Proposal Tracking"
Const PageNameProgramInformation = "RFP Review Program Information"
Const PageNamePostCongressFeedback = "RFP Review Post Congress Feedback"
'
Const RequestNameSessionTitle = "sessiontitle"
Const RequestNameDescription = "description"
Const RequestNameTypeID = "typeid"
'
Const RequestNameTrackID = "TrackID"
'
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
'
Const RequestNameAccepted = "accepted"
Const RequestNameAcceptedFor = "acceptedfor"
Const RequestNameJointProposal = "jointproposal"
Const RequestNameJointProposalContact = "jointproposalcontact"
Const RequestNameDateNotified = "datenotified"
Const RequestNameWithdrawn = "withdrawn"
Const RequestNameDateWithdrawn = "datewithdrawn"
Const RequestNameReviewedBy = "reviewedby"
Const RequestNameSupportLetter = "supportletter"
Const RequestNameSupportLetterComments = "supportlettercomments"
Const RequestNameHandout = "handout"
Const RequestNameHandoutComments = "handoutcomments"
Const RequestNamePostHandout = "posthandout"
Const RequestNamePostHandoutComments = "posthandoutcomments"
Const RequestNameAlternateContact = "alternatecontact"
Const RequestNameSchoolCorrespondence = "schoolcorrespondence"
Const RequestNameProgramDate = "programdate"
Const RequestNameProgramTime = "programtime"
Const RequestNameLocation = "location"
Const RequestNameComments = "comments"
Const RequestNamePostParticipantCount = "postparticipantcount"
Const RequestNameSessionRating = "sessionrating"
Const RequestNamePresenterID = "presenterid"
Const RequestNameAudiovisualNotes = "audiovisualnotes"
Const RequestNameEditedDescription = "editeddescription"
'
Const RequestNameParticipantCount = "participantcount"
Const RequestNameSubmissionID = "submissionid"
'
Const ValueYes = True
Const ValueNo = False
'
Private FormID As Long
Private SourceFormID As Long
Private ButtonValue As String
Private WorkingQueryString As String
'
Private ErrorString As String
'
Private SessionTitle As String
Private Description As String
Private TypeID As Long
Private TrackID As Long
Private Outcome As String
Private Name As String
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
Private AidDescription As String
Private SubmissionID As Long
Private ParticipantCount As Long
'
Private Accepted As Boolean
Private AcceptedFor As String
Private JointProposal As Boolean
Private JointProposalContact As String
Private DateNotified As String
Private Withdrawn As Boolean
Private DateWithdrawn As String
Private ReviewedBy As String
Private SupportLetter As Boolean
Private SupportLetterComments As String
Private Handout As Boolean
Private HandoutComments As String
Private PostHandout As Boolean
Private PostHandoutComments As String
Private AlternateContact As String
Private SchoolCorrespondence As String
Private ProgramDate As String
Private ProgramTime As String
Private Location As String
Private Comments As String
Private PostParticipantCount As Long
Private SessionRating As Long
Private PresenterID As Long
Private EditedDescription As String
Private AudiovisualNotes As String
'
Private PresenterArray(4) As String
'
Private Main As Object
'
Public Sub Init(MainObject As Variant)
    '
    Call AppendLog(logfilename, "enter init")
    '
    Set Main = MainObject
    logfilename = "export" & GetRandomInteger & ".log"    '
    Call LoadForm
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
    Call AppendLog(logfilename, "enter getContent")
    '
    If Not (Main Is Nothing) Then
        GetContent = GetForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadForm()
    On Error GoTo ErrorTrap
    '
    Call AppendLog(logfilename, "enter loadForm")
    '
    FormID = Main.GetStreamInteger(RequestNameFormID)
    SourceFormID = Main.GetStreamInteger(RequestNameSourceFormID)
    ButtonValue = Main.GetStreamText(RequestNameButtonValue)
    SubmissionID = Main.GetStreamInteger(RequestNameSubmissionID)
    PresenterID = Main.GetStreamInteger(RequestNamePresenterID)
    WorkingQueryString = Main.RefreshQueryString
    If WorkingQueryString <> "" Then
        WorkingQueryString = "?" & WorkingQueryString & "&"
    Else
        WorkingQueryString = "?"
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "LoadForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessForm()
    On Error GoTo ErrorTrap
    '
    Call AppendLog(logfilename, "enter processForm")
    '
    Select Case SourceFormID
        Case FormList
            Select Case ButtonValue
                Case FormButtonExportApproved
                    Call ExportSubmissions(True)
                Case FormButtonExportAll
                    Call ExportSubmissions(False)
            End Select
        Case FormEdit
            Select Case ButtonValue
                Case FormButtonCancel
                    FormID = FormList
                Case FormButtonUpdate
                    Call LoadFormEdit
                    Call ProcessFormEdit
            End Select
        Case FormPeopleEdit
            Select Case ButtonValue
                Case FormButtonDelete
                    Call Main.DeleteContentRecords(ContentNameSessionSubmissionsRules, "MemberID=" & PresenterID)
                    FormID = FormEdit
                Case FormButtonCancel
                    FormID = FormEdit
                Case FormButtonUpdate
                    FormID = FormEdit
                    Call LoadFormPeopleEdit
                    Call ProcessFormPeopleEdit
            End Select
    End Select
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "ProcessForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetForm() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Call AppendLog(logfilename, "enter GetForm")
    '
    Select Case FormID
        Case FormEdit
            Call LoadFormEdit
            Stream = Stream & GetFormEdit
        Case FormPeopleEdit
            Call LoadFormPeopleEdit
            Stream = Stream & GetFormPeopleEdit
        Case Else
            FormID = FormList
            Stream = Stream & GetFormList
    End Select
    '
    GetForm = Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormList() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim Class As String
    Dim RowNumber As Long
    Dim Link As String
    Dim ButtonString As String
    '
    Call AppendLog(logfilename, "enter GetFormList")
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonExportApproved)
    ButtonString = ButtonString & " &nbsp;" & Main.GetFormButton(FormButtonExportAll)
    '
    CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "Completed=1", "DateAdded")
    If Not Main.IsCSOK(CSPointer) Then
        Stream = Stream _
            & vbCrLf & vbTab & "<div>" _
            & vbCrLf & vbTab & vbTab & "<p>There are no Session Submissions records marked complete.</p>" _
            & vbCrLf & vbTab & "</div>" _
            & ""
    Else
        Stream = Stream & Main.GetFormStart()
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
        Stream = Stream & GetPageHead("RFP Overeview Instructions")
        Stream = Stream & "<TR CLASS=""" & ClassPanelHeader & """>"
        Stream = Stream & "<TD>Session Name</TD>"
        Stream = Stream & "<TD>Session Type</TD>"
        Stream = Stream & "<TD>Date Submitted</TD>"
        Stream = Stream & "<TD>Submitted By</TD>"
        Stream = Stream & "<TD>Accepted</TD>"
        Stream = Stream & "</TR>"
        Do While Main.CSOK(CSPointer)
            '
            If (RowNumber Mod 2) = 0 Then
                Class = ClassPanelOdd
            Else
                Class = ClassPanelEven
                End If
            '
            Link = Main.ServerPage & WorkingQueryString & RequestNameSubmissionID & "=" & Main.GetCSInteger(CSPointer, "ID") & "&" & RequestNameFormID & "=" & FormEdit
            '
            Stream = Stream & "<TR CLASS=""" & Class & """>"
            Stream = Stream & "<TD><a href=""" & Link & """>" & Main.GetCSText(CSPointer, "Name") & "</a></TD>"
            Stream = Stream & "<TD>" & Main.GetCSLookup(CSPointer, "TypeID") & "</TD>"
            Stream = Stream & "<TD>" & Main.GetCSDate(CSPointer, "DateAdded") & "</TD>"
            Stream = Stream & "<TD>" & Main.GetCSText(CSPointer, "FirstName") & " " & Main.GetCSText(CSPointer, "LastName") & "</TD>"
            Stream = Stream & "<TD>" & Main.GetYesNo(Main.GetCSBoolean(CSPointer, "Accepted")) & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
            RowNumber = RowNumber + 1
        Loop
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD>" & FormSpacer2 & "</TD>"
        Stream = Stream & "<TD>" & FormSpacer2 & "</TD>"
        Stream = Stream & "<TD>" & FormSpacer2 & "</TD>"
        Stream = Stream & "<TD>" & FormSpacer2 & "</TD>"
        Stream = Stream & "<TD>" & FormSpacer2 & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD colspan=5>" & ButtonString & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "</TABLE>"
        Stream = Stream & Main.GetFormEnd()
    End If
    Call Main.CloseCS(CSPointer)
    '
    GetFormList = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetFormList", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetFormEdit() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ButtonString As String
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSubmissionID, SubmissionID)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonUpdate)
    ButtonString = ButtonString & " &nbsp;" & Main.GetFormButton(FormButtonCancel)
    '
    Stream = Stream & Main.GetFormStart
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & GetPageHead(PageNameInstructions)
    Stream = Stream & GetFormRow(GetProperty("Suggested Session Title"), GetField(RequestNameSessionTitle, SessionTitle, , 50))
    Stream = Stream & GetFormRow(GetProperty("Session Description"), Description)
    Stream = Stream & GetFormRow(GetProperty("Edited Description"), GetField(RequestNameEditedDescription, EditedDescription, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Session Type"), Main.GetFormInputSelect(RequestNameTypeID, TypeID, ContentNameSessionType))
    '
    '
    Stream = Stream & GetFormRow(GetProperty("Track Type"), Main.GetFormInputSelect(RequestNameTrackID, TrackID, ContentNameSessionTracks))
    '
    '
    Stream = Stream & GetFormRow(GetProperty("Anticipated Outcomes"), GetField(RequestNameOutcome, Outcome, 10, 40))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy(PageNamePrimaryPresenter))
    '
    Stream = Stream & GetFormRow(GetProperty("First Name"), GetField(RequestNameFirstName, FirstName, , 50))
    Stream = Stream & GetFormRow(GetProperty("Last Name"), GetField(RequestNameLastName, LastName, , 50))
    Stream = Stream & GetFormRow(GetProperty("Title"), GetField(RequestNameTitle, Title, , 50))
    Stream = Stream & GetFormRow(GetProperty("College/Organization"), GetField(RequestNameCompany, Company, , 50))
    Stream = Stream & GetFormRow(GetProperty("Address"), GetField(RequestNameAddress, Address, , 50))
    Stream = Stream & GetFormRow(GetProperty("Address2"), GetField(RequestNameAddress2, Address2, , 50))
    Stream = Stream & GetFormRow(GetProperty("City"), GetField(RequestNameCity, City, , 50))
    Stream = Stream & GetFormRow(GetProperty("State"), GetField(RequestNameState, State, , 50))
    Stream = Stream & GetFormRow(GetProperty("Zip"), GetField(RequestNameZip, Zip, , 50))
    Stream = Stream & GetFormRow(GetProperty("Phone"), GetField(RequestNamePhone, Phone, , 50))
    Stream = Stream & GetFormRow(GetProperty("Fax"), GetField(RequestNameFax, Fax, , 50))
    Stream = Stream & GetFormRow(GetProperty("Email"), GetField(RequestNameEmail, eMail, , 50))
    Stream = Stream & GetFormRow(GetProperty("Audiovisual Aids"), GetField(RequestNameAidDescription, AidDescription, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Audiovisual Notes"), GetField(RequestNameAudiovisualNotes, AudiovisualNotes, 10, 40))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy(PageNameAdditionalPresenter))
    Stream = Stream & GetFormRow3(FormSpacer, GetAdditionalPresenters(SubmissionID))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy(PageNameProposalTracking))
    Stream = Stream & GetFormRow(GetProperty("Accepted"), GetYesNo(RequestNameAccepted, Accepted))
    Stream = Stream & GetFormRow(GetProperty("Accepted For"), GetField(RequestNameAcceptedFor, AcceptedFor, , 50))
    Stream = Stream & GetFormRow(GetProperty("Joint Proposal"), GetYesNo(RequestNameJointProposal, JointProposal))
    Stream = Stream & GetFormRow(GetProperty("If yes, then who"), GetField(RequestNameJointProposalContact, JointProposalContact, , 50))
    Stream = Stream & GetFormRow(GetProperty("Date Notification Letter Sent"), GetField(RequestNameDateNotified, DateNotified, , 50))
    Stream = Stream & GetFormRow(GetProperty("Withdrew"), GetYesNo(RequestNameWithdrawn, Withdrawn))
    Stream = Stream & GetFormRow(GetProperty("If yes, withdrawn date"), GetField(RequestNameDateWithdrawn, DateWithdrawn, , 50))
    '
    Stream = Stream & GetFormRow(GetProperty("Reviewed by (Educational Program Review Cmte Member(s))"), GetField(RequestNameReviewedBy, ReviewedBy, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Board Support Letter Received"), GetYesNo(RequestNameSupportLetter, SupportLetter))
    Stream = Stream & GetFormRow(GetProperty("Support Letter Comments"), GetField(RequestNameSupportLetterComments, SupportLetterComments, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Handouts Received"), GetYesNo(RequestNameHandout, Handout))
    Stream = Stream & GetFormRow(GetProperty("Handout Comments"), GetField(RequestNameHandoutComments, HandoutComments, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("OK to post handouts to web"), GetYesNo(RequestNamePostHandout, PostHandout))
    Stream = Stream & GetFormRow(GetProperty("Post Handout Comments"), GetField(RequestNamePostHandoutComments, PostHandoutComments, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Alternate/Additional Primary Contact"), GetField(RequestNameAlternateContact, AlternateContact, 10, 40))
    Stream = Stream & GetFormRow(GetProperty("Correspondence with school"), GetField(RequestNameSchoolCorrespondence, SchoolCorrespondence, 10, 40))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy(PageNameProgramInformation))
    Stream = Stream & GetFormRow(GetProperty("Date in Program"), GetField(RequestNameProgramDate, ProgramDate, , 50))
    Stream = Stream & GetFormRow(GetProperty("Time in Program"), GetField(RequestNameProgramTime, ProgramTime, , 50))
    Stream = Stream & GetFormRow(GetProperty("Session Room/Roundtable Location"), GetField(RequestNameLocation, Location, , 50))
    '
    Stream = Stream & GetFormRow2(Main.GetContentCopy(PageNamePostCongressFeedback))
    Stream = Stream & GetFormRow(GetProperty("Number of Participants in Session"), Main.GetFormInputSelectList(RequestNamePostParticipantCount, PostParticipantCount, PostParticipantCountString))
    Stream = Stream & GetFormRow(GetProperty("Session Rating"), Main.GetFormInputSelectList(RequestNameSessionRating, SessionRating, SessionRatingString))
    Stream = Stream & GetFormRow(GetProperty("Additional Comments"), GetField(RequestNameComments, Comments, 10, 40))
    '
    Stream = Stream & GetFormRow3(FormSpacer, ButtonString)
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd
    '
    GetFormEdit = Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetFormEdit", Err.Number, Err.Source, Err.Description, True, False)
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
    Stream = Stream & "<TD " & AlignmentString & ">" & FieldCaption & ":</TD>"
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
    Call HandleError("RFPReviewClass", "GetProperty", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetField(RequestName As String, Value As String, Optional Height As String, Optional Width As String, Optional MaxLenghth As String) As String
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
    Stream = Main.GetFormInputText(RequestName, Value, Height, Width)
    Stream = Replace(Stream, "<INPUT ", "<INPUT maxlength=" & MaxLenghth & " ", 1, 99, 1)
    '
    GetField = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetField", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadFormEdit()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim TargetMemberID As Long
    '
    If (SourceFormID = FormEdit) Then
        SessionTitle = Main.GetStreamText(RequestNameSessionTitle)
        EditedDescription = Main.GetStreamText(RequestNameEditedDescription)
        TypeID = Main.GetStreamInteger(RequestNameTypeID)
        '
        TrackID = Main.GetStreamInteger(RequestNameTrackID)
        '
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
        AudiovisualNotes = Main.GetStreamText(RequestNameAudiovisualNotes)
        '
        Accepted = Main.GetStreamBoolean(RequestNameAccepted)
        AcceptedFor = Main.GetStreamText(RequestNameAcceptedFor)
        JointProposal = Main.GetStreamBoolean(RequestNameJointProposal)
        JointProposalContact = Main.GetStreamText(RequestNameJointProposalContact)
        DateNotified = Main.GetStreamText(RequestNameDateNotified)
        Withdrawn = Main.GetStreamBoolean(RequestNameWithdrawn)
        DateWithdrawn = Main.GetStreamText(RequestNameDateWithdrawn)
        ReviewedBy = Main.GetStreamText(RequestNameReviewedBy)
        SupportLetter = Main.GetStreamBoolean(RequestNameSupportLetter)
        SupportLetterComments = Main.GetStreamText(RequestNameSupportLetterComments)
        Handout = Main.GetStreamBoolean(RequestNameHandout)
        HandoutComments = Main.GetStreamText(RequestNameHandoutComments)
        PostHandout = Main.GetStreamBoolean(RequestNamePostHandout)
        PostHandoutComments = Main.GetStreamText(RequestNamePostHandoutComments)
        AlternateContact = Main.GetStreamText(RequestNameAlternateContact)
        SchoolCorrespondence = Main.GetStreamText(RequestNameSchoolCorrespondence)
        ProgramDate = Main.GetStreamText(RequestNameProgramDate)
        ProgramTime = Main.GetStreamText(RequestNameProgramTime)
        Location = Main.GetStreamText(RequestNameLocation)
        Comments = Main.GetStreamText(RequestNameComments)
        PostParticipantCount = Main.GetStreamInteger(RequestNamePostParticipantCount)
        SessionRating = Main.GetStreamInteger(RequestNameSessionRating)
        '
    Else
        '
        CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "ID=" & Main.EncodeSQLNumber(SubmissionID))
        If Main.CSOK(CSPointer) Then
            TargetMemberID = Main.GetCSInteger(CSPointer, "MemberID")
            SessionTitle = Main.GetCSText(CSPointer, "SessionTitle")
            Description = Main.GetCSText(CSPointer, "Description")
            EditedDescription = Main.GetCSText(CSPointer, "EditedDescription")
            TypeID = Main.GetCSInteger(CSPointer, "TypeID")
            '
            TrackID = Main.GetCSInteger(CSPointer, "TrackID")
            '
            Outcome = Main.GetCSText(CSPointer, "Outcome")
            AidDescription = Main.GetCSText(CSPointer, "AidDescription")
            AudiovisualNotes = Main.GetCSText(CSPointer, "AudiovisualNotes")
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
            '
            Accepted = Main.GetCSBoolean(CSPointer, "Accepted")
            AcceptedFor = Main.GetCS(CSPointer, "AcceptedFor")
            JointProposal = Main.GetCSBoolean(CSPointer, "JointProposal")
            JointProposalContact = Main.GetCS(CSPointer, "JointProposalContact")
            DateNotified = Main.GetCS(CSPointer, "DateNotified")
            Withdrawn = Main.GetCSBoolean(CSPointer, "Withdrawn")
            DateWithdrawn = Main.GetCS(CSPointer, "DateWithdrawn")
            ReviewedBy = Main.GetCS(CSPointer, "ReviewedBy")
            SupportLetter = Main.GetCSBoolean(CSPointer, "SupportLetter")
            SupportLetterComments = Main.GetCS(CSPointer, "SupportLetterComments")
            Handout = Main.GetCSBoolean(CSPointer, "Handout")
            HandoutComments = Main.GetCS(CSPointer, "HandoutComments")
            PostHandout = Main.GetCSBoolean(CSPointer, "PostHandout")
            PostHandoutComments = Main.GetCS(CSPointer, "PostHandoutComments")
            AlternateContact = Main.GetCS(CSPointer, "AlternateContact")
            SchoolCorrespondence = Main.GetCS(CSPointer, "SchoolCorrespondence")
            ProgramDate = Main.GetCS(CSPointer, "ProgramDate")
            ProgramTime = Main.GetCS(CSPointer, "ProgramTime")
            Location = Main.GetCS(CSPointer, "Location")
            Comments = Main.GetCS(CSPointer, "Comments")
            PostParticipantCount = Main.GetCSInteger(CSPointer, "ParticipantCount")
            SessionRating = Main.GetCSInteger(CSPointer, "SessionRating")
            '
            End If
        Call Main.CloseCS(CSPointer)
        '
        CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(TargetMemberID))
        If Main.CSOK(CSPointer) Then
            If FirstName = "" Then
                FirstName = Main.GetCSText(CSPointer, "FirstName")
                End If
            If LastName = "" Then
                LastName = Main.GetCSText(CSPointer, "LastName")
                End If
            If LastName = "" Then
                Title = Main.GetCSText(CSPointer, "Title")
                End If
            If Company = "" Then
                Company = Main.GetCSText(CSPointer, "Company")
                End If
            If Address = "" Then
                Address = Main.GetCSText(CSPointer, "Address")
                End If
            If Address2 = "" Then
                Address2 = Main.GetCSText(CSPointer, "Address2")
                End If
            If City = "" Then
                City = Main.GetCSText(CSPointer, "City")
                End If
            If State = "" Then
                State = Main.GetCSText(CSPointer, "State")
                End If
            If Zip = "" Then
                Zip = Main.GetCSText(CSPointer, "Zip")
                End If
            If Phone = "" Then
                Phone = Main.GetCSText(CSPointer, "Phone")
                End If
            If Fax = "" Then
                Fax = Main.GetCSText(CSPointer, "Fax")
                End If
            If eMail = "" Then
                eMail = Main.GetCSText(CSPointer, "Email")
                End If
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
    Call HandleError("RFPReviewClass", "LoadFormEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetYesNo(RequestName As String, Value As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=2>"
    Stream = Stream & "<TR>"
    Stream = Stream & "<TD>" & Main.GetFormInputRadioBox(RequestName, ValueYes, Value) & "</TD>"
    Stream = Stream & "<TD>Yes</TD>"
    Stream = Stream & "<TD>" & Main.GetFormInputRadioBox(RequestName, ValueNo, Value) & "</TD>"
    Stream = Stream & "<TD width=""100%"">No</TD>"
    Stream = Stream & "</TR>"
    Stream = Stream & "</TABLE>"
    '
    GetYesNo = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetYesNo", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub ProcessFormEdit()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, "ID=" & Main.EncodeSQLNumber(SubmissionID))
    If Main.CSOK(CSPointer) Then
        Call Main.SetCS(CSPointer, "Name", SessionTitle)
        Call Main.SetCS(CSPointer, "SessionTitle", SessionTitle)
        Call Main.SetCS(CSPointer, "EditedDescription", EditedDescription)
        Call Main.SetCS(CSPointer, "TypeID", TypeID)
        '
        Call Main.SetCS(CSPointer, "TrackID", TrackID)
        '
        Call Main.SetCS(CSPointer, "Outcome", Outcome)
        Call Main.SetCS(CSPointer, "FirstName", FirstName)
        Call Main.SetCS(CSPointer, "LastName", LastName)
        Call Main.SetCS(CSPointer, "Title", Title)
        Call Main.SetCS(CSPointer, "Company", Company)
        Call Main.SetCS(CSPointer, "Address", Address)
        Call Main.SetCS(CSPointer, "Address2", Address2)
        Call Main.SetCS(CSPointer, "City", City)
        Call Main.SetCS(CSPointer, "State", State)
        Call Main.SetCS(CSPointer, "Zip", Zip)
        Call Main.SetCS(CSPointer, "Phone", Phone)
        Call Main.SetCS(CSPointer, "Fax", Fax)
        Call Main.SetCS(CSPointer, "Email", eMail)
        Call Main.SetCS(CSPointer, "AidDescription", AidDescription)
        Call Main.SetCS(CSPointer, "AudiovisualNotes", AudiovisualNotes)
        '
        Call Main.SetCS(CSPointer, "Accepted", Accepted)
        Call Main.SetCS(CSPointer, "AcceptedFor", AcceptedFor)
        Call Main.SetCS(CSPointer, "JointProposal", JointProposal)
        Call Main.SetCS(CSPointer, "JointProposalContact", JointProposalContact)
        Call Main.SetCS(CSPointer, "DateNotified", DateNotified)
        Call Main.SetCS(CSPointer, "Withdrawn", Withdrawn)
        Call Main.SetCS(CSPointer, "DateWithdrawn", DateWithdrawn)
        Call Main.SetCS(CSPointer, "ReviewedBy", ReviewedBy)
        Call Main.SetCS(CSPointer, "SupportLetter", SupportLetter)
        Call Main.SetCS(CSPointer, "SupportLetterComments", SupportLetterComments)
        Call Main.SetCS(CSPointer, "Handout", Handout)
        Call Main.SetCS(CSPointer, "HandoutComments", HandoutComments)
        Call Main.SetCS(CSPointer, "PostHandout", PostHandout)
        Call Main.SetCS(CSPointer, "PostHandoutComments", PostHandoutComments)
        Call Main.SetCS(CSPointer, "AlternateContact", AlternateContact)
        Call Main.SetCS(CSPointer, "SchoolCorrespondence", SchoolCorrespondence)
        Call Main.SetCS(CSPointer, "ProgramDate", ProgramDate)
        Call Main.SetCS(CSPointer, "ProgramTime", ProgramTime)
        Call Main.SetCS(CSPointer, "Location", Location)
        Call Main.SetCS(CSPointer, "Comments", Comments)
        Call Main.SetCS(CSPointer, "ParticipantCount", PostParticipantCount)
        Call Main.SetCS(CSPointer, "SessionRating", SessionRating)
        '
        End If
    Call Main.CloseCS(CSPointer)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "ProcessFormEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'GetAdditionalPresenters(SubmissionID)
'
Private Function GetAdditionalPresenters(SubmissionID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim Link As String
    Dim PresenterID As Long
    Dim RowNumber As Long
    Dim Class As String
    '
    CSPointer = Main.OpenCSContent(ContentNameSessionSubmissionsRules, "SessionSubmissionsID=" & Main.EncodeSQLNumber(SubmissionID))
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
        Stream = Stream & "<TR CLASS=""" & ClassPanelHeader & """>"
        Stream = Stream & "<TD>Name</TD>"
        Stream = Stream & "<TD>Title</TD>"
        Stream = Stream & "<TD>Organization</TD>"
        Stream = Stream & "<TD>eMail</TD>"
        Stream = Stream & "</TR>"
        Do While Main.CSOK(CSPointer)
            PresenterID = Main.GetCSInteger(CSPointer, "MemberID")
            Call LoadPresenterArray(PresenterID)
            '
            If (RowNumber Mod 2) = 0 Then
                Class = ClassPanelOdd
            Else
                Class = ClassPanelEven
                End If
            '
            Link = Main.ServerPage & WorkingQueryString & RequestNameSubmissionID & "=" & SubmissionID & "&" & RequestNameFormID & "=" & FormPeopleEdit & "&" & RequestNamePresenterID & "=" & PresenterID
            '
            Stream = Stream & "<TR CLASS=""" & Class & """>"
            Stream = Stream & "<TD><a href=""" & Link & """>" & PresenterArray(0) & "</a></TD>"
            Stream = Stream & "<TD>" & PresenterArray(1) & "</TD>"
            Stream = Stream & "<TD>" & PresenterArray(2) & "</TD>"
            Stream = Stream & "<TD>" & PresenterArray(3) & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
            RowNumber = RowNumber + 1
        Loop
        '
        Link = Main.ServerPage & WorkingQueryString & RequestNameSubmissionID & "=" & SubmissionID & "&" & RequestNameFormID & "=" & FormPeopleEdit & "&" & RequestNamePresenterID & "=0"
        '
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD colspan=4><a href=""" & Link & """><b>Add an additional presenter</b></a></TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "</TABLE>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetAdditionalPresenters = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetAdditionalPresenters", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadPresenterArray(MemberID As Long)
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    '
    PresenterArray(0) = ""
    PresenterArray(1) = ""
    PresenterArray(2) = ""
    PresenterArray(3) = ""
    '
    CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(MemberID))
    If Main.CSOK(CSPointer) Then
        PresenterArray(0) = Main.GetCSText(CSPointer, "Name")
        PresenterArray(1) = Main.GetCSText(CSPointer, "Title")
        PresenterArray(2) = Main.GetCSText(CSPointer, "Company")
        PresenterArray(3) = Main.GetCSText(CSPointer, "eMail")
        End If
    Call Main.CloseCS(CSPointer)
    '
    Exit Sub
ErrorTrap:
    Call HandleError("RFPReviewClass", "LoadPresenterArray", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Function GetFormPeopleEdit() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim ButtonString As String
    '
    ButtonString = Main.GetFormInputHidden(RequestNameSourceFormID, FormID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNameSubmissionID, SubmissionID)
    ButtonString = ButtonString & Main.GetFormInputHidden(RequestNamePresenterID, PresenterID)
    ButtonString = ButtonString & Main.GetFormButton(FormButtonUpdate)
    ButtonString = ButtonString & " &nbsp;" & Main.GetFormButton(FormButtonCancel)
    ButtonString = ButtonString & " &nbsp;" & Main.GetFormButton(FormButtonDelete)
    '
    Stream = Stream & Main.GetFormStart
    Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
    Stream = Stream & GetPageHead(PageNameFormParticipantEdit)
    Stream = Stream & GetFormRow("First Name", GetField(RequestNameFirstName, FirstName, , 50))
    Stream = Stream & GetFormRow("Last Name", GetField(RequestNameLastName, LastName, , 50))
    Stream = Stream & GetFormRow("Title", GetField(RequestNameTitle, Title, , 50))
    Stream = Stream & GetFormRow("College/Organization", GetField(RequestNameCompany, Company, , 50))
    Stream = Stream & GetFormRow("Address", GetField(RequestNameAddress, Address, , 50))
    Stream = Stream & GetFormRow3("", GetField(RequestNameAddress2, Address2, , 50))
    Stream = Stream & GetFormRow("City", GetField(RequestNameCity, City, , 50))
    Stream = Stream & GetFormRow("State", GetField(RequestNameState, State, , 50))
    Stream = Stream & GetFormRow("Zip", GetField(RequestNameZip, Zip, , 50))
    Stream = Stream & GetFormRow("Phone", GetField(RequestNamePhone, Phone, , 50))
    Stream = Stream & GetFormRow("Fax", GetField(RequestNameFax, Fax, , 50))
    Stream = Stream & GetFormRow("Email", GetField(RequestNameEmail, eMail, , 50))
    Stream = Stream & GetFormRow3(FormSpacer, ButtonString)
    Stream = Stream & "</TABLE>"
    Stream = Stream & Main.GetFormEnd
    '
    GetFormPeopleEdit = Stream
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetFormPeopleEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Sub LoadFormPeopleEdit()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim SpacePos As Long
    '
    If ButtonValue <> "" Then
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
        '
    Else
        '
        CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(PresenterID))
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
    Call HandleError("RFPReviewClass", "LoadFormPeopleEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ProcessFormPeopleEdit()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim LocalMemberID As Long
    '
    If PresenterID = 0 Then
        CSPointer = Main.InsertCSContent(ContentNamePeople)
    Else
        CSPointer = Main.OpenCSContent(ContentNamePeople, "ID=" & Main.EncodeSQLNumber(PresenterID))
        End If
    If Main.CSOK(CSPointer) Then
        LocalMemberID = Main.GetCSInteger(CSPointer, "ID")
        Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName)
        Call Main.SetCS(CSPointer, "FirstName", FirstName)
        Call Main.SetCS(CSPointer, "LastName", LastName)
        Call Main.SetCS(CSPointer, "Title", Title)
        Call Main.SetCS(CSPointer, "Company", Company)
        Call Main.SetCS(CSPointer, "Address", Address)
        Call Main.SetCS(CSPointer, "Address2", Address2)
        Call Main.SetCS(CSPointer, "City", City)
        Call Main.SetCS(CSPointer, "State", State)
        Call Main.SetCS(CSPointer, "Zip", Zip)
        Call Main.SetCS(CSPointer, "Phone", Phone)
        Call Main.SetCS(CSPointer, "Fax", Fax)
        Call Main.SetCS(CSPointer, "Email", eMail)
        Call Main.SaveCSRecord(CSPointer)
        '
        End If
    Call Main.CloseCS(CSPointer)
    '
    If PresenterID = 0 Then
        CSPointer = Main.InsertCSContent(ContentNameSessionSubmissionsRules)
        If Main.CSOK(CSPointer) Then
            Call Main.SetCS(CSPointer, "SessionSubmissionsID", SubmissionID)
            Call Main.SetCS(CSPointer, "MemberID", LocalMemberID)
            Call Main.SetCS(CSPointer, "Active", 1)
            Call Main.SaveCSRecord(CSPointer)
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("RFPReviewClass", "ProcessFormPeopleEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
Private Sub ExportSubmissions(ApprovedOnly As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Criteria As String
    Dim ExportStream As String
    Dim RecordCount As Long
    Dim AddressString As String
    Dim CurrentSubmissionID As Long
    Dim CS As Long
    Dim SQL As String
    '
    Call AppendLog(logfilename, "enter ExportSubmissions")
    '
    RecordCount = 1
    '
    If ApprovedOnly Then
        Criteria = "(Accepted=1) AND (Completed=1)"
    Else
        Criteria = "Completed=1"
    End If
    '
    Call Main.SetStreamTimeout(600)
    Call Main.ClearStream
    Call Main.SetStreamBuffer(False)
    Call Main.SetStreamType("application/msword")
    Call Main.SetStreamHeader("Content-Disposition", "attachment; filename=SubmissionReview.doc")
    '
    Call AppendLog(logfilename, "mail criteria=" & Criteria)
    CSPointer = Main.OpenCSContent(ContentNameSessionSubmissions, Criteria)
    If Main.CSOK(CSPointer) Then
        ExportStream = ExportStream & "<html>"
        ExportStream = ExportStream & "<body>"
        Do While Main.CSOK(CSPointer)
            Call AppendLog(logfilename, "submission [" & Main.GetCSText(CSPointer, "name") & "]")
            CurrentSubmissionID = Main.GetCSInteger(CSPointer, "ID")
            ExportStream = ExportStream & "<p style=""PAGE-BREAK-BEFORE: always;mso-pagination:none;mso-layout-grid-align:none;Text-autospace:none;"">"
            ExportStream = ExportStream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
            ExportStream = ExportStream & GetFormRow("Proposal ID", CStr(CurrentSubmissionID))
            ExportStream = ExportStream & GetFormRow2("&nbsp;&nbsp;&nbsp;&nbsp;<b>Key Contact Information</b>")
            
            SQL = "SELECT ccMembers.FirstName, ccMembers.LastName, ccMembers.Company, ccMembers.Title,"
            SQL = SQL & " ccMembers.Address, ccMembers.Address2, ccMembers.City, ccMembers.State, ccMembers.Zip, ccMembers.Email, ccMembers.Phone, ccMembers.Fax"
            SQL = SQL & " FROM SessionSubmissionsRules LEFT JOIN ccMembers ON SessionSubmissionsRules.MemberID = ccMembers.ID"
            SQL = SQL & " WHERE (SessionSubmissionsRules.ContentControlID=" & Main.GetRecordID("Content", ContentNameContactRules) & ")"
            SQL = SQL & " AND (SessionSubmissionsRules.Active=1) AND (ccMembers.Active=1)"
            SQL = SQL & " AND (SessionSubmissionsRules.SessionSubmissionsID=" & CurrentSubmissionID & ")"
            
            Call AppendLog(logfilename, "....member sql=" & SQL)
            CS = Main.OpenCSSQL("Default", SQL)
            Do While Main.CSOK(CS)
                Call AppendLog(logfilename, "....member [" & Main.GetCSText(CS, "firstname") & "]")
                AddressString = Main.GetCSText(CS, "Address")
                AddressString = AddressString & " " & Main.GetCSText(CS, "Address2") & "<br />"
                AddressString = AddressString & Main.GetCSText(CS, "City") & ", "
                AddressString = AddressString & Main.GetCSText(CS, "State") & " "
                AddressString = AddressString & Main.GetCSText(CS, "Zip")
                ExportStream = ExportStream & GetFormRow("Name", Main.GetCSText(CS, "FirstName") & " " & Main.GetCSText(CS, "LastName"))
                ExportStream = ExportStream & GetFormRow("Title", Main.GetCSText(CS, "Title"))
                ExportStream = ExportStream & GetFormRow("Organization", Main.GetCSText(CS, "Company"))
                ExportStream = ExportStream & GetFormRow("Address", AddressString)
                ExportStream = ExportStream & GetFormRow("Fax", Main.GetCSText(CS, "Fax"))
                ExportStream = ExportStream & GetFormRow("Phone", Main.GetCSText(CS, "Phone"))
                ExportStream = ExportStream & GetFormRow("Email", Main.GetCSText(CS, "Email"))
                ExportStream = ExportStream & GetFormRow2("&nbsp;")
                Call Main.NextCSRecord(CS)
            Loop
            Call Main.CloseCS(CS)
            
            ExportStream = ExportStream & GetFormRow2("&nbsp;&nbsp;&nbsp;&nbsp;<b>Program Information</b>")
            ExportStream = ExportStream & GetFormRow("Type", Main.GetCSLookup(CSPointer, "TypeID"))
            ExportStream = ExportStream & GetFormRow("Track", Main.GetCSLookup(CSPointer, "TrackID"))
            '
            ExportStream = ExportStream & GetFormRow("Title", Main.GetCSText(CSPointer, "SessionTitle"))
            ExportStream = ExportStream & GetFormRow("Description", Main.GetCSText(CSPointer, "Description"))
            ExportStream = ExportStream & GetFormRow("Outcome", Main.GetCSText(CSPointer, "Outcome"))
            '
            '
            ExportStream = ExportStream & GetFormRow("Audiovisual Description", Main.GetCSText(CSPointer, "AidDescription"))
'            ExportStream = ExportStream & GetFormRow("Audiovisual Notes", Main.GetCSText(CSPointer, "AudiovisualNotes"))
            '
            '
            ExportStream = ExportStream & GetFormRow2("&nbsp;&nbsp;&nbsp;&nbsp;<b>Additional Presenters</b>")
            ExportStream = ExportStream & GetFormRow2(GetExportPresenters(CurrentSubmissionID))
            ExportStream = ExportStream & GetFormRow3(FormSpacer2, "")
            ExportStream = ExportStream & "</TABLE>"
            ExportStream = ExportStream & "</p>"
            RecordCount = RecordCount + 1
            Call Main.NextCSRecord(CSPointer)
        Loop
        ExportStream = ExportStream & "</body>"
        ExportStream = ExportStream & "</html>"
    End If
    '
    Call Main.CloseCS(CSPointer)
    '
    'Call Main.SaveVirtualFile("RFPExport\Submission.doc", ExportStream)
    Call Main.WriteStream(ExportStream)
    Call Main.CloseStream
    '
    Call AppendLog(logfilename, "exit ExportSubmissions")
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call AppendLog(logfilename, "errortrap ExportSubmissions")
    Call HandleError("RFPReviewClass", "ExportSubmissions", Err.Number, Err.Source, Err.Description, True, False)
End Sub
'
Private Function GetExportPresenters(SubmissionID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    Dim Link As String
    Dim PresenterID As Long
    Dim RowNumber As Long
    Dim Class As String
    '
    CSPointer = Main.OpenCSContent(ContentNameSessionSubmissionsRules, "SessionSubmissionsID=" & Main.EncodeSQLNumber(SubmissionID))
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
        Stream = Stream & "<TR CLASS=""" & ClassPanelHeader & """>"
        Stream = Stream & "<TD>Name</TD>"
        Stream = Stream & "<TD>Title</TD>"
        Stream = Stream & "<TD>Organization</TD>"
        Stream = Stream & "<TD>eMail</TD>"
        Stream = Stream & "</TR>"
        Do While Main.CSOK(CSPointer)
            PresenterID = Main.GetCSInteger(CSPointer, "MemberID")
            Call LoadPresenterArray(PresenterID)
            '
            Link = Main.ServerPage & WorkingQueryString & RequestNameSubmissionID & "=" & SubmissionID & "&" & RequestNameFormID & "=" & FormPeopleEdit & "&" & RequestNamePresenterID & "=" & PresenterID
            '
            Stream = Stream & "<TR CLASS=""" & Class & """>"
            Stream = Stream & "<TD>" & PresenterArray(0) & "</TD>"
            Stream = Stream & "<TD>" & PresenterArray(1) & "</TD>"
            Stream = Stream & "<TD>" & PresenterArray(2) & "</TD>"
            Stream = Stream & "<TD>" & PresenterArray(3) & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
            RowNumber = RowNumber + 1
        Loop
        Stream = Stream & "</TABLE>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetExportPresenters = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("RFPReviewClass", "GetExportPresenters", Err.Number, Err.Source, Err.Description, True, False)
    End Function
