VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AwardNominationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const RequestNameFirstName = "firstname"
Const RequestNameLastName = "lastname"
Const RequestNameEmail = "email"
Const RequestNamePhone = "phone"
Const RequestNameInstitution = "institution"
Const RequestNameNominationForm = "nominationform"
Const RequestNameBoardLetter = "boardletter"
Const RequestNameNominationPacket = "nominationpacket"
Const RequestNamePhotograph = "photograph"
Const RequestNameAwardCategory = "awardcategory"
'
Const FormButtonSubmit = "Submit Nomination"
'
Dim CompletedFlag As Boolean
'
Private ErrorMessage As String
'
Private Main As Object

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("AwardNominationClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'----------------------------------------------
'   Aggregate Object Interface
'----------------------------------------------
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    If Not (Main Is Nothing) Then
        GetContent = GetNominationForm()
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("AwardNominationClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'----------------------------------------------
'   Returns sign up form
'----------------------------------------------
'
Private Function GetNominationForm() As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As Long
    '
    Dim Institution As String
    Dim CategoryString As String
    Dim FirstName As String
    Dim LastName As String
    Dim Phone As String
    Dim eMail As String
    '
    Dim NominationForm As String
    Dim BoardLetter As String
    Dim NominationPacket As String
    Dim Photograph As String
    '
    Call ProcessNominationForm
    '
    If ErrorMessage <> "" Then
        Institution = Main.GetStreamText(RequestNameInstitution)
        CategoryString = Main.GetStreamText(RequestNameAwardCategory)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Phone = Main.GetStreamText(RequestNamePhone)
        eMail = Main.GetStreamText(RequestNameEmail)
    Else
        CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
        If Main.CSOK(CSPointer) Then
            FirstName = Main.GetCSText(CSPointer, "FirstName")
            LastName = Main.GetCSText(CSPointer, "LastName")
            Institution = Main.GetCSText(CSPointer, "Company")
            Phone = Main.GetCSText(CSPointer, "Phone")
            eMail = Main.GetCSText(CSPointer, "eMail")
            End If
        Call Main.CloseCS(CSPointer)
        End If
    '
Call Main.TestPoint("CategoryString: " & CategoryString)
    '
    If CompletedFlag Then
        Stream = Stream & Main.GetContentPage("Award Nomination Thank You", , , False, False, 0)
    Else
        Stream = Stream & Main.GetUploadFormStart()
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
        If ErrorMessage <> "" Then
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD colspan=2 class=""ccError"">" & ErrorMessage & "</TD>"
            Stream = Stream & "</TR>"
            End If
        '
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Name of Institution:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameInstitution, Institution, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Award Categories:</TD>"
        Stream = Stream & "<TD width=100%>" & GetAwardCategories(CategoryString) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>First Name:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameFirstName, FirstName, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Last Name:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameLastName, LastName, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Phone Number:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNamePhone, Phone, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>eMail Address:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputText(RequestNameEmail, eMail, , 25) & "</TD>"
        Stream = Stream & "</TR>"
        '
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Nomination Form:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputFile(RequestNameNominationForm) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Board Letter of Support:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputFile(RequestNameBoardLetter) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Nomination Packet:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputFile(RequestNameNominationPacket) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right>Photograph:</TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormInputFile(RequestNamePhotograph) & "</TD>"
        Stream = Stream & "</TR>"
        '
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD align=right><img src=""/ccLib/images/spacer.gif"" height=1 width=150></TD>"
        Stream = Stream & "<TD width=100%>" & Main.GetFormButton(FormButtonSubmit) & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "</TABLE>"
        Stream = Stream & Main.GetUploadFormEnd()
        End If
    '
    GetNominationForm = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("AwardNominationClass", "GetNominationForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
' Processes the Form
'
Private Sub ProcessNominationForm()
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Institution As String
    Dim CategoryString As String
    Dim FirstName As String
    Dim LastName As String
    Dim Phone As String
    Dim eMail As String
    '
    Dim NominationForm As String
    Dim BoardLetter As String
    Dim NominationPacket As String
    Dim Photograph As String
    '
    Dim NominationID As Long
    Dim Filename As String
    Dim Copy As String
    '
    Dim AtPosition As Double
    Dim Dotposition As Double
    '
    If Main.GetStreamText("Button") = FormButtonSubmit Then
        '
Call Main.TestPoint("Nomination Form Submitted")
        '
        Institution = Main.GetStreamText(RequestNameInstitution)
        CategoryString = Main.GetStreamText(RequestNameAwardCategory)
        FirstName = Main.GetStreamText(RequestNameFirstName)
        LastName = Main.GetStreamText(RequestNameLastName)
        Phone = Main.GetStreamText(RequestNamePhone)
        eMail = Main.GetStreamText(RequestNameEmail)
        '
        NominationForm = Main.GetStreamText(RequestNameNominationForm)
        BoardLetter = Main.GetStreamText(RequestNameBoardLetter)
        NominationPacket = Main.GetStreamText(RequestNameNominationPacket)
        Photograph = Main.GetStreamText(RequestNamePhotograph)
        '
        If Institution = "" Then
            ErrorMessage = ErrorMessage & "Please enter the name of your institution<br />"
            End If
        '
        If FirstName = "" Then
            ErrorMessage = ErrorMessage & "Please enter your first name<br />"
            End If
        '
        If LastName = "" Then
            ErrorMessage = ErrorMessage & "Please enter your last name<br />"
            End If
        '
        AtPosition = InStr(1, eMail, "@")
        If AtPosition = 0 Then
            AtPosition = 1
            End If
        Dotposition = InStr(AtPosition, eMail, ".")
        '
        If (Dotposition < AtPosition + 2) Or (AtPosition = 1) Then
            ErrorMessage = ErrorMessage & "Please enter a valid eMail address<br />"
            End If
        If Phone = "" Then
            ErrorMessage = ErrorMessage & "Please enter your phone number<br />"
            End If
        '
        If NominationForm = "" Then
            ErrorMessage = ErrorMessage & "Please attach a copy of the nomination form<br />"
            End If
        If BoardLetter = "" Then
            ErrorMessage = ErrorMessage & "Please attach a copy of the board letter<br />"
            End If
        'If NominationPacket = "" Then
            'ErrorMessage = ErrorMessage & "Please attach a copy of the nomination packet<br />"
            'End If
        'If Photograph = "" Then
            'ErrorMessage = ErrorMessage & "Please attach a copy of the Photograph<br />"
            'End If
        '
        If ErrorMessage = "" Then
            '
            Call Main.TestPoint("Error Message is empty")
            '
            CSPointer = Main.OpenCSContent("People", "ID=" & Main.EncodeSQLNumber(Main.MemberID))
            If Main.CSOK(CSPointer) Then
                Call Main.SetCS(CSPointer, "FirstName", FirstName)
                Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName)
                Call Main.SetCS(CSPointer, "LastName", LastName)
                Call Main.SetCS(CSPointer, "Company", Institution)
                Call Main.SetCS(CSPointer, "Email", eMail)
                Call Main.SetCS(CSPointer, "Phone", Phone)
                Call Main.SetCS(CSPointer, "AllowBulkEmail", 1)
                Call Main.SaveCSRecord(CSPointer)
                End If
            Call Main.CloseCS(CSPointer)
            '
            Call Main.AddGroupMember("Award Nominators")
            '
            CSPointer = Main.InsertCSContent("Award Nominations")
            If Main.CSOK(CSPointer) Then
                NominationID = Main.GetCSInteger(CSPointer, "ID")
                '
                Call Main.SetCS(CSPointer, "VisitID", Main.VisitID)
                Call Main.SetCS(CSPointer, "MemberID", Main.MemberID)
                '
                Filename = Main.GetCSFilename(CSPointer, "NominationFormFileName", NominationForm, "Award Nominations")
                Filename = Replace(Filename, NominationForm, "")
                Call Main.ProcessFormInputFile(RequestNameNominationForm, Filename)
                '
                Filename = Main.GetCSFilename(CSPointer, "BoardLetterFileName", BoardLetter, "Award Nominations")
                Filename = Replace(Filename, BoardLetter, "")
                Call Main.ProcessFormInputFile(RequestNameBoardLetter, Filename)
                '
                Filename = Main.GetCSFilename(CSPointer, "NominationPacketFileName", NominationPacket, "Award Nominations")
                Filename = Replace(Filename, NominationPacket, "")
                Call Main.ProcessFormInputFile(RequestNameNominationPacket, Filename)
                '
                Filename = Main.GetCSFilename(CSPointer, "Photograph", Photograph, "Award Nominations")
                Filename = Replace(Filename, Photograph, "")
                Call Main.ProcessFormInputFile(RequestNamePhotograph, Filename)
                '
                Call Main.SetCS(CSPointer, "Name", FirstName & " " & LastName & " " & Now())
                Call Main.SetCS(CSPointer, "Institution", Institution)
                Call Main.SetCS(CSPointer, "AwardCategory", CategoryString)
                Call Main.SetCS(CSPointer, "FirstName", FirstName)
                Call Main.SetCS(CSPointer, "LastName", LastName)
                Call Main.SetCS(CSPointer, "Email", eMail)
                Call Main.SetCS(CSPointer, "Phone", Phone)
                '
                Call Main.SaveCSRecord(CSPointer)
                '
                Copy = "The following information was submitted on " & Main.ServerHost & " at " & Now() & ".<br /><br />"
                Copy = Copy & "<b>Institution:</b> " & Institution & "<br />"
                Copy = Copy & "<b>Award Category:</b> " & CategoryString & "<br />"
                Copy = Copy & "<b>Name:</b> " & FirstName & " " & LastName & "<br />"
                Copy = Copy & "<b>eMail:</b> " & eMail & "<br />"
                Copy = Copy & "<b>Phone:</b> " & Phone & "<br />"
                Copy = Copy & "<b>Nomination Form:</b> <a href=""http://" & Main.ServerHost & Main.ServerFilePath & Main.GetCS(CSPointer, "NominationFormFileName") & """>Click here to view the nomination form</a><br />"
                Copy = Copy & "<b>Board Letter:</b> <a href=""http://" & Main.ServerHost & Main.ServerFilePath & Main.GetCS(CSPointer, "BoardLetterFileName") & """>Click here to view the board letter</a><br />"
                If NominationPacket <> "" Then
                    Copy = Copy & "<b>Nomination Packet:</b> <a href=""http://" & Main.ServerHost & Main.ServerFilePath & Main.GetCS(CSPointer, "NominationPacketFileName") & """>Click here to view the nomination packet</a><br />"
                    End If
                If Photograph <> "" Then
                    Copy = Copy & "<b>Photograph:</b> <a href=""http://" & Main.ServerHost & Main.ServerFilePath & Main.GetCS(CSPointer, "Photograph") & """>Click here to view the photograph</a><br />"
                    End If
                Copy = Copy & "<a target=_blank href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?cid=" & Main.GetContentID("Award Nominations") & "&af=4&id=" & NominationID & """>Click here to view the nomination record</a>"
                '
                Call Main.SendSystemEmail("Award Nomination Form Auto Responder", , Main.MemberID)
                Call Main.SendSystemEmail("Award Nomination Notification", Copy)
                '
                CompletedFlag = True
                '
                End If
            Call Main.CloseCS(CSPointer)
            End If
        '
    Else
        Call Main.TestPoint("Button Not Pushed")
        End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("AwardNominationClass", "ProcessNominationForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'
'
Private Function GetAwardCategories(SelectedCategoryString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim Stream As String
    Dim Count As Long
    Dim CategoryName As String
    Dim CheckFlag As Boolean
    '
    Count = 0
    '
    CSPointer = Main.OpenCSContent("Award Categories", , "Name")
    If Main.CSOK(CSPointer) Then
        Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>"
        Do While Main.CSOK(CSPointer)
            CategoryName = Main.GetCSText(CSPointer, "Name")
            CheckFlag = InStr(1, SelectedCategoryString, CategoryName, vbTextCompare)
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD><input type=checkbox "
            If CheckFlag Then
                Stream = Stream & "checked "
                End If
            Stream = Stream & "value=""" & CategoryName & """ name=" & RequestNameAwardCategory & "></TD>"
            Stream = Stream & "<TD width=100%>" & CategoryName & "</TD>"
            Stream = Stream & "</TR>"
            Call Main.NextCSRecord(CSPointer)
            Count = Count + 1
        Loop
        Stream = Stream & "</TABLE>"
        End If
    Call Main.CloseCS(CSPointer)
    '
    GetAwardCategories = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("AwardNominationClass", "GetAwardCategories", Err.Number, Err.Source, Err.Description, True, False)
    End Function
