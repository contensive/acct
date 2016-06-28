VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "payPalConfirmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private Main As Object
Private Csv As Object
'
Public Function Execute(CsvObject As Object, MainObj As Object, OptionString As String, FilterInput As String)
    '
    Set Main = MainObj
    '
    Dim SubmissionID As Integer
    Dim amountPaid As Double
    Dim okToCompleteRegistration As Boolean
    Dim echoString As String
    Dim echoVar() As String
    Dim CS As Long
    '
    okToCompleteRegistration = True
    '
    echoString = Main.GetStreamText("echo")
    If echoString = "" Then
        Call AppendLogFile("PayPal Confirmation hit received but item_name was empty. It should be the echoString with RFPID and AmountPaid.")
        okToCompleteRegistration = False
    Else
        echoVar = Split(echoString, ",")
        If UBound(echoVar) <> 1 Then
            Call AppendLogFile("PayPal Confirmation hit received but item_name was invalid [" & echoString & "]. It should be the echoString with RFPID and AmountPaid.")
            okToCompleteRegistration = False
        Else
            SubmissionID = kmaEncodeInteger(echoVar(0))
            amountPaid = KmaEncodeNumber(echoVar(1))
            If SubmissionID = 0 Then
                Call AppendLogFile("PayPal Confirmation hit received but item_name returned and invalid PrimaryAttendeeID (first arg of echostring). item_name [" & echoString & "]")
                okToCompleteRegistration = False
            ElseIf amountPaid = 0 Then
                Call AppendLogFile("PayPal Confirmation hit received but item_name returned and invalid amountPaid (second arg of echostring). item_name [" & echoString & "]")
                okToCompleteRegistration = False
            Else
                If okToCompleteRegistration Then
                    Call CompleteApp(SubmissionID)
                End If
            End If
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("acct.PayPalConfirmClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
Private Sub CompleteApp(SubmissionID As Integer)
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim LocalQueryString As String
    Dim Copy As String
    '
    CSPointer = Main.OpenCSContent("Session Submissions", "ID=" & Main.EncodeSQLNumber(SubmissionID))
    If Main.CSOK(CSPointer) Then
        Call Main.SetCS(CSPointer, "Completed", 1)
        Call Main.SaveCSRecord(CSPointer)
    End If
    Call Main.CloseCS(CSPointer)
    '
    LocalQueryString = "cid=" & Main.GetContentID("Session Submissions")
    LocalQueryString = LocalQueryString & "&id=" & SubmissionID
    LocalQueryString = LocalQueryString & "&af=4"
    '
    Copy = "<br /><br /><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & "admin/index.asp?" & LocalQueryString & """>Click here to view the submission</a>"
    '
    Call Main.AddGroupMember("Session Submissions")
    '
    Call Main.SendSystemEmail("Session Submission Auto Responder", , Main.MemberID)
    Call Main.SendSystemEmail("Session Submission Notification", Copy)
    '
    Exit Sub
ErrorTrap:
    Call HandleError("RFPSubmissionClass", "CompleteApp", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
