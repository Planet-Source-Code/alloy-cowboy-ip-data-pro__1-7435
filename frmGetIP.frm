VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmGetIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Alloy Cowboy's IP Data Pro"
   ClientHeight    =   2790
   ClientLeft      =   2385
   ClientTop       =   1530
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetIP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstDupes 
      Height          =   270
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4200
      Width           =   735
   End
   Begin VB.Frame fraPrefs 
      Caption         =   "User Preferences :"
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Width           =   6495
      Begin VB.CommandButton cmdNoPrefs 
         Caption         =   "×"
         Height          =   255
         Left            =   6120
         MouseIcon       =   "frmGetIP.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Click here too hide the User Preferences."
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   5160
         MouseIcon       =   "frmGetIP.frx":0594
         MousePointer    =   99  'Custom
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Click here too browse for a cowboy response script."
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkBeep 
         Caption         =   "Beep when new data is collected."
         Height          =   255
         Left            =   2040
         MouseIcon       =   "frmGetIP.frx":06E6
         MousePointer    =   99  'Custom
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox txtResponse 
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         MousePointer    =   12  'No Drop
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "80"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblResponse 
         Caption         =   "Server Response :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblPort 
         Caption         =   "Listening Port :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   6000
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraControls 
      Caption         =   "User Controls :"
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   6495
      Begin VB.CommandButton cmdPrefs 
         Caption         =   "Preferences"
         Height          =   375
         Left            =   5160
         MouseIcon       =   "frmGetIP.frx":0838
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Click here too view the User Preferences."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   375
         Left            =   4320
         MouseIcon       =   "frmGetIP.frx":098A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Click here too copy the server URL."
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtURL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2520
         Locked          =   -1  'True
         MousePointer    =   12  'No Drop
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "http:// Your IP Address"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdDisabled 
         Caption         =   "Disabled"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         MouseIcon       =   "frmGetIP.frx":0ADC
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Click here too disable the server."
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdEnabled 
         Caption         =   "Enabled"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGetIP.frx":0C2E
         MousePointer    =   99  'Custom
         TabIndex        =   1
         ToolTipText     =   "Click here too enable the server."
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblURL 
         Caption         =   "URL :"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraDataView 
      Caption         =   "Link Data Collected :"
      Height          =   1935
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6495
      Begin ComctlLib.ListView lstDataView 
         Height          =   1575
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2778
         View            =   3
         Arrange         =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   12582912
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "IP Address"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Time Logged"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Browser Version"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "System OS"
            Object.Width           =   2277
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGetIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'* Original concept formulated by Plasma:     *
'* andrewarmstrong@hotmail.com                *
'* Reworked and revamped by the Alloy Cowboy: *
'* ThrillKillKid20@aol.com                    *
'* Copyright © 2000 - Alloy Cowboy INC.       *
'* Prosecutors will be violated.              *
'**********************************************

Option Explicit

'Socket counter'
Dim iSockets As Integer
'Data sent flag array'
Dim BufferEmpty(0 To 666) As Boolean

Private Sub cmdBrowse_Click()

    'Browse for .cow response script'
    On Error Resume Next

    CommonDialog1.FilterIndex = 5
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = &H1000
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "cowboy response script (*.cow)|*.cow"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen

    If Err Then
        If Err.Number = 32755 Then: Exit Sub
    End If

    txtResponse.Text = CommonDialog1.FileName

End Sub

Private Sub cmdCopy_Click()

    'Copies server URL to the clipboard'
    On Error Resume Next
    'Exit if app is disabled'
    If txtURL.Text = "http:// Your IP Address" Then
        Beep
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText txtURL.Text
        MsgBox "The server URL has been copied to the Windows clipboard.", _
        vbApplicationModal + vbInformation + vbOKOnly, " URL COPIED"
    End If

End Sub

Private Sub cmdDisabled_Click()

    Dim i As Integer

    On Error GoTo ErrHandler
    
    'Close main socket'
    sckServer(0).Close
    
    'Close all other open sockets'
    For i% = 1 To iSockets%
        If sckServer(i%).State <> sckClosed Then sckServer(i%).Close
    Next i%
    
    'Reset data sent flag array'
    For i% = 0 To 666
        If BufferEmpty(i%) <> False Then BufferEmpty(i%) = False
    Next i%
    
    'Reset User Controls URL text box'
    txtURL.Text = "http:// Your IP Address"
    
    'Enable User Prefs port text box'
    txtPort.Enabled = True
    
    'Enable User Prefs Browse button'
    cmdBrowse.Enabled = True
    
    'Disable User Controls Disable button'
    cmdDisabled.Enabled = False
    
    'Enable User Controls Enable button'
    cmdEnabled.Enabled = True
    
    'Set focus to User Controls Enable button'
    cmdEnabled.SetFocus
    
    Exit Sub
    
ErrHandler:

    'Error message box sub'
    Call ErrorMsgOut(Err.Number, Err.Description)

End Sub

Private Sub cmdEnabled_Click()

    On Error GoTo ErrHandler
    
    'Set socket to listen to user port'
    sckServer(0).LocalPort = txtPort.Text
    sckServer(0).Listen
    
    'Setup User Controls URL text box'
    If txtPort.Text = 80 Then
        txtURL.Text = "http://" & sckServer(0).LocalIP
    Else
        txtURL.Text = "http://" & sckServer(0).LocalIP & ":" & txtPort.Text
    End If
    
    'Disable User Prefs port text box'
    txtPort.Enabled = False
    
    'Disable User Prefs Browse button'
    cmdBrowse.Enabled = False
    
    'Disable User Controls Enable button'
    cmdEnabled.Enabled = False
    
    'Enable User Controls Disable button'
    cmdDisabled.Enabled = True
    
    'Set focus to User Controls Disable button'
    cmdDisabled.SetFocus
    
    Exit Sub
    
ErrHandler:

    'Error message box sub'
    Call ErrorMsgOut(Err.Number, Err.Description)
    
End Sub

Private Sub cmdNoPrefs_Click()

    On Error Resume Next
    'Disable tab key for prefs'
    txtPort.TabStop = False
    chkBeep.TabStop = False
    cmdNoPrefs.TabStop = False
    cmdBrowse.TabStop = False
    'Reduce form height to hide prefs'
    frmGetIP.Height = 3165
    'Set focus to link data list'
    lstDataView.SetFocus
    'Enable User Controls prefs button'
    cmdPrefs.Enabled = True

End Sub

Private Sub cmdPrefs_Click()

    On Error Resume Next
    'Enable tab key for prefs'
    txtPort.TabStop = True
    chkBeep.TabStop = True
    cmdNoPrefs.TabStop = True
    cmdBrowse.TabStop = True
    'Increase form height to show prefs'
    frmGetIP.Height = 4360
    'Set focus to port text box'
    txtPort.SetFocus
    'Disable User Controls prefs button'
    cmdPrefs.Enabled = False

End Sub

Private Sub ErrorMsgOut(iErr As Integer, sErr As String)

    Dim retVal As Integer

    On Error Resume Next
    
    'Displays appropriate error message box'
    If iErr% = 0 Or sErr$ = "" Then
        retVal% = MsgBox("An error has occurred !" & vbCrLf & _
        "Program may not function properly." & vbCrLf & "Continue ?" _
        , vbApplicationModal + vbCritical + vbDefaultButton2 + vbYesNo, " ERROR")
    Else
        retVal% = MsgBox("Error #" & iErr% & " has occurred !" & vbCrLf & sErr$ & "." & _
        vbCrLf & "Program may not function properly." & vbCrLf & "Continue ?" _
        , vbApplicationModal + vbCritical + vbDefaultButton2 + vbYesNo, " ERROR")
    End If
    
    'Check return value, 7 = No'
    If retVal% = 7 Then cmdDisabled_Click

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    'Welcome user, give credit and contact info'
    MsgBox "- Thank you for choosing Alloy Cowboy's IP Data Pro !" & vbCrLf & _
    "- The original concept was formulated by Plasma." & vbCrLf & _
    "- The source was reworked and revamped by the Alloy Cowboy." & vbCrLf & vbCrLf & _
    "- Contact Plasma :" & vbCrLf & "  andrewarmstrong@hotmail.com" & vbCrLf & vbCrLf & _
    "- Contact the Alloy Cowboy :" & vbCrLf & "  ThrillKillKid20@aol.com" _
    , vbApplicationModal + vbInformation + vbOKOnly, " WELCOME !"
    
    'Setup form'
    Me.Caption = " Alloy Cowboy's IP Data Pro - Version " & App.Major & "." & App.Minor & "." & App.Revision & " !"
    txtResponse.Text = App.Path & "\errmsg.cow"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    'Terminate application'
    End

End Sub

Private Sub sckServer_Close(Index As Integer)

    On Error GoTo ErrHandler

    'Check if main socket is being closed'
    If Index% <> 0 Then
        'Close other sockets'
        sckServer(Index%).Close
        Unload sckServer(Index%)
        iSockets% = iSockets% - 1
    Else
        'Close main socket.
        sckServer(Index%).Close
    End If
    
    Exit Sub
    
ErrHandler:

    'Error message box sub'
    Call ErrorMsgOut(Err.Number, Err.Description)

End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    On Error GoTo ErrHandler
    
    If Index% = 0 Then
        'Load a new socket to accept the new connection'
        iSockets% = iSockets% + 1
        Load sckServer(iSockets%)
        'Sets a random port for the new connection'
        sckServer(iSockets%).LocalPort = 0
        'Accept the new connection'
        sckServer(iSockets%).Accept requestID&
    End If
    
    Exit Sub
    
ErrHandler:

    'Error message box sub'
    Call ErrorMsgOut(Err.Number, Err.Description)
    
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim sBrowser As String
    Dim sData As String, sIP As String
    Dim vCow As Variant
    Dim OSData, OSData1, OSData2
    Dim SplitData, i As Integer
    
    On Error GoTo ErrHandler
    
    'Get incoming data'
    sckServer(Index%).GetData sData$, vbString
    
    'Divide data into 1 dimensional arrays'
    SplitData = Split(sData$, ";")
    OSData1 = Split(sData$, vbCrLf)
    
    'Look for system information line'
    For i% = LBound(OSData1) To UBound(OSData1)
        If Left(OSData1(i%), 11) = "User-Agent:" Then
            OSData2 = Split(OSData1(i%), ";")
            Exit For
        End If
    Next i%
    
    'Get operating system version'
    If Right(OSData2(2), 1) = ")" Then
        OSData = Left(OSData2(2), Len(OSData2(2)) - 1)
    Else
        OSData = OSData2(2)
    End If

    'Get web browser version'
    sBrowser$ = OSData2(1)
    
    'Get remote host IP address'
    sIP$ = sckServer(Index%).RemoteHostIP
    
    'Open selected .cow response script'
    Open txtResponse.Text For Input As #1
        vCow = StrConv(InputB(LOF(1), 1), vbUnicode)
    Close #1
    
    'Send reply'
    sckServer(Index%).SendData vCow & Chr$(10)
    
    'Wait until all data is sent'
    Do: DoEvents
    Loop Until BufferEmpty(Index%) = True
    
    'Reset data sent flag'
    BufferEmpty(Index%) = False
    
    'Close the socket'
    sckServer(Index%).Close
    
    'Write obtained data to text file'
    Open App.Path & "\debug.txt" For Output As #1
        Print #1, sData$ & vbCrLf
    Close #1
    
    'Check for duplicate IP address'
    For i% = 0 To lstDupes.ListCount - 1
        If lstDupes.List(i%) = sIP$ Then
            Exit Sub
        End If
    Next i%
    
    'Add new IP to duplicate check list'
    lstDupes.AddItem sIP$
    
    'Add new info to link data list'
    With lstDataView
        .ListItems.Add = sIP$
        .ListItems.Item(.ListItems.Count).SubItems(1) = Format(Now, "Long Time")
        .ListItems.Item(.ListItems.Count).SubItems(2) = sBrowser$
        .ListItems.Item(.ListItems.Count).SubItems(3) = OSData
    End With
    
    'Beep if set too'
    If chkBeep.Value = 1 Then Beep
    
    Exit Sub
    
ErrHandler:

    'Error message box sub'
    Call ErrorMsgOut(Err.Number, Err.Description)

End Sub

Private Sub sckServer_SendComplete(Index As Integer)

    'Flags a completed send operation'
    On Error Resume Next
    BufferEmpty(Index%) = True

End Sub
