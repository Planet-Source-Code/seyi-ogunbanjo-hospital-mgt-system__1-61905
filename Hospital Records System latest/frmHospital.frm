VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHospital 
   Caption         =   "Hospital Records System"
   ClientHeight    =   9060
   ClientLeft      =   75
   ClientTop       =   765
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   13410
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   1191
      ButtonWidth     =   609
      ButtonHeight    =   1032
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbrMainForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8685
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Status: ""NOT Logged In"""
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   7080
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      Height          =   9495
      Left            =   -120
      TabIndex        =   1
      Top             =   840
      Width           =   15975
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu mnuUserLogon 
         Caption         =   "&Log in"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuUserLogOut 
         Caption         =   "&Log Out"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserExit 
         Caption         =   "&Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Enabled         =   0   'False
      Begin VB.Menu mnuToolsAddNew 
         Caption         =   "&Add A New Patient Record"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsEditRec 
         Caption         =   "&Edit/Update an existing record"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuToolsSearch 
         Caption         =   "&Search for an existing record"
         Begin VB.Menu mnuSearchHospNo 
            Caption         =   "By Patient's Hospital Number"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuSearchSurname 
            Caption         =   "By Patient's Surname"
            Shortcut        =   ^{F3}
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHospHist 
         Caption         =   "View Patient's &Hospital History..."
      End
      Begin VB.Menu mnuDiagHist 
         Caption         =   "View Patient's &Diagnosis History..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsViewExistingRec 
         Caption         =   "&View All Records"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutHelp 
         Caption         =   "&Help"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAboutSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutThisApp 
         Caption         =   "&About Hospital Records System"
      End
   End
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 LoginSucceeded = False
 With App
    Me.Caption = .Title & " " & .Major & "." & .Minor & "." & .Revision
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim strMsg As String
 If LoginSucceeded = True Then
    'Prevent the user from exitting the app without logging - out first.
    strMsg = "You need to log out before the application can close." + vbCrLf
    strMsg = strMsg + "You can log out using the User menu, or Ctrl+O."
    MsgBox strMsg, vbExclamation
    Cancel = 1
 Else
    'Confirm before quitting
    If MsgBox("Quit to windows?", vbYesNo + vbQuestion, "Quit?") = vbYes Then
        Cancel = 0
        End
    Else
        Cancel = 1
    End If
 End If
End Sub

Private Sub mnuAboutThisApp_Click()
 Load frmAbout
 frmAbout.Show 1
End Sub

Private Sub mnuDiagHist_Click()
On Error GoTo errHnd
 hospNo = Val(InputBox("Please enter the HOSPITAL NUMBER:"))
 Load frmDiagnosisHistory
 frmDiagnosisHistory.Show 1
errHnd:
 If Err.Number = 364 Then
    Debug.Print Err.Description
    Exit Sub
 End If
End Sub

Private Sub mnuHospHist_Click()
On Error GoTo errHnd
 hospNo = Val(InputBox("Please enter the HOSPITAL NUMBER:"))
 Load frmHospHistory
 frmHospHistory.Show 1
errHnd:
 If Err.Number = 364 Then
    Debug.Print Err.Description
    Exit Sub
 End If
End Sub

Private Sub mnuSearchHospNo_Click()
On Error GoTo errHnd
 hospNo = Val(InputBox("Please enter the HOSPITAL NUMBER:"))
 Load frmSearchResultHospNo
 frmSearchResultHospNo.Show 1
errHnd:
 If Err.Number = 364 Then
    Debug.Print Err.Description
    Exit Sub
 End If
End Sub

Private Sub mnuSearchSurname_Click()
'traps errors without terminatin program without user knowledge
 On Error GoTo errHnd
 'trim removes any leading or trailing spaces since db doesn't contain that
 strSName = Trim(InputBox("Please enter the patient's SURNAME:"))
 Load frmSearchResultsn
 frmSearchResultsn.Show 1

errHnd:
 If Err.Number = 364 Then
    'Unloading a form b4 it was loaded generates this error 364
    'in this case the form is frmSearchResultsn which will be unloaded
    'if the search using surname was unsuccesful
    Debug.Print Err.Description
    ' puts the error description in the immediate window- a debuggin tool in vb
    ' occurs when record is not found
    Exit Sub
 End If
End Sub

Private Sub mnuToolsAddNew_Click()
 Load frmNewPatientReg1
 frmNewPatientReg1.Show 1

 'Clear all input controls (text boxes, etc)
End Sub

Private Sub mnuToolsEditRec_Click()
On Error GoTo errHnd
 hospNo = Val(InputBox("Please enter the HOSPITAL NUMBER of record to update:"))
 Load frmUpdateHospNo
 frmUpdateHospNo.Show 1
errHnd:
 If Err.Number = 364 Then
    Debug.Print Err.Description
    Exit Sub
 End If
End Sub

Private Sub mnuToolsViewExistingRec_Click()
 Load frmViewExisting
 frmViewExisting.Show 1
End Sub

Private Sub mnuUserExit_Click()
 Unload Me
End Sub

Private Sub mnuUserLogon_Click()
 Load frmLogin
 frmLogin.Show 1
End Sub

Private Sub mnuUserLogOut_Click()
 If MsgBox("Log out of the system?", vbYesNo + vbInformation, "Logout?") = vbNo Then Exit Sub
 Call Unload_Startup_Screen
End Sub

