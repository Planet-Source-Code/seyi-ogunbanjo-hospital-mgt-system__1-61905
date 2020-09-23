VERSION 5.00
Begin VB.Form frmHospHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hospital History for "
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datHistSQL 
      Caption         =   "Patient's Hospital History"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close This Window"
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   6600
      Width           =   3735
   End
   Begin VB.Data datPerInfo 
      Caption         =   "Personal Info"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame fraHospHistory 
      Caption         =   "&Hospital History of Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   3735
      Begin VB.TextBox txtDateAdmitted 
         DataSource      =   "datHistSQL"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtReferedBy 
         DataSource      =   "datHistSQL"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtPhySurgeon 
         DataSource      =   "datHistSQL"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtWardClinic 
         DataSource      =   "datHistSQL"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtDateDischarged 
         DataSource      =   "datHistSQL"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtDischargedTo 
         DataSource      =   "datHistSQL"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtDischargedStatus 
         DataSource      =   "datHistSQL"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Data datHospHistory 
         Caption         =   "Hospital History Table"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label8 
         Caption         =   "Date Admitted:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Referred By:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Physician/Surgeon:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Ward/Clinic:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Date Discharged:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Discharged To:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Discharged Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1575
      End
   End
   Begin VB.Frame fraPersonal 
      Caption         =   "&Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox cboSex 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         ItemData        =   "frmHospHistory.frx":0000
         Left            =   1560
         List            =   "frmHospHistory.frx":000A
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "F"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtDOB 
         DataField       =   "Date_of_Birth"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtOName 
         DataField       =   "Middle_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtFName 
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtSName 
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtHospNo 
         DataField       =   "Hospital_No"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblOName 
         Alignment       =   1  'Right Justify
         Caption         =   "Middle Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblFName 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblSName 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblHospNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Hospital Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmHospHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Dim strSQL As String
 Dim foundRec As Boolean
 Dim foundRec2 As Boolean
 'datPerInfo
 datPerInfo.DatabaseName = App.Path & "\alldb.mdb"
 datPerInfo.RecordSource = "personal_info"
 datPerInfo.Refresh

 'datDiagnosis
 datHospHistory.DatabaseName = App.Path & "\alldb.mdb"
 datHospHistory.RecordSource = "hospital_history"
 datHospHistory.Refresh
 
 'datHistSQL
 datHistSQL.DatabaseName = App.Path & "\alldb.mdb"
 strSQL = "select  personal_info.hospital_no, Date_Admitted, Referred_By, Physician_or_Surgeon, Ward_or_Clinic, Date_Discharged, Discharged_to, Status "
 strSQL = strSQL + "from hospital_history, personal_info where personal_info.hospital_no = hospital_history.hospital_no"
 datHistSQL.RecordSource = strSQL
 datHistSQL.Refresh
 
 txtDateAdmitted.DataField = "Date_Admitted"
 txtReferedBy.DataField = "Referred_By"
 txtPhySurgeon.DataField = "Physician_or_Surgeon"
 txtWardClinic.DataField = "Ward_or_Clinic"
 txtDateDischarged.DataField = "Date_Discharged"
 txtDischargedTo.DataField = "Discharged_to"
 txtDischargedStatus.DataField = "Status"
  
 datPerInfo.Recordset.MoveFirst
 With datPerInfo.Recordset
    Do
        If .Fields("hospital_no") = hospNo Then
            foundRec = True
        Else
            .MoveNext
        End If
    Loop Until (.EOF) Or (foundRec)
 End With
 
 datHospHistory.Recordset.MoveFirst
 With datHospHistory.Recordset
    Do
        If .Fields("hospital_no") = hospNo Then
            foundRec2 = True
        Else
            .MoveNext
        End If
    Loop Until (.EOF) Or (foundRec2)
 End With
 
 If (foundRec = False) Or (foundRec2 = False) Then
    MsgBox "The specified patient does not have any records in the hospital history table.", , "Hospital History"
    Unload Me
 End If
End Sub
