VERSION 5.00
Begin VB.Form frmDiagnosisHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnosis History for "
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datDiagSQL 
      Caption         =   "Diagnosis History"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Width           =   2580
   End
   Begin VB.Frame fraDiagnosis 
      Caption         =   "&Diagnosis Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   3735
      Begin VB.TextBox txtDiagID 
         DataField       =   "ID"
         DataSource      =   "datDiagnosis"
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtDiagDate 
         DataField       =   "Date"
         DataSource      =   "datDiagnosis"
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtDiagCodeNo 
         DataField       =   "Code_Number"
         DataSource      =   "datDiagnosis"
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtDiagnosis 
         DataField       =   "Diagnosis"
         DataSource      =   "datDiagnosis"
         Height          =   1005
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Data datDiagnosis 
         Caption         =   "Diagnosis Info DataBase"
         Connect         =   "Access"
         DatabaseName    =   " "
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Diagnosis"
         Top             =   3000
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.Label lblDiagID 
         Caption         =   "Diagnosis ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblDiagDate 
         Caption         =   "Diagnosis Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDiagCodeNo 
         Caption         =   "Code Number:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDiagnosis 
         Caption         =   "Diagnosis:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close This Window"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6600
      Width           =   3735
   End
   Begin VB.Data datPerInfo 
      Caption         =   "Personal Info"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Personal_Info"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
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
         ItemData        =   "frmDiagnosisHistory.frx":0000
         Left            =   1560
         List            =   "frmDiagnosisHistory.frx":000A
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "F"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtOccupation 
         DataField       =   "Occupation"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   1935
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
         Width           =   1935
      End
      Begin VB.TextBox txtOName 
         DataField       =   "Middle_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtFName 
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtSName 
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtHospNo 
         DataField       =   "Hospital_No"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1935
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
Attribute VB_Name = "frmDiagnosisHistory"
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
 datDiagnosis.DatabaseName = App.Path & "\alldb.mdb"
 datDiagnosis.RecordSource = "diagnosis"
 datDiagnosis.Refresh
 
 'datDiagSQL
 datDiagSQL.DatabaseName = App.Path & "\alldb.mdb"
 strSQL = "select  personal_info.hospital_no, id, Date, Code_Number, Diagnosis "
 strSQL = strSQL + "from diagnosis, personal_info where personal_info.hospital_no = diagnosis.hospital_no"
 datDiagSQL.RecordSource = strSQL
 datDiagSQL.Refresh
 
 txtDiagID.DataField = "id"
 txtDiagDate.DataField = "Date"
 txtDiagCodeNo.DataField = "Code_Number"
 txtDiagnosis.DataField = "Diagnosis"
  
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
 
 datDiagnosis.Recordset.MoveFirst
 With datDiagnosis.Recordset
    Do
        If .Fields("hospital_no") = hospNo Then
            foundRec2 = True
        Else
            .MoveNext
        End If
    Loop Until (.EOF) Or (foundRec2)
 End With
 
 If (foundRec = False) Or (foundRec2 = False) Then
    MsgBox "The specified patient does not have any records in the diagnosis history table.", , "Hospital History"
    Unload Me
 End If
End Sub

