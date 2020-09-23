VERSION 5.00
Begin VB.Form frmNewPatientReg2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Patient Registration - II"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1575
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   10335
      Begin VB.TextBox txtHospNo 
         DataField       =   "Hospital_No"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtSName 
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtFName 
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   36
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtOName 
         DataField       =   "Middle_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   7800
         TabIndex        =   35
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtDOB 
         DataField       =   "Date of Birth"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7800
         TabIndex        =   33
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cboSex 
         DataField       =   "Sex"
         DataSource      =   "datPerInfo"
         Height          =   315
         ItemData        =   "frmNewPatientReg2.frx":0000
         Left            =   4560
         List            =   "frmNewPatientReg2.frx":000A
         TabIndex        =   32
         Text            =   "F"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblHospNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Hospital Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblSName 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblFName 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblOName 
         Alignment       =   1  'Right Justify
         Caption         =   "Middle Name:"
         Height          =   255
         Left            =   6600
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   3360
         TabIndex        =   40
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   6600
         TabIndex        =   39
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   8760
      TabIndex        =   30
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6600
      TabIndex        =   29
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   8760
      TabIndex        =   28
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6600
      TabIndex        =   27
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame fraBloodTest 
      Caption         =   "&Blood Test Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6600
      TabIndex        =   22
      Top             =   1920
      Width           =   3975
      Begin VB.TextBox Text14 
         Height          =   1245
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Blood Test Result:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Blood Test Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraOperations 
      Caption         =   "&Operation Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   6255
      Begin VB.TextBox txtOpCodeNo 
         Height          =   285
         Left            =   4680
         TabIndex        =   15
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4200
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   885
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Data Data1 
         Caption         =   "Lab Info DataBase"
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
         Top             =   2040
         Width           =   5940
      End
      Begin VB.Label lblOpCodeNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Operation Code No:"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Surgeon:"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Operation:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Operation ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Operation:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame fraLabInfo 
      Caption         =   "&Laboratory Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   6255
      Begin VB.Data datLab_Info 
         Caption         =   "Lab Info DataBase"
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
         Top             =   1440
         Width           =   5940
      End
      Begin VB.TextBox txtLabGenotype 
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtLabRefNo 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtLabBloodGroup 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtLabRhesus 
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtLabAllergy 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblLabGenotype 
         Alignment       =   1  'Right Justify
         Caption         =   "Genotype:"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblLabRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Lab. Ref. Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLabBloodGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabRhesus 
         Alignment       =   1  'Right Justify
         Caption         =   "Rhesus:"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblLabAllergy 
         Alignment       =   1  'Right Justify
         Caption         =   "Allergy:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmNewPatientReg2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
