VERSION 5.00
Begin VB.Form frmSearchResultsn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Result for Patient named "
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRecCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   47
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data datPerInfo 
      Caption         =   "Personal Info"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Data datContactInfo 
      Caption         =   "Contact Info"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Data datPatientInfo1 
      Caption         =   "Patient Info"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Personal_Info"
      Top             =   360
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   46
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Other Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7080
      TabIndex        =   35
      Top             =   3480
      Width           =   3495
      Begin VB.TextBox txtPlaceOfBirth 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtStateOfOrigin 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtLGA 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtReligion 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtNationality 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Place of Birth:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "State Of Origin:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Govt. Area:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Religion:"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Nationality:"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame fraNextOfKin 
      Caption         =   "&Next of Kin's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   24
      Top             =   3480
      Width           =   6735
      Begin VB.TextBox txtRelationship 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtKinOtherNames 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtPhoneNok 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtAddressNok 
         DataSource      =   "datPatientInfo1"
         Height          =   1005
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox txtKinSName 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Relationship:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Other Names:"
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Phone Number:"
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblKinSName 
         Caption         =   "Surname:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraContact 
      Caption         =   "&Contact Information"
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
      TabIndex        =   15
      Top             =   1800
      Width           =   10335
      Begin VB.TextBox txtOfficePhone 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtOfficeAdd 
         DataSource      =   "datPatientInfo1"
         Height          =   885
         Left            =   4080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtHomeAdd 
         DataSource      =   "datPatientInfo1"
         Height          =   885
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtHomePhone 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Office Phone Number:"
         Height          =   255
         Left            =   7920
         TabIndex        =   22
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblOfficeAdd 
         Caption         =   "Office Address:"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblHomeAdd 
         Caption         =   "Home Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Home Phone Number:"
         Height          =   255
         Left            =   7920
         TabIndex        =   20
         Top             =   240
         Width           =   2175
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
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.ComboBox cboSex 
         DataField       =   "Sex"
         DataSource      =   "datPatientInfo1"
         Height          =   315
         ItemData        =   "frmSearchResult.frx":0000
         Left            =   4560
         List            =   "frmSearchResult.frx":000A
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "F"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtOccupation 
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtDOB 
         DataField       =   "Date_of_Birth"
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtOName 
         DataField       =   "Middle_Name"
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtFName 
         DataField       =   "First_Name"
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtSName 
         DataField       =   "Surname"
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtHospNo 
         DataField       =   "Hospital_No"
         DataSource      =   "datPatientInfo1"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex:"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblOName 
         Alignment       =   1  'Right Justify
         Caption         =   "Middle Name:"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblFName 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblSName 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblHospNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Hospital Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSearchResultsn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form locates a record by surname

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = Me.Caption + UCase(strSName)
  Dim strSQL As String
 Dim criteria As String
 '/
 'datPerInfo
 datPerInfo.DatabaseName = App.Path & "\alldb.mdb"
 datPerInfo.RecordSource = "personal_info"
 datPerInfo.Refresh
 
 'datContactInfo
 datContactInfo.DatabaseName = App.Path & "\alldb.mdb"
 datContactInfo.RecordSource = "contact_info"
 datContactInfo.Refresh
 
 'datPerInfo.Recordset.Bookmark = bkMk
 
 'datPatientInfo1
 'datPatientInfo1 .DatabaseName = App.Path & "\alldb.mdb"
 strSQL = "select personal_info.hospital_no, Surname, First_Name, Middle_Name, Date_of_Birth, Sex, Place_of_Birth, Local_Govt_Area, Nationality, Religion, State_of_Origin, "
 strSQL = strSQL + "Home_Address, Office_Address, Home_Phone, Office_Phone, Occupation, surname_nok, First_Name_NoK, relationship_to_nok, Address_of_NoK, Phone_No_of_NoK "
 strSQL = strSQL + "from contact_info, personal_info where contact_info.hospital_no = personal_info.hospital_no"
 datPatientInfo1.RecordSource = strSQL
 datPatientInfo1.Refresh
 
 'contact_info
 'next_of_kin, other_personal_info,
 ' and personal_info.hospital_no = next_of_kin.hospital_no and personal_info.hospital_no = other_personal_info.hospital_no
 
 'personal info
 txtHospNo.DataField = "Hospital_No"
 txtSName.DataField = "Surname"
 txtFName.DataField = "First_Name"
 txtOName.DataField = "Middle_Name"
 txtDOB.DataField = "Date_of_Birth"
 cboSex.DataField = "Sex"
 'personal information2
 txtPlaceOfBirth.DataField = "Place_of_Birth"
 txtLGA.DataField = "Local_Govt_Area"
 txtNationality.DataField = "Nationality"
 txtReligion.DataField = "Religion"
 txtStateOfOrigin.DataField = "State_of_Origin"
 
 'Contact info
 txtHomeAdd.DataField = "Home_Address"
 txtOfficeAdd.DataField = "Office_Address"
 txtHomePhone.DataField = "Home_Phone"
 txtOfficePhone.DataField = "Office_Phone"
 txtOccupation.DataField = "Occupation"
 'Contact info2
 txtKinSName.DataField = "surname_nok"
 txtKinOtherNames.DataField = "First_Name_NoK"
 txtRelationship.DataField = "relationship_to_nok"
 txtAddressNok.DataField = "Address_of_NoK"
 txtPhoneNok.DataField = "Phone_No_of_NoK"
 
 'to locate record by surname
 criteria = "surname = " & "'" & strSName & "'"
 datPatientInfo1.Recordset.FindFirst criteria
 If datPatientInfo1.Recordset.NoMatch Then
    MsgBox "Record not found. Sorry!", , "Search Failed"
    Unload Me
 End If

 '//
End Sub

