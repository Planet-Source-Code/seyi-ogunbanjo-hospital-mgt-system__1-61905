VERSION 5.00
Begin VB.Form frmUpdateHospNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Record with Hospital Number : "
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel Operation"
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
      TabIndex        =   48
      Top             =   5760
      Width           =   2175
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
      DatabaseName    =   ""
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
      RecordSource    =   "Contact_Info"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Record"
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
      Left            =   6000
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
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   37
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtStateOfOrigin 
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   45
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtLGA 
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   41
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtReligion 
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   43
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtNationality 
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
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
         DataSource      =   "datContactInfo"
         Height          =   285
         Left            =   1080
         TabIndex        =   30
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtKinOtherNames 
         DataSource      =   "datContactInfo"
         Height          =   285
         Left            =   4200
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtPhoneNok 
         DataSource      =   "datContactInfo"
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtAddressNok 
         DataSource      =   "datContactInfo"
         Height          =   1005
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox txtKinSName 
         DataSource      =   "datContactInfo"
         Height          =   285
         Left            =   840
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
         DataSource      =   "datContactInfo"
         Height          =   285
         Left            =   7920
         TabIndex        =   23
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtOfficeAdd 
         DataSource      =   "datContactInfo"
         Height          =   885
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtHomeAdd 
         DataSource      =   "datContactInfo"
         Height          =   885
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtHomePhone 
         DataSource      =   "datContactInfo"
         Height          =   285
         Left            =   7920
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
         DataSource      =   "datPerInfo"
         Height          =   315
         ItemData        =   "frmUpdateByHospNo.frx":0000
         Left            =   4560
         List            =   "frmUpdateByHospNo.frx":000A
         TabIndex        =   12
         Text            =   "F"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtOccupation 
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   7800
         TabIndex        =   14
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtDOB 
         DataField       =   "Date_of_Birth"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         ToolTipText     =   "DD/MM/YYYY"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtOName 
         DataField       =   "Middle_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   7800
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtFName 
         DataField       =   "First_Name"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtSName 
         DataField       =   "Surname"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtHospNo 
         DataField       =   "Hospital_No"
         DataSource      =   "datPerInfo"
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
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
Attribute VB_Name = "frmUpdateHospNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form locates a record according to hospital number
'and allows the user to edit it.
Dim flgCancelOp As Boolean

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo errHandler
With datPerInfo.Recordset
    'personal info
    .Fields("Hospital_No") = txtHospNo
    .Fields("Surname") = txtSName
    .Fields("First_Name") = txtFName
    .Fields("Middle_Name") = txtOName
    .Fields("Date_of_Birth") = txtDOB
    .Fields("Sex") = cboSex
    .Fields("Occupation") = txtOccupation
    'personal information2
    .Fields("Place_of_Birth") = txtPlaceOfBirth
    .Fields("Nationality") = txtNationality
    .Fields("Local_Govt_Area") = txtLGA
    .Fields("Religion") = txtReligion
    .Fields("State_of_Origin") = txtStateOfOrigin
 End With
 
 With datContactInfo.Recordset
    'Contact info
    .Fields("Hospital_No") = txtHospNo
    .Fields("Home_Address") = txtHomeAdd
    .Fields("Office_Address") = txtOfficeAdd
    .Fields("Home_Phone") = txtHomePhone
    .Fields("Office_Phone") = txtOfficePhone
    'Contact info2
    .Fields("surname_nok") = txtKinSName
    .Fields("First_Name_NoK") = txtKinOtherNames
    .Fields("relationship_to_nok") = txtRelationship
    .Fields("Phone_No_of_NoK") = txtPhoneNok
    .Fields("Address_of_NoK") = txtAddressNok
 End With
 datPerInfo.Recordset.Update
 datContactInfo.Recordset.Update
 MsgBox "Record Update Operation Completed Successfully.", vbInformation, "Congrats"
 Unload Me
 Exit Sub
errHandler:
 Select Case Err.Number
    Case 3316
        MsgBox "Invalid entry: Check the Sex field." & vbCrLf & "Note: All fields on this screen are required", vbCritical, "Input Error"
        cboSex.SetFocus
        Exit Sub
    Case 3421
        MsgBox "Invalid entry: Check the date of birth field." & vbCrLf & "Date of birth format is 'MM/DD/YYYY'" & vbCrLf & "Note: All fields on this screen are required", vbCritical, "Input Error"
        txtDOB.SetFocus
        Exit Sub
    Case 3315
        MsgBox "Invalid entry: Check the 'Other Information' section field." & vbCrLf & "Note: All fields on this screen are required", vbCritical, "Input Error"
        txtReligion.SetFocus
        Exit Sub
    Case Else
        MsgBox "Unhandled Error: Critical." & vbCrLf & "Note: All fields on this screen are required", vbCritical, "Input Error"
        Debug.Print Err.Number
        Debug.Print Err.Description
        Exit Sub
 End Select
 datPerInfo.Recordset.Update
 datContactInfo.Recordset.Update
 
End Sub

Private Sub datContactInfo_Validate(Action As Integer, Save As Integer)
 If flgCancelOp = True Then
    Save = 0
 End If
End Sub

Private Sub datPerInfo_Validate(Action As Integer, Save As Integer)
If flgCancelOp = True Then
    Save = 0
 End If
End Sub

Private Sub Form_Load()
 Me.Caption = Me.Caption + Str(hospNo)
 Dim strSQL As String
 Dim foundRec As Boolean
 Dim foundRec2 As Boolean
 '/
 'datPerInfo
 datPerInfo.DatabaseName = App.Path & "\alldb.mdb"
 datPerInfo.RecordSource = "personal_info"
 datPerInfo.Refresh
 
 'datContactInfo
 datContactInfo.DatabaseName = App.Path & "\alldb.mdb"
 datContactInfo.RecordSource = "contact_info"
 datContactInfo.Refresh
 
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
 
 datContactInfo.Recordset.MoveFirst
 With datContactInfo.Recordset
    Do
        If .Fields("hospital_no") = hospNo Then
            foundRec2 = True
        Else
            .MoveNext
        End If
    Loop Until (.EOF) Or (foundRec2)
 End With
 
 If (foundRec = False) And (foundRec2 = False) Then
    MsgBox "Record not found. Sorry!", , "Input Error"
    Unload Me
    Exit Sub
 End If
 
 datPerInfo.Recordset.Edit
 datContactInfo.Recordset.Edit
 flgCancelOp = False
End Sub

