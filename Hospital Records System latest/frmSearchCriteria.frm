VERSION 5.00
Begin VB.Form frmSearchCriteria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Specify Search Criteria"
   ClientHeight    =   2085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtSearchBySurname 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2655
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Search by &Patient's Surname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtSearchByHospNo 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Search by &Hospital Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearchCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
 Unload Me
End Sub

Private Sub OKButton_Click()
 Dim strSName As String     'name being sought
 Dim hospNo As Integer      'hosp_no being sought
 searchSucceeded = False
 
 If optSearch(1).Value = True Then
    'Search records by name
    strSName = Trim(txtSearchBySurname)
    If strSName = "" Then
        MsgBox "You need to enter a surname to search.", vbInformation, "Error"
        Exit Sub
    End If
    'code to search by surname
    MsgBox "Search Result.", vbInformation, "Surname Search"
 Else
    'Search records by hospital number
    hospNo = Val(txtSearchByHospNo)
    If hospNo = 0 Then
        MsgBox "You need to enter a hospital number to search.", vbInformation, "Error"
        Exit Sub
    End If
    'code to search by hospital_no
    frmSearchResult.datPatientInfo1.Recordset.FindFirst "Hospital_No = " & "'" & hospNo & "'"
    If frmSearchResult.datPatientInfo1.Recordset.NoMatch = True Then
        'search failed
        MsgBox "Search Failed. Try again", vbInformation, "Hospital No Search"
    Else
        'A record matching the hosp_no was found.
        searchSucceeded = True
        bkMk = frmSearchResult.datPatientInfo1.Recordset.Bookmark
        'Unload this form, so as to display record in the other "form".
        Unload Me
    End If
 End If
End Sub

Private Sub optSearch_Click(Index As Integer)
 If Index = 1 Then  'search by surname
    txtSearchByHospNo.Enabled = False
    txtSearchByHospNo.BackColor = &H80000013  'disabled color (brown)
    txtSearchBySurname.Enabled = True
    txtSearchBySurname.BackColor = &H80000005  'enabled color (white)
    txtSearchByHospNo = ""
    txtSearchBySurname.SetFocus
 Else   'search by hospital no
    txtSearchByHospNo.Enabled = True
    txtSearchByHospNo.BackColor = &H80000005  'enabled color (white)
    txtSearchBySurname = ""
    txtSearchBySurname.Enabled = False
    txtSearchBySurname.BackColor = &H80000013  'disabled color (brown)
    txtSearchByHospNo.SetFocus
 End If
End Sub
