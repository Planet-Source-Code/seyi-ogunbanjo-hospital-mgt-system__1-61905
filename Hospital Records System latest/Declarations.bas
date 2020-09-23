Attribute VB_Name = "Declarations"
 Public strSName As String     'name being sought
 Public hospNo As Integer      'hosp_no being sought
 Public bkMk As String
 Public searchSucceeded As Boolean
 Public LoginSucceeded As Boolean
Option Explicit
Public Sub Load_Startup_Screen(strUserName As String)
 'Initialize all environment variables.
 With frmHospital
    .mnuUserLogon.Enabled = False
    .mnuUserLogOut.Enabled = True
    .mnuUserExit.Enabled = False
    .mnuTools.Enabled = True
    .mnuWindow.Enabled = True
    If strUserName = "RECORDS STAFF" Then
        .mnuDiagHist.Enabled = False
    ElseIf strUserName = "DOCTOR" Then
        .mnuDiagHist.Enabled = True
    End If
    
    'Refresh the status bar
    .sbrMainForm.SimpleText = "User Name: " & strUserName _
      & "                Status: LOGGED-IN"
    
 End With
 LoginSucceeded = True
End Sub

Public Sub Unload_Startup_Screen()
 'Close all environment variables.
 With frmHospital
    .mnuUserLogon.Enabled = True
    .mnuUserLogOut.Enabled = False
    .mnuUserExit.Enabled = True
    .mnuTools.Enabled = False
    .mnuWindow.Enabled = False
    
    'Refresh the status bar
    .sbrMainForm.SimpleText = "Status: LOGGED-OUT"
    
 End With
 LoginSucceeded = False
End Sub

Public Sub ClrRegForm1()
 'Clear all input controls (text boxes, etc)
 

 With frmNewPatientReg1
    'personal info (7)
    '.txtHospNo = ""
    .txtSName = ""
    .txtFName = ""
    .txtOName = ""
    .txtDOB = ""
    .txtOccupation = ""
    .cboSex.Text = ""
 
    'contact info(5)
    .txtHomeAdd = ""
    .txtOfficeAdd = ""
    .txtHomePhone = ""
    .txtOfficePhone = ""
 
    'next of kin info(5)
    .txtKinSName = ""
    .txtRelationship = ""
    .txtKinOtherNames = ""
    .txtAddressNok = ""
    .txtPhoneNok = ""
     
    'other info(5)
    .txtPlaceOfBirth = ""
    .txtNationality = ""
    .txtLGA = ""
    .txtReligion = ""
    .txtStateOfOrigin = ""
    .txtSName.SetFocus
 End With
 
End Sub
