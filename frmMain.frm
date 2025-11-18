VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   ";RELEASE INFO!!!"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7965
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    End
End Sub

Private Sub btnConfig_Click()
    Functions.SaveSettings
    frmConfig.Show
End Sub

Private Sub btnHelp_Click()
    frmHelp.Show
End Sub

Private Sub btnOK_Click()
    Me.Hide
     If GetSetting(sPostName, "Settings", "sUnits", "") = "" Then
        ElseIf GetSetting(sPostName, "Settings", "sUnits", "") <> frmMain.cmbUnits.Text Then
            MsgBox "Units selection changed! Please re-select the post to allow the changes to take effect. Failure to do so will create bad code!"
            Call Functions.SaveSettings
            End
    End If
        
    Functions.SaveSettings

End Sub

Private Sub cmbOrigins_Change()
    Post.sOrigin = Me.cmbOrigins.Text
End Sub

Private Sub cmbUnits_Change()
    Post.sUnits = Me.cmbUnits.Text
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
    End
End Sub
