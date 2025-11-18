VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfig 
   Caption         =   "Biesse CIX Post Configuration"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   OleObjectBlob   =   "frmConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutoLabels_Click()

    If Me.chkAutoLabels.Value Then
        Post.bFlagAutoLabel = True
        Me.lblExportFileExtension.Enabled = True
        Me.cmbImageType.Enabled = True
        
        Me.lblOutlineNote.Enabled = True
        Me.txtOutlineNote.Enabled = True
    Else
        Post.bFlagAutoLabel = False
        Me.lblExportFileExtension.Enabled = False
        Me.cmbImageType.Enabled = False
        
        Me.lblOutlineNote.Enabled = False
        Me.txtOutlineNote.Enabled = False
    End If
End Sub

Private Sub chkHideGUI_Click()

    If chkHideGUI.Value Then
        Post.bHideGUI = True
    Else
        Post.bHideGUI = False
    End If
    
End Sub

Private Sub chkIgnoreCustomUserDataString_Click()

    If Me.chkIgnoreCustomUserDataString.Value Then
        Post.bIgCustString = True
        Me.chkUseCustomUserString.Enabled = False
        Me.txtUseCustomUserString.Enabled = False
    Else
        Post.bIgCustString = False
        Me.chkUseCustomUserString.Enabled = True
        Me.txtUseCustomUserString.Enabled = True
        
        If Me.chkUseCustomUserString.Value Then
            Me.txtUseCustomUserString.Enabled = True
            Post.bUseCustomUserSt = True
        Else
            Me.txtUseCustomUserString.Enabled = False
            Post.bUseCustomUserSt = False
        End If
    
    End If
    
End Sub

Private Sub chkLeadMsgs_Click()
    If chkLeadMsgs.Value Then
        Post.bSupLeadMsgs = True
    Else
        Post.bSupLeadMsgs = False
    End If
End Sub

Private Sub chkMachineFeeds_Click()

    If chkMachineFeeds.Value Then
        Post.bUseCNCFeeds = True
    Else
        Post.bUseCNCFeeds = False
    End If

End Sub

Private Sub chkUseCustomUserString_Click()

    If Me.chkUseCustomUserString.Value Then
        Me.txtUseCustomUserString.Enabled = True
        Post.bUseCustomUserSt = True
    Else
        Me.txtUseCustomUserString.Enabled = False
        Post.bUseCustomUserSt = False
    End If
    
End Sub

Private Sub cmbImageType_Change()

    Post.sImageType = Me.cmbImageType.Text

End Sub



Private Sub UserForm_Activate()

    Me.txtUseCustomUserString.Text = GetSetting(DEF_POST_NAME, "Settings", "txtUseCustomUserString", "")
    Me.txtXCUT.Text = GetSetting(DEF_POST_NAME, "Settings", "txtXCUT", "0")
    Me.txtYCUT.Text = GetSetting(DEF_POST_NAME, "Settings", "txtYCUT", "0")

    
    
    Functions.GetSettings
    
    Me.cmbImageType.Clear
    Me.cmbImageType.AddItem "BMP: Bit Map File"
    Me.cmbImageType.AddItem "JPG: Joint Photo Experts Group"
    Me.cmbImageType.AddItem "EMF: Enhanced Metafile"
    Me.cmbImageType.AddItem "WMF: Windows Metafile"
    Me.cmbImageType.AddItem "GIF: Graphic Illustration File"
    Me.cmbImageType.Text = GetSetting(DEF_POST_NAME, "Settings", "cmbImageType", "BMP: Bit Map File")

End Sub

Private Sub UserForm_Terminate()

    Post.sCustomUserStr = Me.txtUseCustomUserString.Text
    Post.sXCUT = Me.txtXCUT.Text
    Post.sYCUT = Me.txtYCUT.Text
    Post.sOutlineNote = Me.txtOutlineNote.Text
    Post.sImageType = Me.cmbImageType.Text
    
    Functions.SaveSettings
    
End Sub
