Attribute VB_Name = "Events"
Option Explicit
Public sCNCName                 As String
Public bNestingCalls            As Boolean

Function InitAlphacamAddIn(AcamVersion As Long) As Integer
                
    Dim FSO
    Dim sTitle                  As String
    Dim sICO                    As String
        
    On Error GoTo EHandler
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    With App.Frame

        sTitle = "BiesseCIX_USA"
                        
        If .AddMenuItem4("BiesseCIX_USA", "hh_BiesseCIX", acamCmdFILE_INSERT, True, sTitle) Then
                sICO = App.Frame.PathOfThisAddin & "\BiesseCIX_USA.bmp"
                If FSO.FileExists(sICO) Then
                  Call .AddButton(acamButtonBarFILE, sICO, .LastMenuCommandID)
                End If
        End If

    End With
            
    InitAlphacamAddIn = 0

    Exit Function

EHandler:

MsgBox Err.Description & " - Err No: " & Err.Number & " -InitAlphacamAddIn BiesseCIX_USA"

End Function

Function hh_BiesseCIX()
    
    Functions.GetSettings
    Functions.fReadVolMat
    
    frmMain.Caption = DEF_REL_INFO
    frmMain.Show
    
    Call Functions.SaveSettings
    
End Function

Function BeforeOutputNc()

    If VBA.InStr(App.PostFileName, DEF_POST_NAME) < 1 Then Exit Function
        
    bNestingCalls = False
    bPromptDone = False
    
    Dim sFileName As String
    Dim sFilter As String
    Dim sOutput As String
    
    sOutput = ""
    
    CMD.DialogTitle = "Biesse CIX Post Output"
    CMD.AllowMultiSelect = False
    sFilter = "Biesse CIX Files (*.cix)" & Chr(0) & "*.cix" & Chr(0)
    CMD.Filter = sFilter
    
    If App.ActiveDrawing.Name <> "" Then
        sOutput = sOutput & Trim(App.ActiveDrawing.Name) & ".cix"
        CMD.FileName = sOutput
    Else
        CMD.FileName = "Untitled.cix"
    End If
    
    If CMD.ShowSave = False Then
        BeforeOutputNc = 2
        Exit Function
    End If
    
    If CMD.FileName <> "" Then
        sFileName = CMD.FileName
        
        Select Case Right(sFileName, 3)
        
        Case "cix"
            sFileName = Left(sFileName, Len(sFileName) - 3) & "cix"
        Case Else
            sFileName = sFileName & ".cix"
        End Select
        
        BeforeOutputNc = sFileName
        sCNCName = CMD.FileTitle
    Else
        BeforeOutputNc = 2
    End If

'    If CBool(GetSetting(DEF_POST_NAME, "Settings", "chkFlagForAutoLabel", "0")) Then
'        'Create lbl info files
'        If Post.bFlagAutoLabel Then
'            Functions.AddAutoLabelInfo
'        End If
'    End If

End Function

Function BeforeCreateNC()
    'BeforeCreateNC = 0
    
    'License check
'        If Not Licensing.GetLicense Then
'            MsgBox "This Post is not licensed, please obtain a licensed copy!"
'            BeforeCreateNC = 1
'            Exit Function
'        End If
       


'        'Post name check / Option to copy post name to clipboard
'        Dim Response As String
'        If VBA.InStr(App.PostFileName, sPostNameControl) < 1 Then
'            MsgBox "Error! The name of the Post has been modified from the original! Please rename the post to " & sPostNameControl, vbCritical
'            Response = MsgBox("Copy Post Name To Clipboard ? " & sPostNameControl, vbYesNo + vbQuestion + vbDefaultButton2)
'                If Response = 6 Then    ' User chose Yes = 6 , No = 7
'                    Call Events.CopyPostNameToClipboard
'                Else
'                    MsgBox (sPostNameControl + " was not copied to your clipboard, Please close Alphacam and correct the name of your post.")
'                End If
'            BeforeCreateNC = 1
'            Exit Function
'        Else
             BeforeCreateNC = 0
'        End If
    
        'License date check
        If bCheckDate Then
            If Not Licensing.PstExpirey Then
                MsgBox "This Post license has expired, please obtain a licensed copy!"
                BeforeCreateNC = 1
                Exit Function
            ElseIf (Date + 7) >= strExpiryDate Then
                MsgBox "WARNING: Post Processor will expire on " & strExpiryDate & vbCrLf & ", contact Alphacam."
                BeforeCreateNC = 0
            End If
        End If
End Function
Public Sub CopyPostNameToClipboard()
    Dim myData As DataObject
    Dim Output As String
    
    Output = sPostNameControl
    
    If Output = "" Then
    End
    Else
    End If
    Set myData = New DataObject
    myData.SetText Output
    myData.PutInClipboard
    MsgBox (Output + " text has been copied , Please close Alphacam and correct the name of your post.")
End Sub
Function BeforeOutputNcDialogBox()
    
    BeforeOutputNcDialogBox = 1 ' Output to File Only, Suppress OutputNC Dialog Box

End Function
Function AfterOutputNc(sFileName As String)
    
   If VBA.InStr(App.PostFileName, DEF_POST_NAME) < 1 Then Exit Function
 
    App.ActiveDrawing.Attribute("Biesse_USA_CIX") = DEF_REL_INFO

    bPromptDone = False
    
    If CBool(GetSetting(DEF_POST_NAME, "Settings", "chkFlagForAutoLabel", "0")) Then
        Call Functions.AddAutoLabelInfo(sFileName)
    End If
    
    Call SplitNestedPrograms(sFileName)
    
End Function

Public Sub AfterRoughFinish(PS As Paths, Redo As Integer)
    
    Dim p           As Path
    Dim eleOverlap  As Element
    Dim eleInitial  As Element
    Dim stx         As Double
    Dim sty         As Double
    Dim endx        As Double
    Dim endy        As Double
    
    'Logic: Overlap will ALWAYS be immediately before a Lead-Out (if exists)
    'or overlap will ALWAYS be LAST element of the path (non rapid)
    'Once a candidate for overlap is found, then compare its Start Coordinates
    'to the End Coordinates of the FIRST Non-rapid, non-lead element.
    'if they match, flag it!
    
    For Each p In App.ActiveDrawing.ToolPaths
        If p.GetMillData.ProcessType2 = acamProcessROUGH_FINISH Then
            If p.GetMillData.Bidirectional Then GoTo NextP
            If p.Elements.Count = 1 And p.Elements(1).IsRapid Then GoTo NextP
            Set eleOverlap = p.GetLastElem
            While eleOverlap.IsRapid
                Set eleOverlap = eleOverlap.GetPrevious
            Wend
            While eleOverlap.LeadOut
                Set eleOverlap = eleOverlap.GetPrevious
            Wend
            stx = eleOverlap.StartXL
            sty = eleOverlap.StartYL
            
            Set eleInitial = p.GetFirstElem
            While eleInitial.IsRapid
                Set eleInitial = eleInitial.GetNext
            Wend
            While eleInitial.LeadIn
                Set eleInitial = eleInitial.GetNext
            Wend
            endx = eleInitial.StartXL
            endy = eleInitial.StartYL
            
            If (stx = endx) And (sty = endy) And Not eleInitial.IsSame(eleOverlap) Then
                eleOverlap.Attribute("_hhhOverlapElement") = 1
            End If
        End If
NextP:
    Next p

End Sub


