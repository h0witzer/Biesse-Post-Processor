Attribute VB_Name = "Functions"
Option Explicit
Option Private Module

Public dMidX  As Double
Public dMidY  As Double

Public Sub SplitNestedPrograms(ByRef sFileName As String)

If VBA.InStr(1, App.PostFileName, DEF_POST_NAME) < 1 Then Exit Sub

If gb_HasNesting Then  'Section for Nesting ONLY

    Dim iSheetCounter As Integer

    Dim sTempFile           As String
    Dim sBuffer             As String
    Dim sSheetName          As String
    Dim sLine2              As String
    Dim sLine3              As String
    Dim sLine4              As String
    Dim sLabelFileName     As String
    Dim sLabelBuff          As String
    
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TriStateUseDefault = -2, TriStateTrue = -1, TriStateFalse = 0
    
    Dim FSO, TheFile, TempFile, LabelFile
    Set FSO = CreateObject("Scripting.FileSystemObject")
                      
    Set TheFile = FSO.OpenTextFile(sFileName, ForReading, TriStateFalse)
    
    iSheetCounter = 0
    
    While Not TheFile.AtEndOfStream
            
        sBuffer = TheFile.ReadLine

        If Trim(sBuffer) = sFirstLine Then
        
            iSheetCounter = iSheetCounter + 1
            
            'Open Label file corresponding to this sheet
            If Post.bFlagAutoLabel Then
                Dim sFileExt As String
                sFileExt = VBA.UCase(VBA.Right(sFileName, 4))
            
                Select Case sFileExt
                    Case ".CIX"
                        sLabelFileName = VBA.Left(sFileName, VBA.Len(sFileName) - 4) & ".autolabel" & CStr(iSheetCounter)
                    Case ".ANC"
                        sLabelFileName = VBA.Left(sFileName, VBA.Len(sFileName) - 4) & ".autolabel" & CStr(iSheetCounter)
                    Case Else
                        sLabelFileName = sFileName & ".autolabel" & CStr(iSheetCounter)
                End Select
                
                Set LabelFile = FSO.OpenTextFile(sLabelFileName, ForReading, TriStateFalse)
            End If
        
            'Sheet Name should be on the FOURTH line of code.
            sLine2 = TheFile.ReadLine
            
            'Read next line
            sLine3 = TheFile.ReadLine
            
            'Read sheet name
            sLine4 = TheFile.ReadLine
            
            'Copy over
            sSheetName = sLine4
            
            'Create proper file name based on sheet name
            sSheetName = VBA.Right(sSheetName, VBA.Len(sSheetName) - VBA.InStr(sSheetName, "Sheet_") - 5)
            
            'sTempFile = VBA.Left(sFileName, VBA.Len(sFileName) - 4) & "_" & sSheetName & "_" & CStr(iSheetCounter) & "(1).cix"
            sTempFile = VBA.Left(sFileName, VBA.Len(sFileName) - 4) & "_" & CStr(iSheetCounter) & "(1).cix"
            
            Set TempFile = FSO.CreateTextFile(sTempFile, True, False)
            TempFile.Close

            Set TempFile = FSO.OpenTextFile(sTempFile, ForAppending, True, TriStateFalse)
            
            'insert lines 1,2,3,4
            TempFile.WriteLine sBuffer
            TempFile.WriteLine sLine2
            TempFile.WriteLine sLine3
            TempFile.WriteLine sLine4
            
        Else
            If sBuffer <> sLastLine Then
                'Write any line NOT the last
                TempFile.WriteLine sBuffer
            Else
                'Last Line.
                TempFile.WriteLine sBuffer
                
                If CBool(GetSetting(DEF_POST_NAME, "Settings", "chkFlagForAutoLabel", "0")) Then
                    'Append Label Data
                    While Not LabelFile.AtEndOfStream
                        sLabelBuff = LabelFile.ReadLine
                        TempFile.WriteLine sLabelBuff
                    Wend
                    LabelFile.Close
                
                    'Kill Label File
                    Kill sLabelFileName
                End If
                                
                'Close file
                TempFile.Close
            End If
        End If
    Wend
    
    TempFile.Close
    TheFile.Close
    
    Kill sFileName

End If
    
End Sub

Public Function gb_HasNesting() As Boolean
        
    On Error GoTo ErrTrap

    Dim blnRet As Boolean
        blnRet = False
    
    Dim oNestInformation        As NestInformation
    Set oNestInformation = App.ActiveDrawing.GetNestInformation
    
    If Not oNestInformation Is Nothing Then
        If oNestInformation.Sheets.Count > 0 Then
            blnRet = True
        End If
    End If
           
    gb_HasNesting = blnRet
    
    Exit Function

ErrTrap:

    If Err.Number = -2147467259 Then
        Resume Next
    Else
        gb_HasNesting = False
        Exit Function
    End If

End Function
Public Sub fReadVolMat(Optional ByVal p As PostData)

    On Error GoTo eh
    
    'AlphaBug... Material looses properties in V8 if Constrained, this line of code fixes this.
    Dim W As Workpiece
    Set W = App.ActiveDrawing.Workpiece
    
    Dim Drw As Drawing
    Set Drw = App.ActiveDrawing
        
    Dim TheEle As Element
    Dim TheMat As Path
    Dim TopZ As Double
    Dim BottomZ As Double
    Dim minx As Double
    Dim miny As Double
    Dim maxx As Double
    Dim maxy As Double
    Dim mix As Double, miy As Double, max As Double, may As Double, maz As Double, miz As Double
    Dim bMatExists As Boolean
    Dim bMatLyrExists As Boolean
    Dim wp As WorkPlane
     
    TopZ = -99999
    BottomZ = 99999
    mix = 99999
    miy = 99999
    max = -99999
    may = -99999
    bMatExists = False
    bMatLyrExists = False
    
    Dim AttBotZ As String
    AttBotZ = "LicomUKDMBGeoZLevelBottom"
    
    Dim AttTopZ As String
    AttTopZ = "LicomUKDMBGeoZLevelTop"
    
    If gb_HasNesting Then
        If Not p Is Nothing Then
            If p.Vars.NSH > 0 Then
            
                mix = 0
                max = Round(p.Vars.SHX, iDecimals)
                miy = 0
                may = Round(p.Vars.SHY, iDecimals)
                
                If p.Vars.SHZ > 0 Then
                    TopZ = 0
                    BottomZ = -Round(p.Vars.SHZ, 4)
                    bMatExists = True
                    GoTo done
                Else
                
                    For Each TheMat In Drw.Geometries
                        
                        Post.dZ = 0
                        
                        'Look for Sheets by Attirbute
                        If TheMat.Attribute("LicomUKsab_sheet_thickness") <> "" Then
                            Post.dZ = TheMat.Attribute("LicomUKsab_sheet_thickness")
                            Exit For
                        End If
                        
                        'Look for Material to reference Thickness
                        If TheMat.Attribute("LicomUKDMBStockType") = "1" Or _
                            TheMat.Attribute("LicomUKDMBStockType") = "2" Or _
                            TheMat.Name = "Material" Then
                    
                            TopZ = TheMat.Attribute(AttTopZ)
                            BottomZ = TheMat.Attribute(AttBotZ)
                            Post.dZ = Round(Abs(TopZ - BottomZ), iDecimals)
                            Exit For
                        End If
                        
                        'Look for Volume to pull thickness
                        If TheMat.IsWorkVolume Then
                            Set TheEle = TheMat.Elements(1)
                            Post.dZ = Round(Abs(TheEle.EndZG), iDecimals)
                            Exit For
                        End If
                        
                    Next TheMat
                    bMatExists = True
                    GoTo done
                    
                End If 'If p.Vars.SHZ > 0 Then
            End If 'If p.Vars.NSH > 0 Then
            
        Else 'If Not p Is Nothing Then
        
            'User Opens the GUI with Nesting, but not posting.
            For Each TheMat In Drw.Geometries
                'This is a sheet, grab it, get DX DY DZ and quit
                If TheMat.Attribute("LicomUKsab_has_bobble") = "1" Then
                    TheMat.GetFeedExtentXYG mix, miy, max, may
                    TopZ = 0
                    BottomZ = -CDbl(TheMat.Attribute("LicomUKsab_sheet_thickness"))
                    Exit For
                End If
            Next
            bMatExists = True
            GoTo done
        
        End If 'If Not p Is Nothing Then
        
    End If 'If gb_HasNesting Then
    
    'Does the Materials LAYER exist?
    Dim lyrMats As Layer
    For Each lyrMats In Drw.Layers
        If lyrMats.Name = "Materials" Then
            bMatLyrExists = True
        End If
    Next
    
    If bMatLyrExists Then
        Set lyrMats = Drw.Layers.Item("Materials")
    Else
        If MsgBox("Error! Material not found! Use Volume?", vbYesNo, "Error Posting") = vbNo Then
            End
        Else
            GoTo ReadVol
        End If
    End If

    For Each TheMat In lyrMats.Geometries
        TheMat.GetFeedExtent minx, miny, maxx, maxy
        
        'If Geo is on a Workplane, check GLOBAL atts
        Set wp = TheMat.GetWorkPlane
        
        For Each TheEle In TheMat.Elements
            minx = TheEle.StartXG
            maxx = TheEle.EndXG
            miny = TheEle.StartYG
            maxy = TheEle.EndYG
            TopZ = TheMat.Attribute("LicomUKDMBGeoZLevelTop")
            BottomZ = TheMat.Attribute("LicomUKDMBGeoZLevelBottom")
            
            If minx <= mix Then mix = minx
            If miny <= miy Then miy = miny
            If BottomZ <= miz Then miz = BottomZ
            
            If maxx >= max Then max = maxx
            If maxy >= may Then may = maxy
            If TopZ >= maz Then maz = TopZ
        Next TheEle
                    
        If maxx >= max Then max = maxx
        If maxy >= may Then may = maxy
        If TopZ >= maz Then maz = TopZ
        
        bMatExists = True
    Next TheMat
    
    'Look Inside Solids
    Dim spMat As SolidPart
    For Each spMat In lyrMats.SolidParts
        
        minx = spMat.minx
        maxx = spMat.maxx
        miny = spMat.miny
        maxy = spMat.maxy
        TopZ = spMat.MaxZ
        BottomZ = spMat.MinZ
        
        If minx <= mix Then mix = minx
        If miny <= miy Then miy = miny
        If BottomZ <= miz Then miz = BottomZ
        
        If maxx >= max Then max = maxx
        If maxy >= may Then may = maxy
        If TopZ >= maz Then maz = TopZ
        
        bMatExists = True
    Next spMat
    
ReadVol:
 
    'If Material has not been found, move to Volume!
    If bMatExists = False Then
        For Each TheMat In Drw.Geometries
            'Sheet not found, Material not found... look for volume in Drw
            If TheMat.IsWorkVolume Then
                Post.dX = VBA.Round(p.Vars.HXW - p.Vars.LXW, iDecimals)
                Post.dY = VBA.Round(p.Vars.HYW - p.Vars.LYW, iDecimals)
                Post.dZ = VBA.Round(p.Vars.HZW - p.Vars.LZW, iDecimals)
                Exit Sub
            End If
        Next TheMat
    End If
    
done:
    If bMatExists Then
        'Once the extent has been found, set it and exit
        Post.dX = Round(max - mix, iDecimals)
        Post.dY = Round(may - miy, iDecimals)
        Post.dZ = Round(Abs(TopZ - BottomZ), iDecimals)
        If Not CheckNegativeMaterial(Post.dX, Post.dY, Post.dZ) Then
            frmMain.txtDX.Text = CStr(Post.dX)
            frmMain.txtDY.Text = CStr(Post.dY)
            frmMain.txtDZ.Text = CStr(Post.dZ)
            frmMain.txtDX.BackColor = vbWhite
            frmMain.txtDY.BackColor = vbWhite
            frmMain.txtDZ.BackColor = vbWhite
        End If
        
    Else
        'Material, Sheet or Volume not found, Let the user know...
        frmMain.txtDX.Text = "Not Set"
        frmMain.txtDY.Text = "Not Set"
        frmMain.txtDZ.Text = "Not Set"
        frmMain.txtDX.BackColor = vbRed
        frmMain.txtDY.BackColor = vbRed
        frmMain.txtDZ.BackColor = vbRed
        Post.dX = 0
        Post.dY = 0
        Post.dZ = 0
        
    End If

Exit Sub

eh:

If Err.Number = 1 Then

Else

End If

MsgBox "Error! Did you forget to add MATERIAL to the Project?", vbCritical
Resume Next

End Sub

'Public Function ReadVolMat(Optional p As PostData) As Boolean
'
'    'AlphaBug... Material looses properties in V8 if Constrained, this line of code fixes this.
'    Dim W As Workpiece
'    Set W = App.ActiveDrawing.Workpiece
'
'    ReadVolMat = False
'
'    Dim Drw As Drawing
'    Set Drw = App.ActiveDrawing
'
'    Dim Ele As Element
'    Dim TheMat As Path
'    Dim TopZ As Double
'    Dim BottomZ As Double
'    Dim minx As Double
'    Dim miny As Double
'    Dim maxx As Double
'    Dim maxy As Double
'    Dim mix As Double
'    Dim miy As Double
'    Dim max As Double
'    Dim may As Double
'    Dim maz As Double
'    Dim miz As Double
'    Dim bMatExists As Boolean
'    Dim bMatLyrExists As Boolean
'    Dim wp As WorkPlane
'
'    Dim AttBotZ As String
'    AttBotZ = "LicomUKDMBGeoZLevelBottom"
'
'    Dim AttTopZ As String
'    AttTopZ = "LicomUKDMBGeoZLevelTop"
'
'    TopZ = 0
'    BottomZ = 0
'    mix = 99999
'    miy = 99999
'    max = -99999
'    may = -99999
'    bMatExists = False
'    bMatLyrExists = False
'
'    With frmMain
'
'    'This is used only when actually generating code
'    'it pulls the right Sheet Size from the screen
'    If Not p Is Nothing Then
'        If p.Vars.NSH > 0 Then
'            Post.dX = Round(p.Vars.SHX, iDecimals)
'            Post.dY = Round(p.Vars.SHY, iDecimals)
'            If p.Vars.SHZ > 0 Then
'                Post.dZ = Round(p.Vars.SHZ, iDecimals)
'            Else
'                For Each TheMat In Drw.Geometries
'                    Post.dZ = 0
'
'                    'Look for Sheets by Attirbute
'                    If TheMat.Attribute("LicomUKsab_sheet_thickness") <> "" Then
'                        .txtDZ.Text = TheMat.Attribute("LicomUKsab_sheet_thickness")
'                        Exit For
'                    End If
'
'                    'Look for Material to reference Thickness
'                    If TheMat.Attribute("LicomUKDMBStockType") = "1" Or _
'                        TheMat.Attribute("LicomUKDMBStockType") = "2" Or _
'                        TheMat.Name = "Material" Then
'
'                        TopZ = TheMat.Attribute(AttTopZ)
'                        BottomZ = TheMat.Attribute(AttBotZ)
'                        Post.dZ = Round(Abs(TopZ - BottomZ), iDecimals)
'                        Exit For
'                    End If
'
'                    'Look for Volume to pull thickness
'                    If TheMat.IsWorkVolume Then
'                        Dim TheEle As Element
'                        Set TheEle = TheMat.Elements(1)
'                        Post.dZ = Round(Abs(TheEle.EndZG), iDecimals)
'                        Exit For
'                    End If
'
'                Next TheMat
'            End If
'
'            CheckNegativeMaterial Post.dX, Post.dY, Post.dZ
'
'            Events.bNestingCalls = Functions.gb_HasNesting
'
'            ReadVolMat = True
'            Exit Function
'
'        End If
'    Else
'        If gb_HasNesting Then
'            Dim sht As NestSheet
'            Dim NI As NestInformation
'            Dim shtpth As Path
'
'            Set NI = Drw.GetNestInformation
'            For Each sht In NI.Sheets
'                Set shtpth = sht.Path
'                shtpth.GetFeedExtent mix, miy, max, may
'                Post.dX = max - mix
'                Post.dY = may - miy
'                Post.dZ = sht.Thickness
'                If Not CheckNegativeMaterial(Post.dX, Post.dY, Post.dZ) Then
'                    frmMain.txtDX.Text = CStr(Post.dX)
'                    frmMain.txtDY.Text = CStr(Post.dY)
'                    frmMain.txtDZ.Text = CStr(Post.dZ)
'                    frmMain.txtDX.BackColor = vbWhite
'                    frmMain.txtDY.BackColor = vbWhite
'                    frmMain.txtDZ.BackColor = vbWhite
'                End If
'                Exit Function
'            Next sht
'        End If
'    End If
'
'
'        'OK Since Nesting Does not Apply, we will look for Volume and for Material(s)
'        For Each TheMat In Drw.Geometries
'            'Populating in case there are Sheets on the Screen
'            If TheMat.Attribute("LicomUKsab_sheet_material") <> "" Then
'                TheMat.GetFeedExtentXYG minx, miny, maxx, maxy
'                Post.dX = Round(maxx - minx, iDecimals)
'                Post.dY = Round(maxy - miny, iDecimals)
'                Post.dZ = TheMat.Attribute("LicomUKsab_sheet_thickness")
'                If Post.dZ = "0" Then
'                    Post.dZ = "Not Set"
'                    .txtDZ.BackColor = vbRed
'                End If
'                 CheckNegativeMaterial Post.dX, Post.dY, Post.dZ
'                ReadVolMat = True
'                Exit Function
'            End If
'        Next TheMat
'
'       'Does the Materials LAYER exist?
'        Dim lyrMats As Layer
'        For Each lyrMats In Drw.Layers
'            If lyrMats.Name = "Materials" Then
'                bMatLyrExists = True
'                Set lyrMats = Drw.Layers.Item("Materials")
'                Exit For
'            End If
'
'            If lyrMats.Name = "Materiales" Then
'                bMatLyrExists = True
'                Set lyrMats = Drw.Layers.Item("Materiales")
'                Exit For
'            End If
'
'        Next
'
'        If Not bMatLyrExists Then
'            GoTo ReadVol
'        End If
'
'        For Each TheMat In lyrMats.Geometries
'
'            For Each Ele In TheMat.Elements
'
'                If Ele.StartXG <= mix Then mix = Ele.StartXG
'                If Ele.StartXG >= max Then max = Ele.StartXG
'                If Ele.StartYG <= miy Then miy = Ele.StartYG
'                If Ele.StartYG >= may Then may = Ele.StartYG
'                If Ele.StartZG <= BottomZ Then BottomZ = Ele.StartZG
'                If Ele.StartZG >= TopZ Then TopZ = Ele.StartZG
'
'                If Ele.EndXG <= mix Then mix = Ele.EndXG
'                If Ele.EndXG >= max Then max = Ele.EndXG
'                If Ele.EndYG <= miy Then miy = Ele.EndYG
'                If Ele.EndYG >= may Then may = Ele.EndYG
'                If Ele.EndZG <= BottomZ Then BottomZ = Ele.EndZG
'                If Ele.EndZG >= TopZ Then TopZ = Ele.EndZG
'
'                'If material geo is in flatland, check the attributes from Dr. B:
'                Set wp = TheMat.GetWorkPlane
'
'                If wp Is Nothing Then
'                    If CDbl(TheMat.Attribute(AttBotZ)) <= BottomZ Then BottomZ = CDbl(TheMat.Attribute(AttBotZ))
'                    If CDbl(TheMat.Attribute(AttTopZ)) >= TopZ Then TopZ = CDbl(TheMat.Attribute(AttTopZ))
'                End If
'
'
'            Next Ele
'
'            bMatExists = True
'        Next TheMat
'
'        'Look Inside Solids
'        Dim spMat As SolidPart
'        For Each spMat In lyrMats.SolidParts
'
'            If spMat.minx <= mix Then mix = spMat.minx
'            If spMat.maxx >= max Then max = spMat.maxx
'            If spMat.miny <= miy Then miy = spMat.miny
'            If spMat.maxy >= may Then may = spMat.maxy
'            If spMat.MinZ <= BottomZ Then BottomZ = spMat.MinZ
'            If spMat.MaxZ >= TopZ Then TopZ = spMat.MaxZ
'
'            bMatExists = True
'        Next spMat
'
'ReadVol:
'
'            'Look for Volume!
'            If bMatExists = False Then
'                For Each TheMat In Drw.Geometries
'                    'Sheet not found, Material not found... look for volume in Drw
'                    If TheMat.IsWorkVolume Then
'                        Post.dX = Round(TheMat.MaxXL - TheMat.MinXL, iDecimals)
'                        Post.dY = Round(TheMat.MaxYL - TheMat.MinYL, iDecimals)
'                        Post.dZ = Round(Abs(p.Vars.HZW - p.Vars.LZW), iDecimals)
'                        CheckNegativeMaterial Post.dX, Post.dY, Post.dZ
'                        bMatExists = True
'                        Exit Function
'                    End If
'                Next TheMat
'            End If
'
'        If bMatExists Then
'            'Once the extent has been found, set it and exit
'            Post.dX = Round(max - mix, iDecimals)
'            Post.dY = Round(may - miy, iDecimals)
'            Post.dZ = Round(Abs(TopZ - BottomZ), iDecimals)
'            If Not CheckNegativeMaterial(Post.dX, Post.dY, Post.dZ) Then
'                frmMain.txtDX.Text = CStr(Post.dX)
'                frmMain.txtDY.Text = CStr(Post.dY)
'                frmMain.txtDZ.Text = CStr(Post.dZ)
'                frmMain.txtDX.BackColor = vbWhite
'                frmMain.txtDY.BackColor = vbWhite
'                frmMain.txtDZ.BackColor = vbWhite
'            End If
'            ReadVolMat = True
'        Else
'            'Material, Sheet or Volume not found, Let the user know...
'            .txtDX.Text = "Not Set"
'            .txtDY.Text = "Not Set"
'            .txtDZ.Text = "Not Set"
'            .txtDX.BackColor = vbRed
'            .txtDY.BackColor = vbRed
'            .txtDZ.BackColor = vbRed
'            Post.dX = 0
'            Post.dY = 0
'            Post.dZ = 0
'            ReadVolMat = False
'        End If
'
'    End With
'
'End Function

Private Function CheckNegativeMaterial(ByVal dDX As Double, ByVal dDY As Double, dDZ As Double) As Boolean
    
    CheckNegativeMaterial = True
    
    If dDX = 0 Then Exit Function
    If dDY = 0 Then Exit Function
    If dDZ = 0 Then Exit Function
    
    If dDX < 0.001 Then
        MsgBox "Material LENGTH is less than 0. Did you not move the MATERIAL to 0,0?", vbCritical
    End If
            
    If dDY < 0.001 Then
        MsgBox "Material WIDTH is less than 0. Did you not move the MATERIAL to 0,0?", vbCritical
    End If
            
    If dDZ < 0.001 Then
        MsgBox "Material THICKNESS is less than 0. Did you put a NEGATIVE Value for Sheet Thickness?", vbCritical
    End If
    
    CheckNegativeMaterial = False
    
End Function

Public Sub GetSettings()
          
    With frmMain
        
        .cmbUnits.Clear
        .cmbUnits.AddItem "MM"
        .cmbUnits.AddItem "IN"
        
        Post.sUnits = GetSetting(DEF_POST_NAME, "Settings", "sUnits", "MM")
        If Post.sUnits = "" Then Post.sUnits = "MM"
        .cmbUnits.Text = sUnits
        
        If sUnits = "MM" Then
            dFeedMult = 1 / 1000
            iDecimals = 5
            dMinArcLen = 0.3
        Else
            dFeedMult = 25.4 / 1000
            iDecimals = 8
            dMinArcLen = 0.3 / 25.4
        End If
        
        .cmbOrigins.Clear
        
        Dim i As Integer
        For i = 1 To 16
            .cmbOrigins.AddItem CStr(i)
        Next
        
        Post.sOrigin = GetSetting(DEF_POST_NAME, "Settings", "sOrigin", "1")
        If Post.sOrigin = "" Then Post.sOrigin = "1"
        .cmbOrigins.Text = Post.sOrigin
        
        Post.sXoff = GetSetting(DEF_POST_NAME, "Settings", "sXoff", "0")
        If Post.sXoff = "" Then Post.sXoff = "0"
        .txtOffX.Text = Post.sXoff
        
        Post.sYoff = GetSetting(DEF_POST_NAME, "Settings", "sYoff", "0")
        If Post.sYoff = "" Then Post.sYoff = "0"
        .txtYoff.Text = Post.sYoff
        
        If sUnits = "MM" Then
            Post.sZoff = GetSetting(DEF_POST_NAME, "Settings", "sZoff", "50")
            If Post.sZoff = "" Then Post.sZoff = "50"
        Else
            Post.sZoff = GetSetting(DEF_POST_NAME, "Settings", "sZoff", "2")
            If Post.sZoff = "" Then Post.sZoff = "2"
        End If
        
        .txtOffZ.Text = Post.sZoff
        
        Post.sZoff = CStr(-VBA.Abs(CDbl(Post.sZoff)))

    End With
    
    With frmConfig
    
        Post.bSupLeadMsgs = CBool(GetSetting(DEF_POST_NAME, "Settings", "bSupLeadMsgs", "0"))
        .chkLeadMsgs.Value = Post.bSupLeadMsgs
        
        Post.bUseCNCFeeds = CBool(GetSetting(DEF_POST_NAME, "Settings", "bUseCNCFeeds", "0"))
        .chkMachineFeeds.Value = Post.bUseCNCFeeds
        
        Post.bHideGUI = CBool(GetSetting(DEF_POST_NAME, "Settings", "bHideGUI", "0"))
        .chkHideGUI.Value = Post.bHideGUI
        
        Post.bIgCustString = CBool(GetSetting(DEF_POST_NAME, "Settings", "bIgCustString", "0"))
        .chkIgnoreCustomUserDataString.Value = Post.bIgCustString
        
        Post.bUseCustomUserSt = CBool(GetSetting(DEF_POST_NAME, "Settings", "chkUseCustomUserString", "0"))
        .chkUseCustomUserString.Value = Post.bUseCustomUserSt
        
        Post.sCustomUserStr = GetSetting(DEF_POST_NAME, "Settings", "txtUseCustomUserString", "")
        
        If Post.bIgCustString Then
            .chkUseCustomUserString.Enabled = False
            .txtUseCustomUserString.Enabled = False
        Else
            .chkUseCustomUserString.Enabled = True
            .txtUseCustomUserString.Enabled = True
        End If
        
        If Post.sCustomUserStr = "" Or Not Post.bUseCustomUserSt Then
            .chkUseCustomUserString.Value = False
            .txtUseCustomUserString.Enabled = False
        End If
        
        Post.sXCUT = GetSetting(DEF_POST_NAME, "Settings", "txtXCUT", "")
        .txtXCUT.Text = Post.sXCUT
        Post.sYCUT = GetSetting(DEF_POST_NAME, "Settings", "txtYCUT", "")
        .txtYCUT.Text = Post.sYCUT
        
        Post.bFlagAutoLabel = CBool(GetSetting(DEF_POST_NAME, "Settings", "chkFlagForAutoLabel", "0"))
        .chkAutoLabels.Value = Post.bFlagAutoLabel
        
        Post.sImageType = GetSetting(DEF_POST_NAME, "Settings", "cmbImageType", "BMP: Bit Map File")
        '.cmbImageType.Text = Post.sImageType
        
        Post.sOutlineNote = GetSetting(DEF_POST_NAME, "Settings", "txtOutlineNote", "")
        .txtOutlineNote.Text = Post.sOutlineNote
        
    End With

End Sub

Public Sub SaveSettings()

    With frmMain
                
        If Post.sUnits = "" Then Post.sUnits = "IN"
        SaveSetting DEF_POST_NAME, "Settings", "sUnits", Post.sUnits
        sUnits = .cmbUnits.Text
                        
        If Post.sOrigin = "" Then Post.sOrigin = "1"
        SaveSetting DEF_POST_NAME, "Settings", "sOrigin", Post.sOrigin
        Post.sOrigin = .cmbOrigins.Text
        
        
        If Post.sXoff = "" Then Post.sXoff = "0"
        SaveSetting DEF_POST_NAME, "Settings", "sXoff", Post.sXoff
        Post.sXoff = .txtOffX.Text
        
        
        If Post.sYoff = "" Then Post.sYoff = "0"
        SaveSetting DEF_POST_NAME, "Settings", "sYoff", Post.sYoff
        Post.sYoff = .txtYoff.Text
            
        
        If sUnits = "MM" Then
            If Post.sZoff = "" Then Post.sZoff = "50"
        Else
            If Post.sZoff = "" Then Post.sZoff = "1"
        End If
        SaveSetting DEF_POST_NAME, "Settings", "sZoff", Post.sZoff
        Post.sZoff = .txtOffZ.Text
        
    End With
    
    If Post.bSupLeadMsgs Then
        SaveSetting DEF_POST_NAME, "Settings", "bSupLeadMsgs", "1"
    Else
        SaveSetting DEF_POST_NAME, "Settings", "bSupLeadMsgs", "0"
    End If
    
    
    If Post.bUseCNCFeeds Then
        SaveSetting DEF_POST_NAME, "Settings", "bUseCNCFeeds", "1"
    Else
        SaveSetting DEF_POST_NAME, "Settings", "bUseCNCFeeds", "0"
    End If
    
    If Post.bHideGUI Then
        SaveSetting DEF_POST_NAME, "Settings", "bHideGUI", "1"
    Else
        SaveSetting DEF_POST_NAME, "Settings", "bHideGUI", "0"
    End If
    
    If Post.bIgCustString Then
        SaveSetting DEF_POST_NAME, "Settings", "bIgCustString", "1"
    Else
        SaveSetting DEF_POST_NAME, "Settings", "bIgCustString", "0"
    End If
    
    If Post.bUseCustomUserSt Then
        SaveSetting DEF_POST_NAME, "Settings", "chkUseCustomUserString", "1"
    Else
        SaveSetting DEF_POST_NAME, "Settings", "chkUseCustomUserString", "0"
    End If
    
    If Post.bFlagAutoLabel Then
        SaveSetting DEF_POST_NAME, "Settings", "chkFlagForAutoLabel", "1"
    Else
        SaveSetting DEF_POST_NAME, "Settings", "chkFlagForAutoLabel", "0"
    End If
    
    If Post.sImageType <> "" Then
        SaveSetting DEF_POST_NAME, "Settings", "cmbImageType", Post.sImageType
    Else
        SaveSetting DEF_POST_NAME, "Settings", "cmbImageType", "BMP: Bit Map File"
    End If
    
    If Post.sOutlineNote <> "" Then
        SaveSetting DEF_POST_NAME, "Settings", "txtOutlineNote", Post.sOutlineNote
    Else
        SaveSetting DEF_POST_NAME, "Settings", "txtOutlineNote", ""
    End If
    
    If Post.sCustomUserStr <> "" Then
        SaveSetting DEF_POST_NAME, "Settings", "txtUseCustomUserString", Post.sCustomUserStr
    Else
        SaveSetting DEF_POST_NAME, "Settings", "txtUseCustomUserString", ""
    End If
    
    Call SetPins
    
End Sub

Public Function RetMachineType(ByRef p As PostData) As String

    RetMachineType = ""
    
    Dim s3DAction   As String
    Dim s3DAxisType As String
    Dim s3DMethod   As String
    Dim s3DProject  As String
    
    s3DAction = p.Path.Attribute("LicomUKDMB3DAction")
    If s3DAction = "" Then s3DAction = "0"
    Select Case CInt(s3DAction)
        Case 1
            s3DAction = "Machine Surfaces"
        Case 2
            s3DAction = "Machine Surfaces with Tool Side"
        Case 3
            s3DAction = "Z-Contour Roughing"
        Case 4
            s3DAction = "Along Intersection"
        Case 5
            s3DAction = "Along Spline or Polyline"
        Case 6
            s3DAction = "Between two Geometries"
        Case 7
            s3DAction = "Edit Tool Angle Command"
        Case 8
            s3DAction = "Manual Toolpath through API"
    End Select
    
    s3DAxisType = p.Path.Attribute("LicomUKDMB3DAxisType")
    If s3DAxisType = "" Then s3DAxisType = "0"
    Select Case CInt(s3DAxisType)
        Case 1
            s3DAxisType = "3-Axis"
        Case 2
            s3DAxisType = "4-Axis (XZ Rot)"
        Case 3
            s3DAxisType = "4-Axis (YZ Rot)"
        Case 4
            s3DAxisType = "5-Axis"
        Case 5
            s3DAxisType = "System Reserved! How did you get this?!"
        Case 6
            s3DAxisType = "4-Axis (XY Rot)"
    End Select
    
    s3DMethod = p.Path.Attribute("LicomUKDMB3DMethod")
    If s3DMethod = "" Then s3DMethod = "0"
    Select Case CInt(s3DMethod)
        Case 1
            s3DMethod = "Parameter Lines"
        Case 2
            s3DMethod = "Horizontal Z Contours"
        Case 3
            s3DMethod = "Along Line in XY Plane"
        Case 4
            s3DMethod = "Projected Contours"
        Case 5
            s3DMethod = "Radial"
        Case 6
            s3DMethod = "Spiral"
        Case 7
            s3DMethod = "Rest Machining"
        Case 8
            s3DMethod = "Drive Curves"
        Case 9
            s3DMethod = "Parallel - Shallow Slopes"
        Case 10
            s3DMethod = "Z Contours - Steep Slopes"
        Case 11
            s3DMethod = "Flat Area Offset"
        Case 12
            s3DMethod = "Parallel - Steep Slopes"
        Case 13
            s3DMethod = "Reserved"
        Case 14
            s3DMethod = "Disk Finish"
        Case 15
            s3DMethod = "Disk Rough"
        Case 16
            s3DMethod = "Disk Sidecut"
        Case 17
            s3DMethod = "Helical Z"
        Case 18
            s3DMethod = "Z Contours + Flat Area Offset"
        Case 19
            s3DMethod = "Cylindrical Parallel"
        Case 20
            s3DMethod = "Constant Cusp"
        Case 21
            s3DMethod = "Contour Roughing"
        Case 22
            s3DMethod = "Linear Roughing"
        Case 23
            s3DMethod = "Spiral or Waveform Roughing"
        Case 24
            s3DMethod = "Cylindrical Profiling"
        Case 25
            s3DMethod = "Z (Enhanced Undercuts)"
    End Select

    s3DProject = p.Path.Attribute("LicomUKDMB3DProject")
    If s3DProject = "" Then s3DProject = "0"
    Select Case CInt(s3DProject)
        Case 1
            s3DProject = "Global 3-axis"
        Case 2
            s3DProject = "Global 4-axis (XZ rot)"
        Case 3
            s3DProject = "Global 4-axis (YZ rot)"
        Case 4
            s3DProject = "Global 5-axis"
        Case 5
            s3DProject = "Local 3-axis (Helical arcs)"
        Case 6
            s3DProject = "System Reserved! How did you get this?!"
        Case 7
            s3DProject = "Local 3-axis"
        Case 8
            s3DProject = "System Reserved! How did you get this?!"
        Case 9
            s3DProject = "As Original"
    End Select
    
    If s3DAction <> "0" Then
        RetMachineType = s3DAction
        If s3DAxisType <> "0" Then RetMachineType = RetMachineType & ", " & s3DAxisType
        If s3DMethod <> "0" Then RetMachineType = RetMachineType & ", " & s3DMethod
        If s3DProject <> "0" Then RetMachineType = RetMachineType & ", " & s3DProject
        RetMachineType = RetMachineType & "!"
    Else
        Select Case p.Vars.MOT
            Case 1
                RetMachineType = "Rough or Finish"
            Case 2
                RetMachineType = "Contour Pocket"
            Case 3
                RetMachineType = "Manual Toolpath"
            Case 4
                RetMachineType = "Spiral Pocket"
            Case 5
                RetMachineType = "Linear Pocket"
            Case 10
                RetMachineType = "3D Engrave"
            Case 11
                RetMachineType = "Surface Machining"
            Case 12
                RetMachineType = "Machine 3D Polyline"
            Case 13
                RetMachineType = "Cut between 2 Geos"
            Case 21
                RetMachineType = "Drill"
            Case 22
                RetMachineType = "Peck"
            Case 23
                RetMachineType = "Tap"
            Case 24
                RetMachineType = "Bore"
            Case Else
                RetMachineType = "Type Not Found"
        End Select
    End If
    
    If RetMachineType = "" Then RetMachineType = "Machining Type Not Found!"

End Function

Public Function MachiningAllowed(ByRef p As PostData) As Boolean

    Dim bRet As Boolean
    bRet = False
    
    Dim s3DAction   As String
    Dim s3DAxisType As String
    Dim s3DMethod   As String
    Dim s3DProject  As String
    
    s3DAction = p.Path.Attribute("LicomUKDMB3DAction")
    If s3DAction = "" Then s3DAction = "0"
    Select Case CInt(s3DAction)
        Case 1
            s3DAction = "Machine Surfaces"
            bRet = True
        Case 2
            s3DAction = "Machine Surfaces with Tool Side"
            bRet = False
        Case 3
            s3DAction = "Z-Contour Roughing"
            bRet = True
        Case 4
            s3DAction = "Along Intersection"
            bRet = False
        Case 5
            s3DAction = "Along Spline or Polyline"
            bRet = False
        Case 6
            s3DAction = "Between two Geometries"
            bRet = False
        Case 7
            s3DAction = "Edit Tool Angle Command"
            bRet = False
        Case 8
            s3DAction = "Manual Toolpath through API"
            bRet = True
    End Select
        
    s3DAxisType = p.Path.Attribute("LicomUKDMB3DAxisType")
    If s3DAxisType = "" Then s3DAxisType = "0"
    Select Case CInt(s3DAxisType)
        Case 1
            s3DAxisType = "3-Axis"
            bRet = True
        Case 2
            s3DAxisType = "4-Axis (XZ Rot)"
            bRet = False
        Case 3
            s3DAxisType = "4-Axis (YZ Rot)"
            bRet = False
        Case 4
            s3DAxisType = "5-Axis"
            bRet = False
        Case 5
            s3DAxisType = "System Reserved! How did you get this?!"
            bRet = False
        Case 6
            s3DAxisType = "4-Axis (XY Rot)"
            bRet = False
    End Select
        
    s3DProject = p.Path.Attribute("LicomUKDMB3DProject")
    If s3DProject = "" Then s3DProject = "0"
    Select Case CInt(s3DProject)
        Case 1
            s3DProject = "Global 3-axis"
            bRet = True
        Case 2
            s3DProject = "Global 4-axis (XZ rot)"
             bRet = False
        Case 3
            s3DProject = "Global 4-axis (YZ rot)"
             bRet = False
        Case 4
            s3DProject = "Global 5-axis"
             bRet = False
        Case 5
            s3DProject = "Local 3-axis (Helical arcs)"
             bRet = True
        Case 6
            s3DProject = "System Reserved! How did you get this?!"
             bRet = False
        Case 7
            s3DProject = "Local 3-axis"
             bRet = True
        Case 8
            s3DProject = "System Reserved! How did you get this?!"
             bRet = False
        Case 9
            s3DProject = "As Original"
             bRet = True
    End Select
    
    If bRet Then
        MachiningAllowed = bRet
        Exit Function
    Else
        Select Case p.Vars.MOT
            Case 1
                'RetMachineType = "Rough or Finish!"
                 bRet = True
'                 If p.Path.GetMillData.HelicalInterpolation And _
'                    p.Path.GetMillData.McComp <> acamCompTOOLCEN Then
'                    MsgBox "Error! If using Helical interpolation you MUST use Tool on Center!", vbCritical
'                    End
'                End If
            Case 2
                'RetMachineType = "Contour Pocket!"
                 bRet = True
            Case 3
                'RetMachineType = "Manual Toolpath!"
                bRet = True
            Case 4
                'RetMachineType = "Spiral Pocket!"
                bRet = True
            Case 5
                'RetMachineType = "Linear Pocket!"
                bRet = True
            Case 10
                'RetMachineType = "3D Engrave!"
                bRet = True
            Case 11
                'RetMachineType = "Surface Machining!"
                bRet = True
            Case 12
                'RetMachineType = "Machine 3D Polyline!"
                bRet = False
            Case 13
                'RetMachineType = "Cut between 2 Geos!"
                bRet = False
            Case 21
                'RetMachineType = "Drill!"
                bRet = True
            Case 22
                'RetMachineType = "Peck!"
                bRet = True
            Case 23
                'RetMachineType = "Tap!"
                bRet = True
            Case 24
                'RetMachineType = "Bore!"
                bRet = True
            Case Else
                'RetMachineType = "Type Not Found!"
                bRet = True
        End Select
    End If
    
    MachiningAllowed = bRet

End Function

Public Sub SetPins()
    
    Dim sAddPins As String
    sAddPins = ""
    
    With frmMain
        'Row1
        If .chkRow1.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
        'Row2
        If .chkRow2.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
        'Row3
        If .chkRow3.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
        'Row4
        If .chkRow4.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
        'Row5
        If .chkRow5.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
        'Row6
        If .chkRow6.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
        'Row7
        If .chkRow7.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
        'Row8
        If .chkRow8.Value Then
            sAddPins = sAddPins & "1"
        Else
            sAddPins = sAddPins & "0"
        End If
    End With
    
    'Store string
    SaveSetting DEF_POST_NAME, "Settings", "sPins", sAddPins
    
End Sub

Public Sub GetPins()

    Post.sPins = GetSetting(DEF_POST_NAME, "Settings", "sPins", "10101010")
    
    Dim sExtract As String
    sExtract = "0"
    
    Dim sTemp As String
    sTemp = ""
    
    'First Pin
    sExtract = VBA.Left(Post.sPins, 1)
    sTemp = VBA.Right(Post.sPins, 7)
    If sExtract = "1" Then
        frmMain.chkRow1.Value = True
    Else
        frmMain.chkRow1.Value = False
    End If
    
    'Second Pin
    sExtract = VBA.Left(sTemp, 1)
    sTemp = VBA.Right(sTemp, 6)
    If sExtract = "1" Then
        frmMain.chkRow2.Value = True
    Else
        frmMain.chkRow2.Value = False
    End If
    
    'Third Pin
    sExtract = VBA.Left(sTemp, 1)
    sTemp = VBA.Right(sTemp, 5)
    If sExtract = "1" Then
        frmMain.chkRow3.Value = True
    Else
        frmMain.chkRow3.Value = False
    End If

    'Fourth Pin
    sExtract = VBA.Left(sTemp, 1)
    sTemp = VBA.Right(sTemp, 4)
    If sExtract = "1" Then
        frmMain.chkRow4.Value = True
    Else
        frmMain.chkRow4.Value = False
    End If
    
    'Fifth Pin
    sExtract = VBA.Left(sTemp, 1)
    sTemp = VBA.Right(sTemp, 3)
    If sExtract = "1" Then
        frmMain.chkRow5.Value = True
    Else
        frmMain.chkRow5.Value = False
    End If
    
    'Sixth Pin
    sExtract = VBA.Left(sTemp, 1)
    sTemp = VBA.Right(sTemp, 2)
    If sExtract = "1" Then
        frmMain.chkRow6.Value = True
    Else
        frmMain.chkRow6.Value = False
    End If
    
    'Seventh Pin
    sExtract = VBA.Left(sTemp, 1)
    sTemp = VBA.Right(sTemp, 1)
    If sExtract = "1" Then
        frmMain.chkRow7.Value = True
    Else
        frmMain.chkRow7.Value = False
    End If
    
    'Eighth Pin
    sExtract = VBA.Left(sTemp, 1)
    sTemp = VBA.Right(sTemp, 1)
    If sExtract = "1" Then
        frmMain.chkRow8.Value = True
    Else
        frmMain.chkRow8.Value = False
    End If
End Sub

Public Function GetUserDataString() As String
    Dim sRet As String
    sRet = ""
    sRet = "CUSTSTR=" & Chr(34) & "0,0,0,1,1,0,0,0,0,"
    
    If Post.bUseCustomUserSt Then
        GetUserDataString = "CUSTSTR=" & GetSetting(DEF_POST_NAME, "Settings", "txtUseCustomUserString", "")
        Exit Function
    End If
    
    Dim bPin1 As Boolean
    Dim bPin2 As Boolean
    Dim bPin3 As Boolean
    Dim bPin4 As Boolean
    Dim bPin5 As Boolean
    Dim bPin6 As Boolean
    Dim bPin7 As Boolean
    Dim bPin8 As Boolean
    
    bPin1 = False
    bPin2 = False
    bPin3 = False
    bPin4 = False
    bPin5 = False
    bPin6 = False
    bPin7 = False
    bPin8 = False

    Call GetPins

    bPin1 = frmMain.chkRow1.Value
    bPin2 = frmMain.chkRow2.Value
    bPin3 = frmMain.chkRow3.Value
    bPin4 = frmMain.chkRow4.Value
    bPin5 = frmMain.chkRow5.Value
    bPin6 = frmMain.chkRow6.Value
    bPin7 = frmMain.chkRow7.Value
    bPin8 = frmMain.chkRow8.Value
    
    'It SEEMS that, as long as the 1st pin is used, the string follows immediately
    If bPin1 Then
        sRet = sRet & "1"
        If bPin2 Then sRet = sRet & "E2"
        If bPin3 Then sRet = sRet & "E3"
        If bPin4 Then sRet = sRet & "E4"
        If bPin5 Then sRet = sRet & "E5"
        If bPin6 Then sRet = sRet & "E6"
        If bPin7 Then sRet = sRet & "E7"
        If bPin8 Then sRet = sRet & "E8"
    Else
        'OK if Pin 1 is not selected, it SEEMS we have to check if this is EVEN or ODD
        'Check Odds
        If (bPin3 Or bPin5 Or bPin7) And Not (bPin2 Or bPin4 Or bPin6 Or bPin8) Then
            'OK, this is ODD numbers, skip one column!
            sRet = sRet & ","
            If bPin3 Then sRet = sRet & "3"
            If bPin5 Then sRet = sRet & "E5"
            If bPin7 Then sRet = sRet & "E7"
        Else
            'This seems to be EVEN numbers
            If bPin2 Then sRet = sRet & "2"
            If bPin4 Then sRet = sRet & "E4"
            If bPin6 Then sRet = sRet & "E6"
            If bPin8 Then sRet = sRet & "E8"
        End If
    End If
    
    'Finish contructing it
    sRet = sRet & ",0,0,0,0,0,0,0,0,0" & Chr(34)
    
    If Post.bIgCustString Then
        GetUserDataString = "CUSTSTR=" & Chr(34) & Chr(34)
    Else
        GetUserDataString = sRet
    End If
    
End Function

' Inverse Sine
 Function ArcSin(X As Double) As Double
    
    ArcSin = Atn(X / Math.Sqr(-X * X + 1))
    ArcSin = ToDegrees(ArcSin)
    
 End Function

 ' Inverse Cosine
 Function ArcCos(X As Double) As Double
     
    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    ArcCos = ToDegrees(ArcCos)
    
 End Function


'-- pi = 4 * Atn(1)
'-- Atn(1)= (pi/4) radian = 45 degrees
'-- 1 degree = pi/180 radian = (pi/4)/45 radian = Atn(1)/45 radian

Function ToRadians(AngInDegrees As Double)

    Dim pi As Double
    pi = 4 * Atn(1)
    
    ToRadians = AngInDegrees * (pi / 180)

End Function

Function ToDegrees(AngInRadians As Double)

    Dim pi As Double
    pi = 4 * Atn(1)

    ToDegrees = AngInRadians * (180 / pi)

End Function

Public Sub AddAutoLabelInfo(ByVal sTheFileName As String)

    'Only for nested objects
    If gb_HasNesting = False Then
        Exit Sub
    End If

    'System
    Dim FSO
    Dim Folder
    Dim TheFile
    Dim TempFile
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TriStateUseDefault = -2, TriStateTrue = -1, TriStateFalse = 0
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    'ACAM
    Dim drwActive       As Drawing
    Dim ptp             As Path
    Dim NI              As NestInformation
    Dim NS              As NestSheet
    Dim NPI             As NestPartInstance
    
    'Strings
    Dim sTemp               As String
    Dim sMNOxml             As String
    Dim sMNOxmlTitle        As String
    Dim sTempFile           As String
    Dim sBuffer             As String
    Dim sLblName            As String
    Dim sLine2              As String
    Dim sLine3              As String
    Dim sLine4              As String
    
    'Numbers
    Dim iNext           As Integer
    Dim iShtCnt         As Integer
    
    'XML
    Dim doc2            As DOMDocument60
    Dim xmlComment      As IXMLDOMComment
    Dim xmlRoot         As IXMLDOMNode
    Dim xmlPart         As IXMLDOMElement
    
    Dim bNoPNI          As Boolean
        
    Set drwActive = App.ActiveDrawing
    Set doc2 = New DOMDocument60
    
    Set NI = drwActive.GetNestInformation

    Dim sFileExt As String
    sFileExt = VBA.UCase(VBA.Right(sTheFileName, 4))

    Dim sFileTitle As String
    sFileTitle = VBA.Right(sTheFileName, VBA.Len(sTheFileName) - VBA.InStrRev(sTheFileName, "\"))
    sFileTitle = VBA.Left(sFileTitle, VBA.Len(sFileTitle) - 4)
    
    Select Case sFileExt
        Case ".CIX"
            sMNOxml = VBA.Left(sTheFileName, VBA.Len(sTheFileName) - 4) & "_lbl.xml"
            sTempFile = VBA.Left(sTheFileName, VBA.Len(sTheFileName) - 4) & ".autolabel"
        Case ".ANC"
            sMNOxml = VBA.Left(sTheFileName, VBA.Len(sTheFileName) - 4) & "_lbl.xml"
            sTempFile = VBA.Left(sTheFileName, VBA.Len(sTheFileName) - 4) & ".autolabel"
        Case Else
            sMNOxml = sTheFileName & "_lbl.xml"
            sTempFile = sTheFileName & ".autolabel"
    End Select
    
    sMNOxmlTitle = VBA.Right(sMNOxml, VBA.Len(sMNOxml) - VBA.InStrRev(sMNOxml, "\"))
    sMNOxmlTitle = VBA.Left(sMNOxmlTitle, VBA.InStr(sMNOxmlTitle, "_lbl.xml") - 1)
                               
    Set xmlComment = doc2.createComment("Created by Hector Henry for Alphacam US. This file CANNOT be loaded to BiesseWorks/BiesseNest!")
    doc2.appendChild xmlComment

    Set xmlRoot = doc2.createElement("CutList")
    doc2.appendChild xmlRoot
            
    iShtCnt = 0

    bNoPNI = True

    For Each NS In NI.Sheets
        
        iShtCnt = iShtCnt + 1
                    
        sLblName = sTempFile & CStr(iShtCnt)
        
        Set TheFile = FSO.CreateTextFile(sLblName, True)
            
        For Each NPI In NS.Parts
        
            bNoPNI = False
            
            'This small algorithm gets the LAST PATH in the NPI

            Dim pNext As Path
            Dim pLastPath As Path
            Dim iNPIcnt As Integer
            
            iNPIcnt = 0
                        
            iNPIcnt = NPI.Paths.Count
            
            Set pLastPath = NPI.Paths(iNPIcnt)
            
            'Check to see if we are using TPs or Geos
            If Post.sOutlineNote <> "" Then
                'Means we are using toolpaths with note
                While pLastPath.Attribute("LicomUKDMBOperationNote01") <> Post.sOutlineNote
                    If iNPIcnt = 0 Then
                        MsgBox "Error! Outline note not found on perimteer path!"
                    End If
                    Set pLastPath = NPI.Paths(iNPIcnt - 1)
                    iNPIcnt = iNPIcnt - 1
                Wend
            Else
                
'                'Means we ARE NOT using toolpaths. Using Geos.
'                Dim dMatBot As Double
'                dMatBot = CDbl(pLastPath.Attribute("LicomUKDMBGeoZLevelBottom"))
'
'                While pLastPath.IsToolPath Or (dMatBot - Post.dZ) > 0.001 Or (pLastPath.ToolInOut <> acamOUTSIDE)
'                    Set pLastPath = NPI.Paths(iNPIcnt - 1)
'                    iNPIcnt = iNPIcnt - 1
'                    dMatBot = CDbl(pLastPath.Attribute("LicomUKDMBGeoZLevelBottom"))
'                Wend
                
                'Deprecated while an update can be made to add the 3 methods in BP: Outline Note, Toolpaths and Geos
                While (IsOuterThruPath(pLastPath, Post.dZ) = False) Or (pLastPath.IsToolPath = False) Or pLastPath.IsPathAllRapids
                    If iNPIcnt = 0 Then
                        MsgBox "Error! All Toolpaths have been checked and no Perimeter Path has been found!"
                        End
                    End If
                    Set pLastPath = NPI.Paths(iNPIcnt - 1)
                    iNPIcnt = iNPIcnt - 1
                Wend
                
            End If
            
            Set ptp = pLastPath
            
            If VBA.Trim(ptp.Attribute("LicomUKDMB_path_uid_new")) <> "" Then
            
                sTemp = ".\" & sFileTitle & "\" & VBA.Trim(ptp.Attribute("LicomUKDMB_path_uid_new")) & "." & VBA.LCase(VBA.Left(sImageType, 3))
                
                'This will add the Part node to the _lbl.xml file, since we have the ID number
                Set xmlPart = doc2.createElement("Part")
                
                'Add attributes
                xmlPart.SetAttribute "Draw_1", sTemp
                xmlPart.SetAttribute "id", ("P" & VBA.Trim(ptp.Attribute("LicomUKDMB_path_uid_new")))
                
                'Add to root
                xmlRoot.appendChild xmlPart
                
                'Add the Macro to the CIX file
                FindMidPoints NS, NPI
                TheFile.WriteLine "BEGIN MACRO"
                TheFile.WriteLine "   NAME=LABEL"
                TheFile.WriteLine "   PARAM,NAME=ID,VALUE=" & Chr(34) & "P" & VBA.Trim(ptp.Attribute("LicomUKDMB_path_uid_new")) & Chr(34)
                TheFile.WriteLine "   PARAM,NAME=X,VALUE=" & VBA.Trim(VBA.CStr(Round(dMidX, 3)))
                TheFile.WriteLine "   PARAM,NAME=Y,VALUE=" & VBA.Trim(VBA.CStr(Round(dMidY, 3)))
                TheFile.WriteLine "   PARAM,NAME=NAME,VALUE=" & Chr(34) & "LBL" & Chr(34)
                TheFile.WriteLine "   PARAM,NAME=DATA,VALUE=" & Chr(34) & sFileTitle & ";1;" & VBA.Trim(ptp.Attribute("LicomUKDMB_path_uid_new")) & ";1" & Chr(34)
                TheFile.WriteLine "   PARAM,NAME=Z,VALUE=0"
                TheFile.WriteLine "   PARAM,NAME=ROT,VALUE=0"
                TheFile.WriteLine "   PARAM,NAME=OPT,VALUE=NO"
                TheFile.WriteLine "   PARAM,NAME=ISO,VALUE=" & Chr(34) & Chr(34)
                TheFile.WriteLine "   PARAM,NAME=LAY,VALUE=" & Chr(34) & Chr(34)
                TheFile.WriteLine "END MACRO"
                TheFile.WriteLine ""
                                                  
            End If
        Next NPI
        
        TheFile.Close
        
    Next NS
               
    doc2.Save sMNOxml
   
    'XML
    'Set doc2 = Nothing
    Set xmlPart = Nothing
    Set xmlRoot = Nothing
    Set xmlComment = Nothing

    If bNoPNI Then
        MsgBox "Error Appending AutoLabel Info! Nest Part Instance Data does not exist!"
        End
    End If

    Exit Sub
    
eh:
    MsgBox Err.Description & " - AddAutoLabelInfo"


End Sub

Private Function FindMidPoints(NS As NestSheet, NPI As NestPartInstance) As Double

    dMidX = 0
    dMidY = 0
    
    Dim SheetMinx As Double
    Dim SheetMiny As Double
    Dim SheetMaxx As Double
    Dim SheetMaxy As Double
    Dim dSheetWid As Double
    
    Dim PartMinx As Double
    Dim PartMiny As Double
    Dim PartMaxx As Double
    Dim PartMaxy As Double
    
    Dim mix As Double
    Dim miy As Double
    Dim max As Double
    Dim may As Double
    
    mix = 999999
    max = -999999
    miy = 999999
    may = -999999
    
    Dim p As Path
    
    Dim sp As Path
    Set sp = NS.Path
    
    sp.GetFeedExtentXYG SheetMinx, SheetMiny, SheetMaxx, SheetMaxy
    
    dSheetWid = SheetMaxy - SheetMiny
    
    'find the part boundries
    For Each p In NPI.Paths
    
        p.GetFeedExtentXYG PartMinx, PartMiny, PartMaxx, PartMaxy
        
        If PartMinx <= mix Then mix = PartMinx
        If PartMiny <= miy Then miy = PartMiny
        If PartMaxx >= max Then max = PartMaxx
        If PartMaxy >= may Then may = PartMaxy
    
    Next p
    
    'Get part center
    dMidX = (mix + max) / 2 - SheetMinx
    dMidY = dSheetWid - ((miy + may) / 2 - SheetMiny)
    
End Function

Public Function IsElementLast(ByRef TheElement As Element, ByRef pTestPath As Path) As Boolean
    
    Dim bRet As Boolean
    bRet = False

    Dim LastEle As Element
    Set LastEle = pTestPath.Elements.Item(pTestPath.Elements.Count)
        
    'Check to make sure that the LastEle is not a rapid, if it is, keep going back until you have a feed path
    'Also, make sure that the element is NOT Zero length. this is a known issue
    If LastEle.IsRapid Then
        While (LastEle.IsRapid Or LastEle.Length < 0.0001)
            Set LastEle = LastEle.GetPrevious
        Wend
    End If

    'Check if this is the last element of the path
    If LastEle.IsSame(TheElement) Then
        bRet = True
    End If

    IsElementLast = bRet

End Function

Private Function IsOuter(ByRef pTestPath As Path) As Boolean

    IsOuter = True
    
    Dim sSheetID As String
    Dim geo As Path

    'For the terms of defining an "outer path" we will assume the path is INSIDE the SHEET (i.e. has the attribute "LicomUKsab_sheet_ident"
    'I just need to test to make sure THIS pTestPath is NOT INSIDE another path

    'Quit if it not a closed path
    If Not pTestPath.ClosedEx Then
        IsOuter = False
        Exit Function
    End If
        
    'Sheet path. Cannot be considered an "Outer Cutting Path/Geo"
    If VBA.Len(pTestPath.Attribute("LicomUKsab_has_bobble")) > 0 Or VBA.Len(pTestPath.Attribute("LicomUKsab_is_bobble")) > 0 Then
        IsOuter = False
        Exit Function
    End If
        
    'Any Geo Path NOT on a sheet. Cannot be considered an "Outer Cutting Path/Geo"
    If VBA.Len(pTestPath.Attribute("LicomUKsab_sheet_ident")) = 0 Then
        IsOuter = False
        Exit Function
    End If
    
'    'The path being tested now is already a thru path. This upgrade uses the tool note to expedite testing.
'    If bUseNote And pTestPath.Attribute("LicomUKDMBOperationNote01") = sToolNote Then
'        Exit Function
'    End If
    
    If pTestPath.IsToolPath Then
        'Quit if drilling. No need to test.
        If pTestPath.GetMillData.IsDrilling Then
            IsOuter = False
            Exit Function
        End If
                
        If pTestPath.Attribute("LicomUKsab_outer_path") = "1" Then
            Exit Function
        End If
        
        'Test until you find an enclosing geo
        For Each geo In App.ActiveDrawing.ToolPaths
            If geo.Closed Then
                If pTestPath.TestInsidePath(geo) And (Not pTestPath.IsSame(geo)) And (geo.Name <> "Material") Then
                    'Only ONE special case: The sheet geo may rightfully contain this geo/path
                    If VBA.Len(geo.Attribute("LicomUKsab_has_bobble")) > 0 And (pTestPath.Attribute("LicomUKsab_sheet_ident") = geo.Attribute("LicomUKsab_sheet_ident")) Then
                        'Nothing
                    Else
                        'Found containing geo
                        IsOuter = False
                        Exit Function
                    End If
                End If
            End If
        Next geo
    Else
    
        'Quit if it not a closed path
        If Not pTestPath.Closed Then
            IsOuter = False
            Exit Function
        End If
        
        'Test until you find an enclosing geo
        For Each geo In App.ActiveDrawing.Geometries
            If geo.Closed Then
                If pTestPath.TestInsidePath(geo) And pTestPath.IsSame(geo) = False Then
                    'Only ONE special case: The sheet geo may rightfully contain this geo/path
                    If VBA.Len(geo.Attribute("LicomUKsab_has_bobble")) > 0 And (pTestPath.Attribute("LicomUKsab_sheet_ident") = geo.Attribute("LicomUKsab_sheet_ident")) Then
                        'Nothing
                    Else
                        'Found containing geo
                        IsOuter = False
                        Exit Function
                    End If
                End If
            End If
        Next geo
    End If

End Function

Public Function IsOuterThruPath(ByRef ThePath As Path, ByVal dThick As Double) As Boolean

    
    Dim md As MillData
    Dim dFinalDepth As Double
    
    IsOuterThruPath = False
    
    If Not ThePath.IsToolPath Then Exit Function
    
    If dThick <= 0 Then
        MsgBox "Error! Sheet Thickness is ZERO!!!"
        End
    End If
    
    If IsOuter(ThePath) Then
        Set md = ThePath.GetMillData
        dFinalDepth = md.FinalDepth
        If dFinalDepth <= -dThick Then
            IsOuterThruPath = True
        End If
    End If
    
End Function
