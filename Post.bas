Attribute VB_Name = "Post"
Option Explicit

   'Customer       :
   'Machine        : All. Not machine specific. bSolid release.
   'Control        : None. This is a CIX file which may be read onto bSolid or BiesseWorks.
   'Attachments    : Saw cutting Support. All workplane Ops.
   'Units          : Metric/Inch (Drawing) Prompt
   'No of Axis     : XYZBC
   'Coord System   : IN/MM
   'Z -Axis        : Vertical
   'Contact        : Justin Maynard
   'Email          : Justin.maynard@alphacam.com
   'Dealer Name    : Vero Software US
   'Base Author    : Hector Henry
   
   '==== NOTES ===================================================
   ' Program Material top MUST be ZERO.
   ' IF COMP IS USED ...
   ' TURN ON "APPLY COMP ON RAPID" FLAG IN ROUGH/FINISH OPERATION
   ' ACAM PROGRAM IS IN THE FIRST QUADRANT, X+ & Y+
   '===============================================================
      
    Public Const sPostName As String = "Biesse_CIX_USA"
    Public Const sFirstLine As String = "BEGIN ID CID3"
    Public Const sLastLine As String = "'EOF"

    '===List of Core Modifications=====================================
    Public Const sPostRelease    As String = "3.0.3"
    'Rel 3.0.0 JEM 1/6/2022
    '   :Updated to new tracking Style
    '   :Added error message after selecting New Units
    '
    'Rel 3.0.1 JEM 1/11/2022
    '   : Oversite on customer license name resolved
    '
    'Rel 3.0.2 JEM 8/5/2022
    '   : Remove Post name check controll due to post being placed iin VBmacro startup causing issues
    '     when seleting another post.
    'Rel 3.0.3 JEM 1/12/2024
    '   :Added License ECO control to the post
    '
    '******* Notes below are from previous developers ********
    '3/4/2015  Rel 1.0.1: Error when calculating face offset for back face -HH.
    '3/6/2015  Rel 1.0.2: Added origins drop-down list - HH.
    '3/6/2015  Rel 1.0.3: Russell from Biesse pointed out a problem with FCN if working in MM. - HH.
    '3/6/2015  Rel 1.0.4: Issue not remembering the origin and persistent data - HH
    '3/6/2015  Rel 1.0.5: Daiek pointed out problem with the WSP being set to the SPEED, not the FEED - HH
    '3/6/2015  Rel 1.0.6: Daiek requests the ability to select the pins from the table - HH
    '6/26/2015 Rel 2.0.0: Daiek wants Pin control. Fair request - HH
    '8/4/2015  Rel 2.0.1: Brian Elliot from RSI found that the Sloping Lead Type is incorrect - HH
    '8/5/2015  Rel 2.0.2: Fixed small bug pertaining to the offset on the3D Lead - HH
    '8/5/2015  Rel 2.0.3: More Lead errors. Also, addressed a problem with Lead-in Feed when using Tool on Center - HH
    '          I also changed the decimal places and the tolerances because the system is having trouble calculating certain offsets - HH
    '8/6/2015  Rel 2.0.4: Had to adjust the precision to 4 decimals in MM and 6 in IN. Ridiculous. Also, added special cases to GetLeadData - HH
    '8/7/2015  Rel 2.0.5: Implemneted BOTTOM side operations on face WVF=5 - HH
    '8/14/2015 Rel 2.0.6: Even more issues with Arc Tol precision  - HH
    '5/25/2016 Rel 2.1.2 : Added support for Automatic Labeling - HH
    '5/28/2016 Rel 2.1.3 : Fixed minor issues with GUI and settings - HH
    '6/8/2016  Rel 2.1.5 Small update regarding settings
    '6/9/2016  Rel 2.1.6 - Added update to add Auto Label Information
    '6/17/2016 Rel 2.1.7 - Corrected Y-axis coordinate on the label position for Auto-labeling poer Giacomo (Top-Left Corner)
    '6/17/2016 Rel 2.1.9 - Added Folder name to  _lbl.xml file for BMP locations
    '6/21/2016 Rel 2.1.10 - Corrected mistake I made in the calculation of the Y-Axis coordinate of the part
    '6/22/2016 Rel 2.1.11 - Corrected mistake I made when finding the last a path if creatingPre-label data
    '6/22/2016 Rel 2.1.12 - I was appending label data even if the user said not to.
    '11/1/2016 Rel 2.1.13 - Issues with tagging and LX LY LZ rounding.
    '2/16/2017 Rel 2.1.20 - Several issues found at Trendwood: Leads and Jig Thickness variable.
    '3/8/2017  Rel 2.1.21 - The checkbox for Cust String is not persistent
    '3/13/2017 Rel 2.1.22 - Support for "slow down for corners"
    '4/25/2017 Rel 2.1.23 - Added support for outline note vs no note, to support Bridged Part Nesting
    '5/15/2017 Rel 2.1.24 - Small error when creating lbl info through the post
    '5/16/2017 Rel 2.1.25 - Once again, NPIs fuck me. Arrgh. Had to change the logic finding perimeter cuts if no tool note is used.
    '6/16/2017 Rel 2.1.27 - Modified the section on drilling per case Reported by Danny.
    '6/16/2017 Rel 2.1.28 - The previous Modification did NOT solve the issue. In the end, Drilling with which must be optimized must have OPT=YES
    '6/27/2017 Rel 2.1.29 - Unreliable behaviour when saving the Customer Header String
    '07/7/2017 Rel 2.1.30 - Final Modification to the AfterOpenPost Parameters for Tolerances.
    '9/12/2017 Rel 2.1.31 - Changed dMinArcLen from 1mm to 0.3mm per email from Fred.
    '8/23/2018 Rel 2.1.32 - Changed in Output Feed depth of cut for p.element.Slope FRW
    '9/26/2018 Rel 2.1.33 - Added line in Make Generic Bore to specify a Lancia bit. FRW
    '7/31/2019 Rel 2.1.34 - Adjusted tolerance of dDeltaZ for 3axis tool path. FRW
    '9/14/2021 Rel 2.1.35 - Resplved issue with Sawing TPD values.
    '9/16/2021 Rel 2.1.36 - Issues with selecting work plane due to compearing doubles.JEM
    '
    '==================================================================
 
    '==Customer Information / Options===========================================
    Public Const bShowCustomerInfo  As Boolean = True
    Public Const bCustomerHasMD     As Boolean = False ' Not used with CIX
    Public Const bListTools         As Boolean = True
    Public Const bCheckDate         As Boolean = False
    Public Const strExpiryDate      As String = "12/30/2021"
    Public Const sCustomerName      As String = "Kimball International Inc"
    Public Const sMultiDrillDB      As String = ""
    '==================================================================
    
    '==List of Customer Based Modifications=============
    Public Const sRev               As String = "V1"
    'V1 JEM 1/6/2022
    '  :Relase to Customer
    '
    '==================================================================
   
'Post file name
Public Const sPostNameControl As String = sPostName & "_" & "_" & sPostRelease & "_" & sRev
      
    'Added for ease of 3.0.0 conversion
    Public Const DEF_REL_INFO         As String = sPostNameControl
    Public Const DEF_POST_NAME        As String = sPostName
   
    Public DWG_Units        As Integer
    Public iDecimals        As Integer
    Public iWPid            As Integer
    Public iNextID          As Integer
    Public iFeedTypeHH      As Integer
    
    Public sUnits           As String
    Public sXoff            As String
    Public sYoff            As String
    Public sZoff            As String
    Public sOrigin          As String
    Public sSawPaths        As String
    Public sPins            As String
    Public sLastPath        As String
    Public sCustomUserStr   As String
    Public sXCUT            As String
    Public sYCUT            As String
    Public sImageType       As String
    Public sOutlineNote     As String
    
    Public lLastOp          As Long
        
    Public dist             As Double
    Public dProfileNo       As Double
    Public arc_error        As Double
    Public dToolDia         As Double
    Public dX               As Double
    Public dY               As Double
    Public dZ               As Double
    Public dFeedMult        As Double
    Public dLastZ           As Double
    Public dDeltaZ          As Double
    Public dFaceOffX        As Double 'Allows for the creation of "Face" code with origin =! bottom left
    Public dFaceOffY        As Double 'Allows for the creation of "Face" code with origin =! bottom left
    Public dFaceOffZ        As Double 'Allows for the creation of "Face" code with origin =! bottom left
    Public dMinArcLen       As Double 'Checks for the min length of an Arc. Will be replaced with LINE if smaller. Set in GetSettings
    Public dLastX           As Double
    Public dLastY           As Double
                    
    Public bSupLeadMsgs     As Boolean
    Public bUseCNCFeeds     As Boolean
    Public bNestingCalls    As Boolean
    Public bPromptDone      As Boolean
    Public bLeadDataDone    As Boolean
    Public bRapidDone       As Boolean
    Public bExtract         As Boolean
    Public bComp            As Boolean
    Public bFirstFeed       As Boolean
    Public bHideGUI         As Boolean 'Hides the interface for support of Automation (CDM, APM, hCAM, etc)
    Public bIgCustString    As Boolean
    Public bUseCustomUserSt As Boolean
    Public bFlagAutoLabel   As Boolean

    
    Public lyrLeadVectors   As Layer

    Public mintParkPos As Integer   ' will hold parking position 1,2 or 3
            
    Public Drw As Drawing
    
    Public Type BiesseLeadData
        bSlopeIn        As Boolean
        bSlopeOut       As Boolean
        dAngleIn        As Double
        dAngleOut       As Double
        dOffIn          As Double
        dOffOut         As Double
        dOverlap        As Double
        dMultiplier     As Double
        sInType         As String
        sOutType        As String
    End Type
    
    'Modifications:
    '3/4/2015  Rel 1.0.1: Error when calculating face offset for back face -HH.
    '3/6/2015  Rel 1.0.2: Added origins drop-down list - HH.
    '3/6/2015  Rel 1.0.3: Russell from Biesse pointed out a problem with FCN if working in MM. - HH.
    '3/6/2015  Rel 1.0.4: Issue not remembering the origin and persistent data - HH
    '3/6/2015  Rel 1.0.5: Daiek pointed out problem with the WSP being set to the SPEED, not the FEED - HH
    '3/6/2015  Rel 1.0.6: Daiek requests the ability to select the pins from the table - HH
    '6/26/2015 Rel 2.0.0: Daiek wants Pin control. Fair request - HH
    '8/4/2015  Rel 2.0.1: Brian Elliot from RSI found that the Sloping Lead Type is incorrect - HH
    '8/5/2015  Rel 2.0.2: Fixed small bug pertaining to the offset on the3D Lead - HH
    '8/5/2015  Rel 2.0.3: More Lead errors. Also, addressed a problem with Lead-in Feed when using Tool on Center - HH
    '          I also changed the decimal places and the tolerances because the system is having trouble calculating certain offsets - HH
    '8/6/2015  Rel 2.0.4: Had to adjust the precision to 4 decimals in MM and 6 in IN. Ridiculous. Also, added special cases to GetLeadData - HH
    '8/7/2015  Rel 2.0.5: Implemneted BOTTOM side operations on face WVF=5 - HH
    '8/14/2015 Rel 2.0.6: Even more issues with Arc Tol precision  - HH
    '5/25/2016 Rel 2.1.2 : Added support for Automatic Labeling - HH
    '5/28/2016 Rel 2.1.3 : Fixed minor issues with GUI and settings - HH
    '6/8/2016  Rel 2.1.5 Small update regarding settings
    '6/9/2016  Rel 2.1.6 - Added update to add Auto Label Information
    '6/17/2016 Rel 2.1.7 - Corrected Y-axis coordinate on the label position for Auto-labeling poer Giacomo (Top-Left Corner)
    '6/17/2016 Rel 2.1.9 - Added Folder name to  _lbl.xml file for BMP locations
    '6/21/2016 Rel 2.1.10 - Corrected mistake I made in the calculation of the Y-Axis coordinate of the part
    '6/22/2016 Rel 2.1.11 - Corrected mistake I made when finding the last a path if creatingPre-label data
    '6/22/2016 Rel 2.1.12 - I was appending label data even if the user said not to.
    '11/1/2016 Rel 2.1.13 - Issues with tagging and LX LY LZ rounding.
    '2/16/2017 Rel 2.1.20 - Several issues found at Trendwood: Leads and Jig Thickness variable.
    '3/8/2017  Rel 2.1.21 - The checkbox for Cust String is not persistent
    '3/13/2017 Rel 2.1.22 - Support for "slow down for corners"
    '4/25/2017 Rel 2.1.23 - Added support for outline note vs no note, to support Bridged Part Nesting
    '5/15/2017 Rel 2.1.24 - Small error when creating lbl info through the post
    '5/16/2017 Rel 2.1.25 - Once again, NPIs fuck me. Arrgh. Had to change the logic finding perimeter cuts if no tool note is used.
    '6/16/2017 Rel 2.1.27 - Modified the section on drilling per case Reported by Danny.
    '6/16/2017 Rel 2.1.28 - The previous Modification did NOT solve the issue. In the end, Drilling with which must be optimized must have OPT=YES
    '6/27/2017 Rel 2.1.29 - Unreliable behaviour when saving the Customer Header String
    '07/7/2017 Rel 2.1.30 - Final Modification to the AfterOpenPost Parameters for Tolerances.
    '9/12/2017 Rel 2.1.31 - Changed dMinArcLen from 1mm to 0.3mm per email from Fred.
    '8/23/2018 Rel 2.1.32 - Changed in Output Feed depth of cut for p.element.Slope FRW
    '9/26/2018 Rel 2.1.33 - Added line in Make Generic Bore to specify a Lancia bit. FRW
    '7/31/2019 Rel 2.1.34 - Adjusted tolerance of dDeltaZ for 3axis tool path. FRW
    '9/14/2021 Rel 2.1.35 - Resplved issue with Sawing TPD values.
    '9/16/2021 Rel 2.1.36 - Issues with selecting work plane due to compearing doubles.JEM
    
Public Sub OutputFileLeadingLines(p As PostData) '$10
    
    If App.AlphacamVersion.Major < 2022 Then
        MsgBox "Alphacam MultiCam 4axis post Supports Alphacam minimum Version 2022, Please contact Alphacam Support!, Email: acam.support@alphacam.com", vbCritical, sPostName
        p.Post "$EXIT"
    End If

    Const AC_ACAM_CUSTOM_DEVELOPMENT = 387

    Dim olic As License
    Set olic = App.License

    If (olic.IsNamedOptionLicensed(AC_ACAM_CUSTOM_DEVELOPMENT, "BiesseCix")) = False Then
        MsgBox "This Alphacam Biesse_CIX_USA Post is not licensed, Please contact Alphacam Support!, Email: acam.support@alphacam.com", vbCritical, sPostName
        p.Post "$EXIT"
    End If
    

    
    Functions.GetSettings
    Functions.fReadVolMat p
    
    If p.Vars.NSH > 1 Then
        bNestingCalls = True
    Else
        bNestingCalls = False
    End If
      
    iWPid = 5
    lLastOp = 0
    iNextID = 1100
    dLastZ = 0
    bExtract = False
    
    Set lyrLeadVectors = App.ActiveDrawing.CreateLayer("hhhLeadVectors")
    lyrLeadVectors.Visible = False
    
End Sub

Public Sub OutputProgramLeadingLines(p As PostData) '$12
               
    With frmMain
        Functions.fReadVolMat p
        If (Not bHideGUI) Then 'And (Not bPromptDone) Then
'            .Caption = DEF_REL_INFO
            .Caption = sPostNameControl
            .txtDX.Text = CStr(dX)
            .txtDY.Text = CStr(dY)
            .txtDZ.Text = CStr(dZ)
            .Show
            bPromptDone = True
            dX = CDbl(.txtDX)
            dY = CDbl(.txtDY)
            dZ = CDbl(.txtDZ)
        End If
        sZoff = .txtOffZ
        Call Functions.SaveSettings
        Call ClearWpAtts(p)
        Call LabelOverlaps
    End With
    
    If dX < 0.001 Or dY < 0.001 Or dZ < 0.001 Then
        MsgBox "Error reading Material Size!!!", vbCritical, "BiesseCIX_USA"
        End
    End If
    
    p.Post "BEGIN ID CID3"
    p.Post "   REL=5.0"
'    p.PostQuick "   '" & sPostNameControl
    
     If bShowCustomerInfo Then
        p.PostQuick "'========================================================================"
        'Customer Information - Post details
        p.PostQuick "   'Company: " & UCase(App.License.GetCustomerName)
        p.PostQuick "   '" & "Post: " & sPostName & " " & "Rel " & sPostRelease & " " & (sRev)
        p.PostQuick "   'ALPHACAM Version: " & App.AlphacamVersion.String
            
               'MultiDrill Database Infromation
            If bCustomerHasMD Then p.PostQuick "//Multidrill database: " & sMultiDrillDB
        p.PostQuick "'========================================================================"
     End If
    
    
    If App.ActiveDrawing.Name <> "" Then
        If gb_HasNesting Then
            p.PostQuick "   'Drawing: " & App.ActiveDrawing.Name & ".ard" & " Sheet_" & p.Path.Attribute("LicomUKsab_sheet_ident")
        Else
            p.PostQuick "   'Drawing: " & App.ActiveDrawing.Name & ".ard"
        End If
    Else
        If gb_HasNesting Then
            p.PostQuick "   'Drawing: Untilted.ard" & " Sheet_" & p.Path.Attribute("LicomUKsab_sheet_ident")
        Else
            p.PostQuick "   'Drawing: Untitled.ard"
        End If
    End If
    p.PostQuick "   'Program Time: " & p.Vars.TIM \ 60 & "min " & Round((p.Vars.TIM - (p.Vars.TIM \ 60) * 60), 0) & "sec"
    p.Post "END ID"
    p.Post ""
    
    p.Post "BEGIN MAINDATA"
    
    p.Post "   LPX=" & VBA.Trim(VBA.CStr(Round(Post.dX, 4)))
    p.Post "   LPY=" & VBA.Trim(VBA.CStr(Round(Post.dY, 4)))
    p.Post "   LPZ=" & VBA.Trim(VBA.CStr(Round(Post.dZ, 4)))
    p.Post "   ORLST=" & Chr(34) & sOrigin & Chr(34)
    p.Post "   " & Functions.GetUserDataString
    
    If sUnits = "MM" Then
        p.Post "   FCN=1"
    Else
        'Dont know why. Verified with Alex in bSolid
        p.Post "   FCN=25.4"
    End If
    
    If Post.sXCUT <> "" Then
        If CDbl(Post.sXCUT) > 0 Then
            p.Post "   XCUT=LPX-" & Post.sXCUT
        End If
    End If
    
    If Post.sYCUT <> "" Then
        If CDbl(Post.sYCUT) > 0 Then
            p.Post "   YCUT=" & Post.sYCUT
        End If
    End If
    
    p.Post "   PUTLST=" & Chr(34) & "1" & Chr(34)
    p.Post "   OPPWKRS=1"
    p.Post "   ENABLELABEL=1"

    'Jig Thickness Variable
    p.Post "   JIGTH=" & VBA.Trim(VBA.CStr(Round(Post.sZoff, iDecimals)))
    
    p.Post "END MAINDATA"
    p.Post ""

    If CDbl(Post.sXoff) <> 0 Or CDbl(Post.sYoff) Then
        p.Post "BEGIN MACRO"
        p.Post "   NAME=OFFSET"
        p.Post "   PARAM,NAME=X,VALUE=" & VBA.Trim(VBA.CStr(Round(Post.sXoff, iDecimals)))
        p.Post "   PARAM,NAME=Y,VALUE=" & VBA.Trim(VBA.CStr(Round(Post.sYoff, iDecimals)))
        'p.Post "   PARAM,NAME=Z,VALUE=" & VBA.Trim(VBA.CStr(-Round(Post.sZoff, iDecimals)))
        p.Post "   PARAM,NAME=SHW,VALUE=NO"
        p.Post "END MACRO"
        p.Post ""
    End If
    
End Sub

Public Sub OutputProgramTrailingLines(p As PostData) '$15
   p.PostQuick sLastLine
End Sub

Public Sub OutputFileTrailingLines(p As PostData) '$17

    lyrLeadVectors.Geometries.Delete
    lyrLeadVectors.Delete
    App.ActiveDrawing.Refresh
        
End Sub

Public Sub OutputRapid(p As PostData) '$20, $21 & $25
       
    'Biesse CIX Does NOT require RAPID Moves. All moves are done in FEED
    If Not MachiningAllowed(p) Then
        MsgBox "Error! Machining not permitted! " & Functions.RetMachineType(p)
        End
    End If
    
    bLeadDataDone = False
    bRapidDone = False
    
    'This is why we apply comp on rapid!
    If p.Vars.TC <> "0" Then
        bComp = True
    Else
        bComp = False
    End If
        
    If bExtract Then
        bExtract = False
        p.Post "BEGIN MACRO"
        p.Post "   NAME=ENDPATH"
        p.Post "END MACRO"
        p.Post ""
    End If
    
    sSawPaths = ""
    
End Sub

Public Sub OutputFeed(p As PostData) '$40, $50 & $60

    'Make Sure User Checked Box if Comp is selected...
    If p.Path.CompOnRapid = False Then
        If p.Path.McComp = acamCompMC And p.Path.IsDrilling = False Then
            Dim sPrompt As String
            sPrompt = "You have made cuts with tool compensation, " & Chr(10)
            sPrompt = sPrompt & "but you did not set the parameter: " & Chr(10)
            sPrompt = sPrompt & "'Apply Compensation on rapid approach/retract'." & Chr(10)
            sPrompt = sPrompt & "Please check your Roughing/Finishing passes and try again."
            sPrompt = sPrompt & "Operation: " & p.Vars.OPN
            MsgBox sPrompt, vbCritical, "Biesse CIX USA Post"
            End
        End If
    End If
    
    If Not MachiningAllowed(p) Then
        MsgBox "Error! Machining not permitted! " & Functions.RetMachineType(p)
        End
    End If

    If p.Tool.TPD(2) = "SAW" Then
        Call OutputSimpleSaw(p)
        Exit Sub
    End If

    Dim dPostNX     As Double
    Dim dPostNY     As Double
    Dim dPostNZ     As Double
    Dim dPostCX     As Double
    Dim dPostCY     As Double
    Dim dPostCZ     As Double
    Dim dPostI      As Double
    Dim dPostJ      As Double
    Dim sFace       As String
    Dim sPost       As String
    Dim TheLead     As BiesseLeadData
    
    If CInt(ReturnWP(p.WorkPlane, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)) > -1 Then
        sFace = ReturnWP(p.WorkPlane, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)
    Else
        sFace = GetWorkplaneID(p)
    End If

    'BiesseWorks is VERY civilised. The code for ANY face is the same.
    'You literally just change the Face and all machining applies from one face to another.
    'You can even copy a milling op to another face!
    
    'Global on top face
    If CInt(ReturnWP(p.WorkPlane, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)) = 0 Then
        dPostNX = p.Vars.GAX + dFaceOffX
        dPostNY = p.Vars.GAY + dFaceOffY
        dPostNZ = -p.Vars.GAZ + dFaceOffZ
        dPostCX = p.Vars.GCX + dFaceOffX
        dPostCY = p.Vars.GCY + dFaceOffY
        dPostCZ = -p.Vars.GCZ + dFaceOffZ
        dPostI = p.Vars.GAI + dFaceOffX
        dPostJ = p.Vars.GAJ + dFaceOffY
    Else
        dPostNX = p.Vars.AX + dFaceOffX
        dPostNY = p.Vars.AY + dFaceOffY
        dPostNZ = -p.Vars.AZ + dFaceOffZ
        dPostCX = p.Vars.CX + dFaceOffX
        dPostCY = p.Vars.CY + dFaceOffY
        dPostCZ = -p.Vars.CZ + dFaceOffZ
        dPostI = p.Vars.AI + dFaceOffX
        dPostJ = p.Vars.AJ + dFaceOffY
    End If
    
    If Not bLeadDataDone And bComp Then
        TheLead = GetLeadData(p)
    End If
    
    If Not bRapidDone Then 'Or bExtract Then
    
        'If bExtract Then bExtract = False
    
        'Rapids / Cut Setup
        p.Post "BEGIN MACRO"
        p.Post "   NAME=ROUT"
        p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & Replace(VBA.Trim(RetMachineType(p) & " " & p.Path.Name), " ", "_") & Chr(34) 'Toolpath name
        p.Post "   PARAM,NAME=SIDE,VALUE=" & VBA.Trim(sFace)
        p.Post "   PARAM,NAME=CRN,VALUE=" & Chr(34) & "2" & Chr(34) 'Always use the bottom-left reference on the face
        p.Post "   PARAM,NAME=Z,VALUE=0" 'Channel' Supposed to create an offset from the top face. See Doc.
        If bComp Then
            p.Post "   PARAM,NAME=DP,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNZ, iDecimals))) 'Depth of Cut
            dLastZ = dPostNZ
        Else
            If p.Element.Slope Then
                p.Post "   PARAM,NAME=DP,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNZ, iDecimals))) 'Depth of Cut '8-23-2018
                dLastZ = dPostNZ    '8-28-2018
'                p.Post "   PARAM,NAME=DP,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostCZ, iDecimals))) 'Depth of Cut
'                dLastZ = dPostCZ
            Else
                p.Post "   PARAM,NAME=DP,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNZ, iDecimals))) 'Depth of Cut
                dLastZ = dPostNZ
            End If
        End If
        'P.Post "   PARAM,NAME=ISO,VALUE=""" 'Possible expansion with direct ISO command?
        p.Post "   PARAM,NAME=OPT,VALUE=NO" 'Optimisable? Changed it due to the fact that BW can change the SECQUENCE of cuts by DEFAULT if using a BW configured as such!
        
        'For Display Purposes, Engraving will be dia 0.1"
        If p.Path.GetMillData.ProcessType2 = acamProcessENGRAVE Then
            p.Post "   PARAM,NAME=DIA,VALUE=0.1"
        Else
            p.Post "   PARAM,NAME=DIA,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Tool.Diameter, iDecimals)))  'Tool Diamater
        End If
        'P.Post "   PARAM,NAME=AZ,VALUE=0" 'Tilt data. Not used by bSolid. Workplane overrides (not used).
        'P.Post "   PARAM,NAME=AR,VALUE=0" 'Tilt data. Not used by bSolid. Workplane overrides (not used).
        'P.Post "   PARAM,NAME=CKA,VALUE=azrNO" 'Tilt data. Not used by bSolid. Workplane overrides (not used).
        p.Post "   PARAM,NAME=ER,VALUE=YES" '"It is recommended that this field be enabled at all times"
        'P.Post "   PARAM,NAME=A21,VALUE=0" 'Enables the use of the AGGRE42 aggregate and allows to enter the direction angle for the lead-in of the aggregate under the piece
        'P.Post "   PARAM,NAME=TOS,VALUE=NO" 'Related to Z. "When the adjacent TOS box is marked, during the calculation to establish the safety position, the value set in field Z is ignored"
        'P.Post "   PARAM,NAME=S21,VALUE=-1" ' Aggrgate Data. Face to be machined -1 is probably making sure we are not machining the wrong thing. Skipped sine bSolid no good.
        'P.Post "   PARAM,NAME=AZS,VALUE=0" ' Additional Safety.
        'Feed Data, depending on existance of leads
        If Not bUseCNCFeeds Then
            If p.Element.LeadIn Then
                p.Post "   PARAM,NAME=DSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, iDecimals)))  'G0/B Feed Rate
                p.Post "   PARAM,NAME=RSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.S, iDecimals)))  'Speed
                p.Post "   PARAM,NAME=IOS,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, iDecimals)))  'Lead Speed
            Else
                p.Post "   PARAM,NAME=DSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, iDecimals)))  'G0/B Feed Rate
                p.Post "   PARAM,NAME=RSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.S, iDecimals)))  'Speed
                p.Post "   PARAM,NAME=IOS,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, iDecimals)))  'Lead Speed
            End If
            'Working Feed
            p.Post "   PARAM,NAME=WSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.FC, iDecimals)))
        End If
        
        p.Post "   PARAM,NAME=TNM,VALUE=" & Chr(34) & VBA.Trim(p.Vars.TNM) & Chr(34) 'Tool Name

        
        'Tool Type. There is an exception for a SAW BLADE
        If p.Tool.Type = acamToolWHEEL Then
            'Disk Tool
            'P.Post "   PARAM,NAME=TTP,VALUE=tclCUT"
            p.Post "   PARAM,NAME=TCL,VALUE=2"
        Else
            'Any other tool
            'P.Post "   PARAM,NAME=TTP,VALUE=tclROUT"
            p.Post "   PARAM,NAME=TCL,VALUE=1"
        End If
        'Tool Comp
        p.Post "   PARAM,NAME=CRC,VALUE=" & VBA.Trim(p.Vars.TC)
        'Lead Data
        If bComp Then
            'Type of Lead-In
            p.Post "   PARAM,NAME=TIN,VALUE=" & TheLead.sInType
            'Angle
            p.Post "   PARAM,NAME=AIN,VALUE=" & TheLead.dAngleIn
            'Type of Lead-Out
            p.Post "   PARAM,NAME=TOU,VALUE=" & TheLead.sOutType
            'Angle
            p.Post "   PARAM,NAME=AOU,VALUE=" & TheLead.dAngleOut
            'Overlap
            p.Post "   PARAM,NAME=DOU,VALUE=" & TheLead.dOverlap
            'Multiplier
            p.Post "   PARAM,NAME=PRP,VALUE=" & TheLead.dMultiplier
            'Comp on Rapid
            p.Post "   PARAM,NAME=CIN,VALUE=YES"
            p.Post "   PARAM,NAME=COU,VALUE=YES"
        Else
            'Type of Lead-In
            p.Post "   PARAM,NAME=TIN,VALUE=0"
            'Angle
            p.Post "   PARAM,NAME=AIN,VALUE=0"
            'Type of Lead-Out
            p.Post "   PARAM,NAME=TOU,VALUE=0"
            'Angle
            p.Post "   PARAM,NAME=AOU,VALUE=0"
            'Overlap
            p.Post "   PARAM,NAME=DOU,VALUE=0"
            'Multiplier
            p.Post "   PARAM,NAME=PRP,VALUE=0"
            'Comp on Rapid
            p.Post "   PARAM,NAME=CIN,VALUE=NO"
            p.Post "   PARAM,NAME=COU,VALUE=NO"
        End If
        
        'Side offset for lead type 7 (sloping lead)
        If TheLead.sInType = "7" Then
            p.Post "   PARAM,NAME=GIN,VALUE=" & TheLead.dOffIn
        Else
            p.Post "   PARAM,NAME=GIN,VALUE=0"
        End If
        
        If TheLead.sOutType = "7" Then
            p.Post "   PARAM,NAME=GOU,VALUE=" & TheLead.dOffOut
        Else
            p.Post "   PARAM,NAME=GOU,VALUE=0"
        End If
        
        'Other Params not used in Lead Data:
        p.Post "   PARAM,NAME=DIN,VALUE=0"
        p.Post "   PARAM,NAME=TBI,VALUE=NO"
        p.Post "   PARAM,NAME=TQI,VALUE=0"
        p.Post "   PARAM,NAME=TLI,VALUE=0"
        p.Post "   PARAM,NAME=TBO,VALUE=NO"
        p.Post "   PARAM,NAME=TQO,VALUE=0"
        p.Post "   PARAM,NAME=TLO,VALUE=0"
                
        'Support for blower
        If p.Path.GetMillData.Coolant = acamCoolNONE Then
            p.Post "   PARAM,NAME=BFC,VALUE=NO"
        Else
            p.Post "   PARAM,NAME=BFC,VALUE=YES"
        End If
        
        'Support for dust collection hood
        If p.Tool.TPD(1) <> "" Then
            'User-defined
            p.Post "   PARAM,NAME=SHP,VALUE=" & p.Tool.TPD(1)
        Else
            'Automatic
            p.Post "   PARAM,NAME=SHP,VALUE=0"
        End If
                
        p.Post "END MACRO"
        p.Post ""
        
        'Start Point
        If Not bComp Then
            p.Post "BEGIN MACRO"
            p.Post "   NAME=START_POINT"
            p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & "StPt_" & GetNextID & Chr(34) 'Sequntial ID Number
            p.Post "   PARAM,NAME=X,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostCX, iDecimals)))
            p.Post "   PARAM,NAME=Y,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostCY, iDecimals)))
            p.Post "   PARAM,NAME=Z,VALUE=0"
            p.Post "END MACRO"
            p.Post ""
        End If

        bRapidDone = True
        bFirstFeed = True
        
    End If ' - If Not bRapidDone Then
            
    'Compensated Leads must use the Biesse CIX Specification. We do not post them.
    If bComp Then
        If p.Element.LeadIn Or p.Element.LeadOut Or p.Element.IsRapid Then GoTo NoPosting
        If p.Element.Attribute("_hhhOverlapElement") = 1 Then GoTo NoPosting
    End If
    
    If bFirstFeed And bComp Then
        p.Post "BEGIN MACRO"
        p.Post "   NAME=START_POINT"
        p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & "StPt_" & GetNextID & Chr(34) 'Sequntial ID Number
        p.Post "   PARAM,NAME=X,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostCX, iDecimals)))
        p.Post "   PARAM,NAME=Y,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostCY, iDecimals)))
        p.Post "   PARAM,NAME=Z,VALUE=0"
        p.Post "END MACRO"
        p.Post ""
        bFirstFeed = False
    End If
        
    p.Post "BEGIN MACRO"
    
    iFeedTypeHH = p.FeedType
    If p.Element.IsArc Then
        If p.Element.IncludedAngle < 350 Then
            If (Sqr((dPostNX - dLastX) ^ 2 + (dPostNY - dLastY) ^ 2) < dMinArcLen) Or ((p.Element.Length < dMinArcLen) And p.Vars.TC = 0) Then
                iFeedTypeHH = 1
            End If
        End If
    End If
    
    'Feeds
    Select Case iFeedTypeHH
        Case 1
            p.Post "   NAME=LINE_EP"
            p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & Replace(Trim(p.Path.Name) & "_" & (p.Element.Name), " ", "_") & Chr(34)
        Case 2
            p.Post "   NAME=ARC_EPCE"
            p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & Replace(Trim(p.Path.Name) & "_" & (p.Element.Name), " ", "_") & Chr(34)
            p.Post "   PARAM,NAME=DIR,VALUE=dirCW"
            p.Post "   PARAM,NAME=XC,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostI, iDecimals)))
            p.Post "   PARAM,NAME=YC,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostJ, iDecimals)))
        Case 3
            p.Post "   NAME=ARC_EPCE"
            p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & Replace(Trim(p.Path.Name) & "_" & (p.Element.Name), " ", "_") & Chr(34)
            p.Post "   PARAM,NAME=DIR,VALUE=dirCCW"
            p.Post "   PARAM,NAME=XC,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostI, iDecimals)))
            p.Post "   PARAM,NAME=YC,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostJ, iDecimals)))
    End Select
    
    p.Post "   PARAM,NAME=XE,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNX, iDecimals)))
    p.Post "   PARAM,NAME=YE,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNY, iDecimals)))
    
    'Can you believe this??? Only Z is INCREMENTAL!?
    
    'Record DeltaZ for actually posted items:
    dDeltaZ = dPostNZ - dLastZ
    dLastZ = dPostNZ
    
'    If (dDeltaZ > 0.001) Or (dDeltaZ < -0.001) Then
    If (dDeltaZ > 0.0001) Or (dDeltaZ < -0.0001) Then   '7-31-2019 FRW
        p.Post "   PARAM,NAME=ZE,VALUE=" & VBA.Trim(VBA.CStr(Round(dDeltaZ, iDecimals)))
    End If
    p.Post "   PARAM,NAME=SC,VALUE=scOFF"
    
    If Not bUseCNCFeeds Then
        'Feeds and Speeds
        If p.Element.LeadIn Then
            'Alpha Bug: Reported F as zero on first feed...
            If p.Vars.f > 0 Then
                p.Post "   PARAM,NAME=FD,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, 0)))
            Else
                p.Post "   PARAM,NAME=FD,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.FD, 0)))
            End If
            
        Else
            'Alpha Bug: Reported F as zero on first feed...
            If p.Vars.f > 0 Then
                p.Post "   PARAM,NAME=FD,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, 0)))
            Else
                p.Post "   PARAM,NAME=FD,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.FC, 0)))
            End If
        End If
        p.Post "   PARAM,NAME=SP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.S, 0)))
    End If
    
    p.Post "END MACRO"
    p.Post ""
    
NoPosting:

    'Flag the Last Feed to Close the Macro in the Rapid Section:
    'If p.Vars.LF > 0 Then
    'If ((p.Path.GetLastElem.IsSame(p.Element)) Or IsNextEleLastAndZeroLength(p.Element)) And p.Vars.LF > 0 Then
    If IsElementLast(p.Element, p.Path) Then
        bExtract = True
        sLastPath = p.Path.Name
    End If
    
    dLastX = dPostNX
    dLastY = dPostNY
    
End Sub


Public Sub OutputCancelTool(p As PostData) '$70
End Sub

Public Sub OutputSelectTool(p As PostData) '$80
    
   If p.Vars.CLT = "M8" Then
      p.Post "BEGIN MACRO"
      p.PostQuick "   NAME=ISO"
      p.Post "   PARAM,NAME=ISO,VALUE=""WL(PRWL)=1 M80"""
      p.PostQuick "END MACRO"
      p.PostQuick " "
   End If
    
End Sub

Public Sub OutputSelectWorkPlane(p As PostData) '$88

End Sub

Public Sub OutputSelectToolAndWorkPlane(p As PostData) '$89
End Sub

Public Sub OutputCallSub(p As PostData) '$90
End Sub

Public Sub OutputBeginSub(p As PostData) '$100
End Sub

Public Sub OutputEndSub(p As PostData) '$110
End Sub

Public Sub OutputOriginShift(p As PostData) '$120
End Sub

Public Sub OutputCancelOriginShift(p As PostData) '$130
End Sub

Public Sub OutputDrillCycleCancel(p As PostData) '$200
End Sub

Public Sub OutputFirstHoleSub(p As PostData) '$205
End Sub

Public Sub OutputNextHoleSub(p As PostData) '$206
End Sub

Public Sub OutputDrillCycleFirstHole(p As PostData) '$210, 214, 220, 224, 230, 234, 240 & 244

   Call MakeGenericBore(p)
            
End Sub

Public Sub OutputDrillCycleNextHoles(p As PostData) '$211, 215, 221, 225, 231, 235, 241 & 245

    Call MakeGenericBore(p)

End Sub

Public Sub OutputDrillCycleSubParameters(p As PostData) '$212, 216, 222, 226, 232, 236, 242 & 246
End Sub

Private Function ReturnWP(ByRef wp As WorkPlane, dDisX As Double, dDisY As Double, dDisZ As Double) As String

    Dim sSide As String
    sSide = "0"
    
    'Applies offsets to the WP origin, only if using cardinal FACES
    dFaceOffX = 0
    dFaceOffY = 0
    dFaceOffZ = 0

    Select Case wp.WVF
        Case 0
            'Top
            sSide = "0"
        Case 1
            'Front
            If dDisY = 0 Then
                sSide = "2"
                dFaceOffX = dDisX
            Else
                sSide = "-1"
            End If
        Case 2
            'Right
            If Round(dDisX, 4) = Round(dX, 4) Then
                sSide = "3"
                dFaceOffX = dDisY
            Else
                sSide = "-1"
            End If

        Case 3
            'Back
            If Round(dDisY, 4) = Round(dY, 4) Then
                sSide = "4"
                dFaceOffX = dX - dDisX
            Else
                sSide = "-1"
            End If
        Case 4
            'Left
            If dDisX = 0 Then
                sSide = "1"
                dFaceOffX = dY - dDisY
            Else
                sSide = "-1"
            End If
        Case 5
            'Bottom
            sSide = "5"
            dFaceOffX = dX - dDisX
            dFaceOffY = dDisY
        Case Else
            'Any INCLINED Plane
            sSide = "-1"
    End Select
    
    If CInt(sSide) > 0 Then
        dFaceOffY = Post.dZ + dDisZ
    End If
    
    ReturnWP = sSide
    
End Function

Private Function GetNextID() As String
    iNextID = iNextID + 1
    GetNextID = VBA.Trim(VBA.CStr(iNextID))
End Function

'Private Function GetLeadData2(ByRef p As PostData) As BiesseLeadData
'
'    Dim op          As Operation
'    Dim subop       As SubOperation
'    Dim SubOp2      As SubOperation
'
'    Dim SubPth      As Path
'
'    Dim md          As MillData
'
'    Dim ld          As LeadData
'    Dim bld         As BiesseLeadData
'
'    Dim bFoundIt    As Boolean
'
'    Dim dMultiIN    As Double
'    Dim dMultiOUT   As Double
'
'    Dim sMsg        As String
'
'    bFoundIt = False
'    sMsg = ""
'
'    With bld
'        .bSlopeIn = False
'        .bSlopeOut = False
'        .dAngleIn = 0
'        .dAngleOut = 0
'        .dMultiplier = 0
'        .dOverlap = 0
'        .sInType = "0"
'        .sOutType = "0"
'    End With
'
'    Set op = App.ActiveDrawing.Operations.Item(p.Path.OpNo)
'
'    For Each subop In op.SubOperations
'        For Each SubPth In subop.ToolPaths
'            If p.Path.Name = SubPth.Name Then
'                Set ld = subop.GetMillData.GetLeadData
'                If Not ld Is Nothing Then
'                    bFoundIt = True
'                End If
'                Exit For
'            End If
'        Next SubPth
'        If bFoundIt Then Exit For
'    Next subop
'
'    If Not bFoundIt Then GoTo NoLeadData
'
'    'Types of Biesse Lead:
''    Values allowed:
''    • 0 = None
''    • 1 = Curve
''    • 2 = Line
''    • 3 = Tg LineCurve
''    • 5 = Helix
''    • 6 = 3DLineCurve
''    • 7 = Corrected 3DLine
''    • 8 = Corrected 3DCurve
''    • 9 = Corrected Line
''    • 14 = 3D profile
'
'
'
'    'Error Catching
'    Select Case ld.LeadIn
'        Case acamLeadLINE
'            If ld.SlopeIn Then
'                bld.sInType = "7"
'                ld.AngleIn = 60
'            Else
'                bld.sInType = "2"
'            End If
'            dMultiIN = 0
'        Case acamLeadARC
'            If ld.SlopeIn Then
'                bld.sInType = "6"
'            Else
'                bld.sInType = "1"
'            End If
'            dMultiIN = ld.RadiusIn * 100
'            ld.AngleIn = ld.AngleIn / 2
'        Case acamLeadBOTH
'            If ld.SlopeIn Then
'                bld.sInType = "6"
'            Else
'                bld.sInType = "3"
'            End If
'            dMultiIN = ld.RadiusIn * 100
'        Case acamLeadBOTH_NOT_TANGENTIAL
'            sMsg = "Error! The CIX Specification does NOT allow the Lead Type 'Both-Non-Tangent' " & Chr(10)
'            sMsg = sMsg & "IF Using G41/G42!" & Chr(10)
'            sMsg = sMsg & "To suppress this message, check 'Supress Lead Error Messages'" & Chr(10)
'            sMsg = sMsg & "in the Config Screen. You may choose ANY Lead if using 'Tool On Center'" & Chr(10)
'            sMsg = sMsg & "Do you wish to continue posting using Lead-In Type LINE?"
'            If ld.SlopeIn Then
'                bld.sInType = "6"
'            Else
'                bld.sInType = "2"
'            End If
'             dMultiIN = 0
'        Case acamLeadNONE
'            bld.sInType = "0"
'            dMultiIN = 0
'    End Select
'
'    Select Case ld.LeadOut
'        Case acamLeadLINE
'            If ld.SlopeOut Then
'                If ld.AngleOut = 0 Then
'                    bld.sOutType = "7"
'                    ld.AngleOut = 60
'                Else
'                    bld.sOutType = "6"
'                End If
'            Else
'                bld.sOutType = "2"
'            End If
'            dMultiOUT = 0
'        Case acamLeadARC
'            If ld.SlopeOut Then
'                bld.sOutType = "6"
'            Else
'                bld.sOutType = "1"
'            End If
'            dMultiOUT = ld.RadiusOut * 100
'            ld.AngleOut = ld.AngleOut / 2
'        Case acamLeadBOTH
'            If ld.SlopeOut Then
'                bld.sOutType = "6"
'            Else
'                bld.sOutType = "3"
'            End If
'            dMultiOUT = ld.RadiusOut * 100
'        Case acamLeadBOTH_NOT_TANGENTIAL
'            sMsg = "Error! The CIX Specification does NOT allow the Lead Type 'Both-Non-Tangent' " & Chr(10)
'            sMsg = sMsg & "IF Using G41/G42!" & Chr(10)
'            sMsg = sMsg & "To suppress this message, check 'Supress Lead Error Messages'" & Chr(10)
'            sMsg = sMsg & "in the Config Screen. You may choose ANY Lead if using 'Tool On Center'" & Chr(10)
'            sMsg = sMsg & "Do you wish to continue posting using Lead-Out Type LINE?"
'            If ld.SlopeIn Then
'                bld.sInType = "6"
'            Else
'                bld.sInType = "2"
'            End If
'            dMultiOUT = 0
'        Case acamLeadNONE
'            bld.sInType = "0"
'            dMultiOUT = 0
'    End Select
'
''    If sUnits = "MM" Then
''        If (Abs(Abs(dMultiIN / 100 * P.Tool.Diameter / 2) - Abs(dMultiOUT / 100 * P.Tool.Diameter / 2))) > 0.05 Then
''            bAsym = True
''        End If
''    Else
''        If (Abs(Abs(dMultiIN / 100 * P.Tool.Diameter / 2) - Abs(dMultiOUT / 100 * P.Tool.Diameter / 2))) > 0.001 Then
''            bAsym = True
''        End If
''    End If
''
''    If bAsym Then
''        sMsg = "Error! The CIX Specification does NOT allow the LENGTH " & Chr(10)
''        sMsg = sMsg & "of the Lead-In to be different from the length " & Chr(10)
''        sMsg = sMsg & "of the Lead-Out, IF Using G41/G42!" & Chr(10)
''        sMsg = sMsg & "To suppress this message, check 'Supress Lead Error Messages'" & Chr(10)
''        sMsg = sMsg & "in the Config Screen. You may choose ANY Lead if using 'Tool On Center'" & Chr(10)
''        sMsg = sMsg & "Do you wish to continue posting using the length of the SHORTEST Lead?"
''        If dMultiIN > dMultiOUT Then dMultiIN = dMultiOUT
''        If dMultiOUT > dMultiIN Then dMultiOUT = dMultiIN
''    End If
'
'    If ((ld.AngleIn > 90) Or (ld.AngleOut > 90)) Then
'        sMsg = "Error! The CIX Specification does NOT allow the ANGLE" & Chr(10)
'        sMsg = sMsg & "of the Lead-In OR the Lead-Out to EXCEED 90 Degrees!" & Chr(10)
'        sMsg = sMsg & "IF Using G41/G42!" & Chr(10)
'        sMsg = sMsg & "To suppress this message, check 'Supress Lead Error Messages'" & Chr(10)
'        sMsg = sMsg & "in the Config Screen. You may choose ANY Lead if using 'Tool On Center'" & Chr(10)
'        sMsg = sMsg & "Do you wish to continue posting using a max Angle of 90?"
'        If ld.AngleIn >= 90 Then ld.AngleIn = 89.999
'        If ld.AngleOut >= 90 Then ld.AngleOut = 89.999
'    End If
'
'    'Error Display
'    If sMsg <> "" And Not bSupLeadMsgs Then
'        If MsgBox(sMsg, vbYesNo, "Lead Data Error!") = vbNo Then
'            End
'        Else
'            bSupLeadMsgs = True
'        End If
'    End If
'
'    bld.bSlopeIn = ld.SlopeIn
'    bld.bSlopeOut = ld.SlopeOut
'    If ld.LeadIn <> acamLeadNONE Then
'        bld.dAngleIn = Round(ld.AngleIn, iDecimals)
'    Else
'        bld.dAngleIn = 0
'    End If
'    If ld.LeadOut <> acamLeadNONE Then
'        bld.dAngleOut = Round(ld.AngleOut, iDecimals)
'    Else
'        bld.dAngleOut = 0
'    End If
'    bld.dOverlap = Round(ld.Overlap, iDecimals)
'    bld.dMultiplier = Round(dMultiIN, 0)
'
'NoLeadData:
'
'    bLeadDataDone = True
'    GetLeadData2 = bld
'
'End Function

Private Function GetLeadData(ByRef p As PostData) As BiesseLeadData
     
    Dim tp          As Path
    Dim tp2         As Path
    
    Dim md          As MillData
    
    Dim eleIN       As Element
    Dim eleOUT      As Element
    Dim Ele1        As Element
    Dim Ele2        As Element
    
    Dim bIN         As Boolean
    Dim bOUT        As Boolean
    
    Dim dOverlap    As Double
    Dim dToolRad    As Double
    Dim dLenIN      As Double
    Dim dLenOUT     As Double
    Dim dTopZ       As Double
    Dim dBotZ       As Double
    Dim dSin        As Double
    Dim dDeg        As Double
    
    Dim sTypeIn     As String
    Dim sTypeOut    As String
    
    Dim bld         As BiesseLeadData
    
    Dim iFace       As Integer
    
    Dim sMsg As String
    
    With bld
        .bSlopeIn = False
        .bSlopeOut = False
        .dAngleIn = 0
        .dAngleOut = 0
        .dOffIn = 0
        .dOffOut = 0
        .dMultiplier = 0
        .dOverlap = 0
        .sInType = "0" '0=None
        .sOutType = "0" '0=None
    End With
    
    bIN = False
    bOUT = False
    sTypeIn = ""
    sTypeOut = ""
    dLenIN = 0
    dLenOUT = 0
    
    'Get Path and MillData
    Set tp = p.Path
    Set md = tp.GetMillData
    dToolRad = p.Tool.Diameter / 2

    'Get First Element, which SHOULD be Lead-In
    Set eleIN = tp.GetFirstElem
    
    'Ensure Element is from the PATH and not the Rapid
    While eleIN.IsRapid
        Set eleIN = eleIN.GetNext
    Wend
    
    'Find Type
    If eleIN.LeadIn Then
        If eleIN.IsLine Then
            If eleIN.GetNext.LeadIn And eleIN.GetNext.IsArc Then
                sTypeIn = "Both"
            Else
                sTypeIn = "Line"
            End If
        Else
            sTypeIn = "Arc"
        End If
    End If
    
    'Get Last Element, which SHOULD be Lead-Out (if exists)
    Set eleOUT = tp.GetLastElem
    
    'Ensure Element is from the PATH and not the Rapid
    While eleOUT.IsRapid
        Set eleOUT = eleOUT.GetPrevious
    Wend
    
    'Find Type
    If eleOUT.LeadOut Then
        If eleOUT.IsLine Then
            If eleOUT.GetPrevious.LeadOut And eleOUT.GetPrevious.IsArc Then
                sTypeOut = "Both"
            Else
                sTypeOut = "Line"
            End If
        Else
            sTypeOut = "Arc"
        End If
    End If
    
    bIN = eleIN.LeadIn
    bOUT = eleOUT.LeadOut
    
    If Not bIN And Not bOUT Then
        'Return since no lead data is available
        bld.sInType = "0"
        bld.sOutType = "0"
        GetLeadData = bld
        Exit Function
    End If
        
'    Types of Biesse Lead:
'    Values allowed:
'    • 0 = None
'    • 1 = Curve
'    • 2 = Line
'    • 3 = Tg LineCurve
'    • 5 = Helix
'    • 6 = 3DLineCurve
'    • 7 = Corrected 3DLine
'    • 8 = Corrected 3DCurve
'    • 9 = Corrected Line
'    • 14 = 3D profile
    
    'A Standard Alphacam SLOPING lead will be considered a CORRECTED3DLINE. The problem is that one must calculate the
    'amount of slope via an ANGLE and the actual alphacam angle is an offset from the SIDE. Yuck.
    
    'Calculate Data Needed for BS
    If bIN Then

        bld.bSlopeIn = eleIN.Slope
        'i HATE alphacam sometimes... found a bug where an element is NOT flagged properly as a slope in Helical Milling
        'so now I must check by hand...
        If Not bld.bSlopeIn And eleIN.Is3D Then
            If Abs(Abs(eleIN.StartXL) - Abs(eleIN.EndXL)) > 0.001 Or _
               Abs(Abs(eleIN.StartYL) - Abs(eleIN.EndYL)) > 0.001 And _
               Abs(Abs(eleIN.StartZL) - Abs(eleIN.EndZL)) < 0.001 Then
               bld.bSlopeIn = True
            End If
        End If
    
        Select Case sTypeIn
            Case "Line"
                'Find Lead Angle
                Set Ele1 = eleIN
                While Ele1.LeadIn
                    Set Ele1 = Ele1.GetNext
                Wend
                
                'Normal Angle. WIll be overriden if SLOPING
                bld.dAngleIn = Abs(eleIN.AngleToElement(Ele1))
                
                'Special case when the first element after lead is an arc at 90
                If Ele1.IsArc And bld.dAngleIn = 90 Then bld.dAngleIn = 0
                
                If bld.bSlopeIn Then
                
                    'As stated above, all sloping leads are 7. but we must now find the "reciprocating" information
                    'needed for Biesse. their ANGLE is the Angle of the lead SLOPE and the Side angle is a distance called offset. Yuck.
                    
                    bld.sInType = "7"
                    
                    'Find the offset from the side (Alphacam Lead Angle)
                    dLenIN = eleIN.Length2D
                    bld.dOffIn = Sin(ToRadians(bld.dAngleIn)) * dLenIN / 2
                    bld.dOffIn = Round(bld.dOffIn, iDecimals)
                    'Simple fix for angles > 30
                    If bld.dAngleIn > 30 Then
                        bld.dOffIn = p.Tool.Diameter / 2
                    End If
                    
                    'Offset must be capped to the tool diameter
                    If bld.dOffIn > p.Tool.Diameter Then bld.dOffIn = Round(p.Tool.Diameter, iDecimals)
                    
                    'Find the angle of the slope:
                    dLenIN = eleIN.Length
                    
                    dTopZ = p.Vars.CZ
                    dBotZ = p.Vars.AZ
            
                    dSin = (dTopZ - dBotZ) / dLenIN
                    If dSin < 1 Then
                        dDeg = ArcSin(dSin)
                    Else
                        dDeg = 45
                    End If
                    
                    'Assign new angle
                    bld.dAngleIn = Round(dDeg, iDecimals)
                    
                    'Multiplier is zero for this lead
                    bld.dMultiplier = 0
                    
                Else
                    bld.sInType = "2"
                    bld.dMultiplier = eleIN.Length2D / dToolRad * 100
                End If
            Case "Arc"
                'Find Lead Angle
                bld.dAngleIn = Abs(FindArcLeadAngle(eleIN))
                
                If bld.bSlopeIn Then
                    bld.sInType = "6"
                Else
                    bld.sInType = "1"
                End If
                
                bld.dMultiplier = eleIN.radius / dToolRad * 100
                bld.dAngleIn = bld.dAngleIn / 2
                
            Case "Both"
                'Find Lead Angle
                Set Ele1 = eleIN
                While Ele1.LeadIn
                    Set Ele1 = Ele1.GetNext
                Wend
                
                bld.dAngleIn = Abs(eleIN.AngleToElement(Ele1))
                
                If bld.bSlopeIn Then
                    bld.sInType = "6"
                Else
                    bld.sInType = "3"
                End If
                bld.dMultiplier = eleIN.Length2D / dToolRad * 100
            Case Else
                bld.sInType = "0"
        End Select
    End If
    
    If bOUT Then
        
        bld.bSlopeOut = eleOUT.Slope
        'i HATE alphacam sometimes... found a bug where an element is NOT flagged properly as a slope in Helical Milling
        'so now I must check by hand...
        If Not bld.bSlopeOut And eleOUT.Is3D Then
            If Abs(Abs(eleOUT.StartXL) - Abs(eleOUT.EndXL)) < 0.001 Or _
               Abs(Abs(eleOUT.StartYL) - Abs(eleOUT.EndYL)) < 0.001 And _
               Abs(Abs(eleOUT.StartZL) - Abs(eleOUT.EndZL)) > 0.001 Then
               bld.bSlopeOut = True
            End If
        End If
    
        Select Case sTypeOut
            Case "Line"
                'Find Lead Angle
                Set Ele1 = eleOUT
                While Ele1.LeadOut
                    Set Ele1 = Ele1.GetPrevious
                Wend
                
                bld.dAngleOut = Abs(eleOUT.AngleToElement(Ele1))
                
                'Special case when the last element before lead is an arc at 90
                If Ele1.IsArc And bld.dAngleOut = 90 Then bld.dAngleOut = 0
                
                If bld.bSlopeOut Then
                
                    bld.sOutType = "7"
                    
                    'Find the offset from the side (Alphacam Lead Angle)
                    dLenOUT = eleOUT.Length2D
                    bld.dOffOut = Sin(ToRadians(bld.dAngleOut)) * dLenOUT / 2
                    bld.dOffOut = Round(bld.dOffOut, iDecimals)
                    
                    'Simple fix for angles > 30
                    If bld.dAngleOut > 30 Then
                        bld.dOffOut = p.Tool.Diameter / 2
                    End If
                    
                    'Offset must be capped to the tool diameter
                    If bld.dOffOut > p.Tool.Diameter Then bld.dOffOut = Round(p.Tool.Diameter, iDecimals)
                                        
                    'Find the angle of the slope:
                    dLenOUT = eleOUT.Length
                    
                    dTopZ = p.Vars.CZ
                    dBotZ = p.Vars.AZ
                    
                    If dLenIN = 0 Then
                        dSin = 1
                    Else
                        dSin = (dTopZ - dBotZ) / dLenIN
                    End If
                    
                    If dSin < 1 Then
                        dDeg = ArcSin(dSin)
                    Else
                        dDeg = 45
                    End If
                    
                    bld.dAngleOut = Round(dDeg, iDecimals)
                    
                Else
                    bld.sOutType = "2"
                End If
                
            Case "Arc"
                
                'Find Lead Angle
                bld.dAngleOut = Abs(FindArcLeadAngle(eleOUT))
                
                If bld.bSlopeOut Then
                    bld.sOutType = "6"
                Else
                    bld.sOutType = "1"
                End If
            
                bld.dAngleOut = bld.dAngleOut / 2
                
            Case "Both"
                'Find Lead Angle
                Set Ele1 = eleOUT
                While Ele1.LeadOut
                    Set Ele1 = Ele1.GetPrevious
                Wend
                
                bld.dAngleOut = Abs(eleOUT.AngleToElement(Ele1))
                If bld.bSlopeOut Then
                    bld.sOutType = "6"
                Else
                    bld.sOutType = "3"
                End If
            Case Else
                bld.sOutType = 0
        End Select
    End If
    
    'Error Catching
    If ((bld.dAngleIn > 90.0001) Or (bld.dAngleOut > 90.0001)) Then
        sMsg = "Error! The CIX Specification does NOT allow the ANGLE" & Chr(10)
        sMsg = sMsg & "of the Lead-In OR the Lead-Out to EXCEED 90 Degrees!" & Chr(10)
        sMsg = sMsg & "IF Using G41/G42! Toolpath: " & p.Path.Name & " Element: " & p.Element.Name & Chr(10)
        sMsg = sMsg & "To suppress this message, check 'Supress Lead Error Messages'" & Chr(10)
        sMsg = sMsg & "in the Config Screen. You may choose ANY Lead if using 'Tool On Center'" & Chr(10)
        sMsg = sMsg & "Do you wish to continue posting using a max Angle of 90?"
    End If
    
    'Insurance
    If bld.dAngleIn >= 90 Then bld.dAngleIn = 89.999
    If bld.dAngleOut >= 90 Then bld.dAngleOut = 89.999
    
    'Error Display
    If sMsg <> "" And Not bSupLeadMsgs Then
        If MsgBox(sMsg, vbYesNo, "Lead Data Error!") = vbNo Then
            End
        Else
            bSupLeadMsgs = True
        End If
    End If
    
    bld.dAngleIn = Round(bld.dAngleIn, 2)
    bld.dAngleOut = Round(bld.dAngleOut, 2)
    bld.dMultiplier = Round(bld.dMultiplier, 0)
    bld.dOffIn = Round(bld.dOffIn, iDecimals)
    bld.dOffOut = Round(bld.dOffOut, iDecimals)
    bld.dOverlap = Round(GetOverlapLength(p.Path), iDecimals)
    
    bLeadDataDone = True

    GetLeadData = bld

End Function

Private Function GetElementAngle2D(ByRef Ele As Element, iFaccia As Integer) As Double

    Dim dStX        As Double
    Dim dStY        As Double
    Dim dEndX       As Double
    Dim dEndY       As Double
    Dim dDeltaX     As Double
    Dim dDeltaY     As Double
    Dim dResult     As Double
    Dim pi          As Double
    
    pi = 4 * Atn(1)
    dResult = 0
    
    dStX = Ele.StartXL
    dStY = Ele.StartYL
    dEndX = Ele.EndXL
    dEndY = Ele.EndYL
    dDeltaX = dEndX - dStX
    dDeltaY = dEndY - dStY
    
    If dDeltaX > -0.0000001 And dDeltaX < 0.0000001 Then
        'Zero Delta X (I HATE ATN(Penis))
        If dEndY > 0 Then
            If dEndY > dStY Then
                dResult = 90
            Else
                dResult = 270
            End If
        Else
            If dEndY < dStY Then
                dResult = 270
            Else
                dResult = 90
            End If
        End If
    Else
        'All other anlges
        dResult = Atn(dDeltaY / dDeltaX) * 180 / pi
        
        'Special Cases due to the nature of TAN
        If dDeltaX < 0 And dDeltaY > 0 Then
            dResult = 180 + dResult
        End If
        
        If dDeltaX < 0 And dDeltaY < 0 Then
            dResult = 180 + dResult
        End If
        
        If dDeltaX > 0 And dDeltaY < 0 Then
           dResult = 360 + dResult
        End If
    End If
        
    While dResult <= 0
        dResult = dResult + 360
    Wend
    
    GetElementAngle2D = Round(dResult, iDecimals)
    
    'Debug.Print "Angle: " & Round(dResult, iDecimals) & " DeltaX: " & Round(dDeltaX, iDecimals) & " DeltaY: " & Round(dDeltaY, iDecimals)
    
End Function


Public Sub AfterOpenPost(Config As PostConfigure) 'Read on AlphaCAM Startup & Post Selection

    Dim NewVal As Double
    
'    Events.BeforeCreateNC
    Functions.GetSettings
                
    With Config
                
        If sUnits = "IN" Then
            Config.FeedMax = 1500
            .ArcChordTolerance = 0.005
            .MaximumArcRadius = 3900
            .RapidXYSpeed = 1500
            .RapidZSpeed = 800
            .XYZNumberFormat.FiguresAfterPoint = 5
            .ArcCentreNumberFormat.FiguresAfterPoint = 5

        Else
            Config.FeedMax = 38100
            .ArcChordTolerance = 0.1
            .MaximumArcRadius = 100000
            .RapidXYSpeed = 38100
            .RapidZSpeed = 20320
            .XYZNumberFormat.FiguresAfterPoint = 3
            .ArcCentreNumberFormat.FiguresAfterPoint = 3

        End If
        
        
'        .XYZNumberFormat.FiguresAfterPoint = iDecimals
'        .ArcCentreNumberFormat.FiguresAfterPoint = iDecimals

        .XYZNumberFormat.Format = acamPostNumberFormat4LEADING
        .XYZNumberFormat.LeadingFigures = 1
    
        .ArcCentreNumberFormat.Format = acamPostNumberFormat4LEADING
        .ArcCentreNumberFormat.LeadingFigures = 1
    
        .SubroutineNumberFormat.Format = acamPostNumberFormat6INTEGER
        .SubroutineNumberFormat.LeadingFigures = 0
        .SubroutineNumberFormat.FiguresAfterPoint = 0
        .SubroutineStartNumber = 0
            
        .LineNumberFormat.Format = acamPostNumberFormat6INTEGER
        .LineNumberFormat.LeadingFigures = 0
        .LineNumberFormat.FiguresAfterPoint = 0
        .LineNumberIncrement = 1
    
        .AllowPositiveAndNegativeTilt = False
        .AllowManagedRapids = True
        .AllowOutputVisibleOnly = True
        
        .CAxisArcsAndLinesAsLines = False
        .CWSpindleRotation = "M3"
        .CCWSpindleRotation = "M4"
        .CoolantMist = "M8"
        .CoolantOff = ""
        .CoolantFlood = ""
        .CoolantThroughTool = ""
        
        .FeedNumberFormat.Format = acamPostNumberFormat6INTEGER
        .FeedNumberFormat.FiguresAfterPoint = 0
        .FeedNumberFormat.LeadingFigures = 0
        
        .FiveAxisOffsetFromPivotPointX = 0
        .FiveAxisOffsetFromPivotPointY = 0
        .FiveAxisProgramPivot = False
        .FiveAxisToolHolderLength = 0
        .FiveAxisToolMaxAngle = 135
        '.FiveAxisToolMaxAngleChange = 0.001
        
        .HelicalArcsAsLines = False
        .HorizontalMCCentre = True
        
        '.LimitArcs = acamPostLimitArcsNONE
        .LimitArcs = acamPostLimitArcsQUAD
        .LocalXorYAxis = acamPostLocalXorYAxisNONE
        
        .MCToolComp5Axis = False
        .MCToolCompAdjustInternalCorners = False
        .MCToolCompBlendPercent = 0
        .MCToolCompCancel = 0
        .MCToolCompLeft = 1
        .MCToolCompRight = 2
        .MCToolCompOnRapidApproach = True
        
        .ModalAbsoluteValues = "X Y Z A C F S U V W"
        
        .NeedPlusSigns = False
        
        .PlanarArcsAsLines = acamPostPlanarArcsAsLinesNONE
    
        
        .SelectWpToolOrder = acamPostSelectWpToolOrderBOTH
        
        .SpindleSpeedMax = 30000
        .SpindleSpeedNumberFormat.Format = acamPostNumberFormat6INTEGER
        .SpindleSpeedNumberFormat.FiguresAfterPoint = 0
        .SpindleSpeedNumberFormat.LeadingFigures = 0
        .SpindleSpeedRound = 100
        
        .SuppressComments = True
        
        .ToolChangeTime = 16
        .ToolNumberFormat.Format = acamPostNumberFormat7INTEGER_LEAD_0
        .ToolNumberFormat.FiguresAfterPoint = 0
        .ToolNumberFormat.LeadingFigures = 1
    
        .SelectWpToolOrder = acamPostSelectWpToolOrderTOOL_FIRST
    
    End With


End Sub

Public Sub OutputStop(p As PostData)
  
      p.Post "BEGIN MACRO"
      p.Post "   NAME=WAIT"
      p.Post "   PARAM,NAME=TYP,VALUE=stNT"
      p.Post "   PARAM,NAME=OG,VALUE=1"
      p.Post "   PARAM,NAME=RT,VALUE=0"
      p.Post "   PARAM,NAME=MR,VALUE=0"
      p.Post "   PARAM,NAME=UK,VALUE=YES"
      p.Post "   PARAM,NAME=SWS,VALUE=NO"
      p.Post "END MACRO"
      p.Post ""
   
End Sub

Public Function GetWorkplaneID(ByRef p As PostData) As String

    Dim wp          As WorkPlane
    Set wp = p.WorkPlane

    Dim swpatt      As String
    swpatt = ""
    
    Dim sRet        As String
    sRet = ""
    
    Dim dOriginX    As Double
    Dim dOriginY    As Double
    Dim dOriginZ    As Double
    Dim dTilt       As Double
    Dim dRotZ       As Double
    
    dOriginX = Round(p.Vars.LGX, iDecimals)
    dOriginY = Round(dY - p.Vars.LGY, iDecimals)
    dOriginZ = Round(-p.Vars.LGZ, iDecimals)
    dTilt = Round(90 - p.Vars.WTC, iDecimals)
    dRotZ = Round(270 - p.Vars.WAC, iDecimals)
    'Modulo 360
    Do While (dRotZ >= 360)
       dRotZ = dRotZ - 360
    Loop
    Do While (dRotZ < 0)
        dRotZ = dRotZ + 360
    Loop
    
    If wp.ToolPaths.Count > 0 And CInt(ReturnWP(wp, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)) < 0 Then
        If wp.Attribute("hhhBiesseWorkplaneID") = "NoID" Then
            iWPid = iWPid + 1
            
            sRet = VBA.Trim(CStr(iWPid))
            wp.Attribute("hhhBiesseWorkplaneID") = sRet
            
            p.Post "BEGIN MACRO"
            p.Post "   NAME=WFL"
            p.Post "   PARAM,NAME=ID,VALUE=" & VBA.Trim(wp.Attribute("hhhBiesseWorkplaneID")) 'ID
            p.Post "   PARAM,NAME=X,VALUE=" & VBA.Trim(CStr(dOriginX)) 'X-origin
            p.Post "   PARAM,NAME=Y,VALUE=" & VBA.Trim(CStr(dOriginY)) 'Y-origin
            p.Post "   PARAM,NAME=Z,VALUE=" & VBA.Trim(CStr(dOriginZ)) 'Z-origin
            p.Post "   PARAM,NAME=AZ,VALUE=" & VBA.Trim(CStr(dTilt)) 'Tilt
            p.Post "   PARAM,NAME=AR,VALUE=" & VBA.Trim(CStr(dRotZ)) 'Rot in Z-Axis
            p.Post "   PARAM,NAME=L,VALUE=" & VBA.Trim(CStr(dX)) '"Length" of the display face
            p.Post "   PARAM,NAME=H,VALUE=lpz"
            p.Post "   PARAM,NAME=VRT,VALUE=0" 'Flag which forces a VERTICAL plane, must be disabled
            p.Post "   PARAM,NAME=VF,VALUE=1" 'Virtual Face (If ON,  Virtual Simulator will not show it)
            p.Post "   PARAM,NAME=AFL,VALUE=0" 'Auto-Length of the Workplane
            p.Post "   PARAM,NAME=AFH,VALUE=0" 'Auto-Height Based on the Thickness of the Piece (like a Work Volume face, I guess)
            p.Post "   PARAM,NAME=UCS,VALUE=1" 'System used for the Reference Corner
            p.Post "   PARAM,NAME=RV,VALUE=0"
            p.Post "   PARAM,NAME=FRC,VALUE=2" 'Reference Corner
            p.Post "END MACRO"
            p.Post ""
        Else
            swpatt = wp.Attribute("hhhBiesseWorkplaneID")
            If swpatt <> "" Then
                sRet = swpatt
            End If
        End If
    End If
    
    GetWorkplaneID = sRet
    
End Function

Private Sub ClearWpAtts(ByRef p As PostData)

    Dim wp As WorkPlane
    
    For Each wp In App.ActiveDrawing.WorkPlanes
        If wp.ToolPaths.Count > 0 Then
            wp.Attribute("hhhBiesseWorkplaneID") = "NoID"
        End If
    Next wp
End Sub

Public Function IsNextEleLastAndZeroLength(ByRef TheL As Element) As Boolean
    
    IsNextEleLastAndZeroLength = False
    
    Dim NextEle As Element
    
    Set NextEle = TheL.GetNext
    
    '1. Check to see if the NEXT element is Zero length
    If NextEle.Length > 0.0001 Then
        'Its Longer than Zero. Quit
        Exit Function
    End If
    
    '2. OK, the NEXT element IS Zero Length. Is the Next Element the LAST ONE of the Collection?
    Dim LastEle As Element
    Set LastEle = TheL.Path.Elements(TheL.Path.Elements.Count)
    If LastEle.IsSame(NextEle) Then
        IsNextEleLastAndZeroLength = True
    End If
    
End Function

Private Sub LabelOverlaps()
    
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

Private Function GetOverlapLength(ByRef tp As Path) As Double
    
    Dim eleOverlap  As Element
    Dim dRet        As Double
    dRet = 0
        
    Set eleOverlap = tp.GetLastElem
    
    While eleOverlap.IsRapid
        Set eleOverlap = eleOverlap.GetPrevious
    Wend
    
    While eleOverlap.LeadOut
        Set eleOverlap = eleOverlap.GetPrevious
    Wend
    
    If eleOverlap.Attribute("_hhhOverlapElement") = 1 Then
        dRet = eleOverlap.Length2D
    End If
    
    GetOverlapLength = dRet
End Function

Private Function HasLeadIn(ByRef tp As Path) As Boolean
    
    Dim bRet As Boolean
    bRet = False
    
    Dim Ele As Element
    Set Ele = tp.GetFirstElem
    
    While Ele.IsRapid
        Set Ele = Ele.GetNext
    Wend
    
    'Lead-in will ALWYAS be the first non-rapid element!
    If Ele.LeadIn Then
        bRet = True
    End If
    
    HasLeadIn = bRet
End Function

Private Function FindArcLeadAngle(ByRef EleArc As Element) As Double
    
    Dim tp          As Path
    Dim tp2          As Path
    
    Dim Ele1        As Element
    Dim Ele2        As Element
    
    Dim dAngle      As Double

    dAngle = 0
    
    Set tp = App.ActiveDrawing.Create2DLine(EleArc.CenterXL, EleArc.CenterYL, EleArc.StartXL, EleArc.StartYL)
    tp.SetLayer lyrLeadVectors
    
    Set Ele1 = tp.GetFirstElem
    
    Set tp2 = App.ActiveDrawing.Create2DLine(EleArc.CenterXL, EleArc.CenterYL, EleArc.EndXL, EleArc.EndYL)
    tp2.SetLayer lyrLeadVectors
    
    Set Ele2 = tp2.GetFirstElem
    
    dAngle = -Ele1.AngleToElement(Ele2)
    
    FindArcLeadAngle = dAngle

End Function

Private Sub MakeGenericBore(ByRef p As PostData)

    If Not MachiningAllowed(p) Then
        MsgBox "Error! Machining not permitted! " & Functions.RetMachineType(p)
        End
    End If

    Dim dPostNX     As Double
    Dim dPostNY     As Double
    Dim dPostNZ     As Double
    Dim dPostCX     As Double
    Dim dPostCY     As Double
    Dim dPostCZ     As Double
    Dim sFace       As String
    
    If CInt(ReturnWP(p.WorkPlane, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)) > -1 Then
        sFace = ReturnWP(p.WorkPlane, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)
    Else
        sFace = GetWorkplaneID(p)
    End If
    
    'BiesseWorks is VERY civilised. The code for ANY face is the same.
    'You literally just change the Face and all machining applies from one face to another.
    'You can even copy a milling op to another face!
    
    'Global on Top Face
    If CInt(ReturnWP(p.WorkPlane, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)) = 0 Then
        dPostNX = p.Vars.GAX + dFaceOffX
        dPostNY = p.Vars.GAY + dFaceOffY
        dPostNZ = -p.Vars.ZB + dFaceOffZ
        dPostCX = p.Vars.GCX + dFaceOffX
        dPostCY = p.Vars.GCY + dFaceOffY
        dPostCZ = -p.Vars.ZB + dFaceOffZ
    Else
        dPostNX = p.Vars.AX + dFaceOffX
        dPostNY = p.Vars.AY + dFaceOffY
        dPostNZ = -p.Vars.ZB + dFaceOffZ
        dPostCX = p.Vars.CX + dFaceOffX
        dPostCY = p.Vars.CY + dFaceOffY
        dPostCZ = -p.Vars.ZB + dFaceOffZ
    End If
    
    p.Post "BEGIN MACRO"
    p.Post "   NAME=BG"
    p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & Replace(VBA.Trim(RetMachineType(p) & " " & p.Path.Name), " ", "_") & Chr(34) 'Toolpath name
    p.Post "   PARAM,NAME=SIDE,VALUE=" & VBA.Trim(sFace)
    p.Post "   PARAM,NAME=CRN,VALUE=" & Chr(34) & "2" & Chr(34) 'Always use the bottom-left reference on the face
    p.Post "   PARAM,NAME=Z,VALUE=0" 'Channel' Supposed to create an offset from the top face. See Doc.
    p.Post "   PARAM,NAME=X,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNX, iDecimals))) 'Depth of Cut
    p.Post "   PARAM,NAME=Y,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNY, iDecimals))) 'Depth of Cut
    p.Post "   PARAM,NAME=DP,VALUE=" & VBA.Trim(VBA.CStr(Round(dPostNZ, iDecimals))) 'Depth of Cut
    'Feed Data, depending on existance of leads
    If Not bUseCNCFeeds Then
        p.Post "   PARAM,NAME=DSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, iDecimals)))  'G0/B Feed Rate
        p.Post "   PARAM,NAME=RSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.S, iDecimals)))  'Speed
        p.Post "   PARAM,NAME=IOS,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.f, iDecimals)))  'Lead Speed
        'Working Feed
        p.Post "   PARAM,NAME=WSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Tool.FixedFeed, iDecimals)))
    End If
    
    If p.Path.GetMillData.ProcessType2 = acamProcessBORE Or p.Path.GetMillData.ProcessType2 = acamProcessDRILL Then
        'Tool Type. Controls hole with Main Spindle vs Drilling Unit.
        If p.Tool.Type = acamToolDRILL Or p.Path.Attribute("AcamUSrg_Mdrill") = "1" Then
            p.Post "   PARAM,NAME=OPT,VALUE=YES" 'Optimisable?
            p.Post "   PARAM,NAME=TCL,VALUE=0"
            p.Post "   PARAM,NAME=TNM,VALUE=" & Chr(34) & Chr(34)
            p.Post "   PARAM,NAME=DIA,VALUE=" & Round(p.Vars.TD, iDecimals)
            
            If p.Vars.OFS = 1 Then
                p.Post "   PARAM,NAME=TTP,VALUE=1"  'to specify a lancia bit 9-26-2018 FRW
            Else
            End If
            
        Else
            p.Post "   PARAM,NAME=OPT,VALUE=NO" 'Optimisable?
            p.Post "   PARAM,NAME=TNM,VALUE=" & Chr(34) & VBA.Trim(p.Vars.TNM) & Chr(34) 'Tool Name
            p.Post "   PARAM,NAME=TCL,VALUE=1"
        End If
    End If
    
    'Support for blower
    If p.Path.GetMillData.Coolant = acamCoolNONE Then
        p.Post "   PARAM,NAME=BFC,VALUE=NO"
    Else
        p.Post "   PARAM,NAME=BFC,VALUE=YES"
    End If
    
    'Support for dust collection hood
    If p.Tool.TPD(1) <> "" Then
        'User-defined
        p.Post "   PARAM,NAME=SHP,VALUE=" & p.Tool.TPD(1)
    Else
        'Automatic
        p.Post "   PARAM,NAME=SHP,VALUE=0"
    End If
    
    p.Post "END MACRO"
    p.Post ""
    

End Sub

Private Sub OutputSimpleSaw(ByRef p As PostData)
    
    Dim sFace As String
        sFace = ""
        
    Dim sStartX     As String
    Dim sStartY     As String
    Dim sEndX       As String
    Dim sEndY       As String
    Dim sDepth      As String
    Dim sPost       As String
        
    Dim Ele         As Element
    
    'Verify if THIS saw path has been executed already or not.
    'sSawPaths will store the names of all the paths which have been processed.
    If InStr(sSawPaths, p.Path.Name) > 0 Then Exit Sub
    
    If CInt(ReturnWP(p.WorkPlane, p.Vars.LGX, p.Vars.LGY, p.Vars.LGZ)) <> 0 Then
        MsgBox "Error! Simple Sawing is only valid in Flatland or a parallel plane to it!", vbCritical
        p.Post "$EXIT"
    End If
    
    'Global on top face
    Set Ele = p.Path.Elements(1)
    While Ele.IsRapid
        Set Ele = Ele.GetNext
    Wend
    
    sStartX = VBA.Trim(VBA.CStr(Round(Ele.StartXG, iDecimals)))
    sStartY = VBA.Trim(VBA.CStr(Round(Ele.StartYG, iDecimals)))
    
    Set Ele = p.Path.Elements(p.Path.Elements.Count)
    While Ele.IsRapid
        Set Ele = Ele.GetPrevious
    Wend
    
    sEndX = VBA.Trim(VBA.CStr(Round(Ele.EndXG, iDecimals)))
    sEndY = VBA.Trim(VBA.CStr(Round(Ele.EndYG, iDecimals)))
    sDepth = VBA.Trim(VBA.CStr(Round(-p.Path.GetMillData.DepthOfCut, iDecimals)))
        
    'Cut Setup
    p.Post "BEGIN MACRO"
    p.Post "   NAME=CUT_G"
    p.Post "   PARAM,NAME=SIDE,VALUE=0" 'Only valid in Flatland or Parallel to it.
    p.Post "   PARAM,NAME=ID,VALUE=" & Chr(34) & Replace(("SimpleSawCut" & " " & p.Path.Name), " ", "_") & Chr(34) 'Toolpath name
    p.Post "   PARAM,NAME=CRN,VALUE=" & Chr(34) & "2" & Chr(34) 'Always use the bottom-left reference on the face
    p.Post "   PARAM,NAME=X,VALUE=" & sStartX 'Start Coordinate for the cut
    p.Post "   PARAM,NAME=Y,VALUE=" & sStartY 'Start Coordinate for the cut
    p.Post "   PARAM,NAME=XE,VALUE=" & sEndX 'End Coordinate for the cut
    p.Post "   PARAM,NAME=YE,VALUE=" & sEndY 'End Coordinate for the cut
    p.Post "   PARAM,NAME=Z,VALUE=0" 'Start Coordinate for the cut. Depth is in DP.
    p.Post "   PARAM,NAME=DP,VALUE=" & sDepth 'Depth of Cut
    p.Post "   PARAM,NAME=TYP,VALUE=cutXY"
    'p.Post "   PARAM,NAME=DP,VALUE=" & sLength 'Length
    p.Post "   PARAM,NAME=OPT,VALUE=NO" 'Optimisable?
    'p.Post "   PARAM,NAME=TH,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Path.GetTool.Diameter, iDecimals))) 'Blade THICKNESS. Will be left blank on purpose to test if the blade works by NAME
    p.Post "   PARAM,NAME=RV,VALUE=NO" 'Reverse direction of cut?
    p.Post "   PARAM,NAME=TTK,VALUE=0" 'Total Cut Thickness. No fucking clue. Came from Russell.
    p.Post "   PARAM,NAME=OVM,VALUE=0" 'Overmaterial
    p.Post "   PARAM,NAME=TOS,VALUE=NO" '"Chanel" Depth
    p.Post "   PARAM,NAME=VTR,VALUE=NO" 'Vertical Runs
    p.Post "   PARAM,NAME=GIP,VALUE=YES" 'Complicated. See page 405 of Biesse Manual. I think it sets the TOP surface as the zero ref.
    p.Post "   PARAM,NAME=TNM,VALUE=" & Chr(34) & VBA.Trim(p.Vars.TNM) & Chr(34) 'Tool Name. THIS is what will select the tool from the catalog.
    p.Post "   PARAM,NAME=TTP,VALUE=200" 'Tool type forced to CUTTING (page 641)
    p.Post "   PARAM,NAME=TCL,VALUE=2" 'Tool Class forced to C_Cutting (page 641)
    'Feed Data, depending on existance of leads
    If Not bUseCNCFeeds Then
        If p.Element.LeadIn Then
            p.Post "   PARAM,NAME=RSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.S, iDecimals)))  'Speed
            p.Post "   PARAM,NAME=IOS,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.FD, iDecimals)))  'Lead Speed
        Else
            p.Post "   PARAM,NAME=RSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.S, iDecimals)))  'Speed
            p.Post "   PARAM,NAME=IOS,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.FD, iDecimals)))  'Lead Speed
        End If
        'Working Feed
        p.Post "   PARAM,NAME=WSP,VALUE=" & VBA.Trim(VBA.CStr(Round(p.Vars.FC, iDecimals)))
    End If
    p.Post "   PARAM,NAME=SPI,VALUE=""" 'Selected Spindle page 434
    'Support for blower
    If p.Path.GetMillData.Coolant = acamCoolNONE Then
        p.Post "   PARAM,NAME=BFC,VALUE=NO"
    Else
        p.Post "   PARAM,NAME=BFC,VALUE=YES"
    End If
    'Support for dust collection hood
    If p.Tool.TPD(1) <> "" Then
        'User-defined
        p.Post "   PARAM,NAME=SHP,VALUE=" & p.Tool.TPD(1)
    Else
        'Automatic
        p.Post "   PARAM,NAME=SHP,VALUE=0"
    End If
    p.Post "   PARAM,NAME=BRC,VALUE=0"  'Tool Comp is ALWAYS ON CENTER. Global ONLY.
    p.Post "   PARAM,NAME=BDR,VALUE=NO" 'Bi-Directional simple cuts NOT supported!
    p.Post "   PARAM,NAME=PRV,VALUE=YES" '"Possible Reverse of Cut Direction". Overrides Blade to cut in the default direction!
    p.Post "   PARAM,NAME=NRV,VALUE=NO" ' Undocumented)
    p.Post "   PARAM,NAME=DIN,VALUE=0"     'Lead Data (mostly ignored
    p.Post "   PARAM,NAME=DOU,VALUE=0"     'Lead Data (mostly ignored
    p.Post "   PARAM,NAME=CRC,VALUE=0"     'Lead Data (mostly ignored
    p.Post "   PARAM,NAME=DSP,VALUE=0"     'Lead Data (mostly ignored
    p.Post "   PARAM,NAME=CEN,VALUE="""     'Lead Data (mostly ignored
    p.Post "   PARAM,NAME=AGG,VALUE="""     'Agg ID
    p.Post "   PARAM,NAME=LAY,VALUE=" & Chr(34) & "SimpleSawCut" & Chr(34)
    p.Post "   PARAM,NAME=DVR,VALUE=0" ' Undocumented
    p.Post "   PARAM,NAME=ETB,VALUE=NO" ' Undocumented
    p.Post "   PARAM,NAME=KDT,VALUE=NO" ' Undocumented
'    p.Post "   PARAM,NAME=NRV,VALUE=YES" ' Compulsory Flip Direction. ONLY for QFC.
    p.Post "END MACRO"
    p.Post ""
        
    'Record Path
    sSawPaths = sSawPaths & " " & p.Path.Name
    
End Sub
