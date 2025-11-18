Attribute VB_Name = "CMD"
Option Explicit
 
DefStr S
DefLng N
DefBool B
DefVar V
 
' OFN constants.
Const OFN_ALLOWMULTISELECT   As Long = &H200
Const OFN_CREATEPROMPT       As Long = &H2000
Const OFN_EXPLORER           As Long = &H80000
Const OFN_EXTENSIONDIFFERENT As Long = &H400
Const OFN_FILEMUSTEXIST      As Long = &H1000
Const OFN_HIDEREADONLY       As Long = &H4
Const OFN_LONGNAMES          As Long = &H200000
Const OFN_NOCHANGEDIR        As Long = &H8
Const OFN_NODEREFERENCELINKS As Long = &H100000
Const OFN_OVERWRITEPROMPT    As Long = &H2
Const OFN_PATHMUSTEXIST      As Long = &H800
Const OFN_READONLY           As Long = &H1
 
' The maximum length of a single file path.
Const MAX_PATH As Long = 256
' This MAX_BUFFER value allows you to select approx.
' 500 files with an average length of 25 characters.
' Change this value as needed.
Const MAX_BUFFER As Long = 50 * MAX_PATH
' String constants:
Const sBackSlash As String = "\"
'Const sPipe As String = "|"
 
#If VBA7 Then
    ' API functions to use the Windows common dialog boxes.
    Private Type OPENFILENAME
        lStructSize         As Long
        hWndOwner           As LongPtr
        hInstance           As LongPtr
        lpstrFilter         As String
        lpstrCustomFilter   As String
        nMaxCustFilter      As Long
        nFilterIndex        As Long
        lpstrFile           As String
        nMaxFile            As Long
        lpstrFileTitle      As String
        nMaxFileTitle       As Long
        lpstrInitialDir     As String
        lpstrTitle          As String
        flags               As Long
        nFileOffset         As Integer
        nFileExtension      As Integer
        lpstrDefExt         As String
        lCustData           As Long
        lpfnHook            As LongPtr
        lpTemplateName      As String
    End Type
    
    Public Type BrowseInfo
        hWndOwner As LongPtr
        pIDLRoot As LongPtr
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfnCallback As LongPtr
        lParam As LongPtr
        iImage As Long
    End Type

    
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As OPENFILENAME) As Boolean
    Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As OPENFILENAME) As Boolean
    Private Declare PtrSafe Function GetActiveWindow Lib "User32.dll" () As LongPtr
    
#Else

    ' API functions to use the Windows common dialog boxes.
    Type OPENFILENAME
      lStructSize As Long
      hWndOwner As Long
      hInstance As Long
      lpstrFilter As String
      lpstrCustomFilter As String
      nMaxCustFilter As Long
      nFilterIndex As Long
      lpstrFile As String
      nMaxFile As Long
      lpstrFileTitle As String
      nMaxFileTitle As Long
      lpstrInitialDir As String
      lpstrTitle As String
      flags As Long
      nFileOffset As Integer
      nFileExtension As Integer
      lpstrDefExt As String
      lCustData As Long
      lpfnHook As Long
      lpTemplateName As String
    End Type
    
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As OPENFILENAME) As Boolean
    Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As OPENFILENAME) As Boolean
    Private Declare Function GetActiveWindow Lib "User32.dll" () As Long

#End If

' Private variables.
Private OFN As OPENFILENAME
Private colFileTitles As New Collection
Private colFileNames As New Collection
Private sFullName
Private sFileTitle
Private sPath
Private sExtension
 
' Public enumeration variable.
Public Enum XFlags
  PathMustExist = OFN_PATHMUSTEXIST
  FileMustExist = OFN_FILEMUSTEXIST
  PromptToCreateFile = OFN_CREATEPROMPT
End Enum
 
Property Let AllowMultiSelect(bFlag)
  SetFlag OFN_ALLOWMULTISELECT, bFlag
End Property
 
Property Let DialogTitle(sCaption)
  OFN.lpstrTitle = sCaption
End Property
 
Property Let Filter(vFilter)
  'If IsArray(vFilter) Then vFilter = Join(vFilter, vbNullChar)
  'OFN.lpstrFilter = Replace(vFilter, sPipe, vbNullChar) & vbNullChar
    OFN.lpstrFilter = vFilter
End Property
 
Property Get Filter()
  With OFN
  '  If .nFilterIndex Then
  '    Dim sTemp()
  '    sTemp = Split(.lpstrFilter, vbNullChar)
  '    Filter = sTemp(.nFilterIndex * 2 - 2) & sPipe & sTemp(.nFilterIndex * 2 - 1)
  '  End If
  
     Filter = .nFilterIndex
  
  End With
End Property
 
Property Let FilterIndex(nIndex)
  OFN.nFilterIndex = nIndex
End Property
 
Property Get FilterIndex() As Long
  FilterIndex = OFN.nFilterIndex
End Property
 
Property Let RestoreCurDir(bFlag)
  SetFlag OFN_NOCHANGEDIR, bFlag
End Property
 
Property Let ExistFlags(nFlags As XFlags)
  OFN.flags = OFN.flags Or nFlags
End Property
 
Property Let CheckBoxVisible(bFlag)
  SetFlag OFN_HIDEREADONLY, Not bFlag
End Property
 
Property Let CheckBoxSelected(bFlag)
  SetFlag OFN_READONLY, bFlag
End Property
 
Property Get CheckBoxSelected() As Boolean
  CheckBoxSelected = OFN.flags And OFN_READONLY
End Property
 
Property Let FileName(sFileName)
  If Len(sFileName) <= MAX_PATH Then OFN.lpstrFile = sFileName
End Property
 
Property Get FileName() As String
  FileName = sFullName
End Property
 
Property Get FileNames() As Collection
  Set FileNames = colFileNames
End Property
 
Property Get FileTitle() As String
  FileTitle = sFileTitle
End Property
 
Property Get FileTitles() As Collection
  Set FileTitles = colFileTitles
End Property
 
Property Let Directory(sInitDir)
  OFN.lpstrInitialDir = sInitDir
End Property
 
Property Get Directory() As String
  Directory = sPath
End Property
 
Property Let Extension(sDefExt)
  OFN.lpstrDefExt = LCase$(Left$(Replace(sDefExt, ".", vbNullString), 3))
End Property
 
Property Get Extension() As String
  Extension = sExtension
End Property
 
Function ShowOpen() As Boolean
  ShowOpen = Show(True)
End Function
 
Function ShowSave() As Boolean
  ' Set or clear appropriate flags for Save As dialog.
  SetFlag OFN_ALLOWMULTISELECT, False
  SetFlag OFN_PATHMUSTEXIST, True
  SetFlag OFN_OVERWRITEPROMPT, True
  ShowSave = Show(False)
End Function
 
Private Function Show(bOpen)
  With OFN
    .lStructSize = LenB(OFN)
    ' Could be zero if no owner is required.
    .hWndOwner = GetActiveWindow
    ' If the RO checkbox must be checked, we should also
    ' display it.
    If .flags And OFN_READONLY Then _
      SetFlag OFN_HIDEREADONLY, False
    ' Create large buffer if multiple file selection
    ' is allowed.
    .nMaxFile = IIf(.flags And OFN_ALLOWMULTISELECT, _
      MAX_BUFFER + 1, MAX_PATH + 1)
    .nMaxFileTitle = MAX_PATH + 1
    ' Initialize the buffers.
    .lpstrFile = .lpstrFile & String$( _
      .nMaxFile - 1 - Len(.lpstrFile), 0)
    .lpstrFileTitle = String$(.nMaxFileTitle - 1, 0)
 
    ' Display the appropriate dialog.
    If bOpen Then
      Show = GetOpenFileName(OFN)
    Else
      Show = GetSaveFileName(OFN)
    End If
 
    If Show Then
      ' Remove trailing null characters.
      Dim nDoubleNullPos
      nDoubleNullPos = InStr(.lpstrFile & vbNullChar, String$(2, 0))
      nDoubleNullPos = InStr(.lpstrFile, vbNullChar)
      If nDoubleNullPos Then
        ' Get the file name including the path name.
        sFullName = Left$(.lpstrFile, nDoubleNullPos - 1)
        If .nFileExtension = 0 Then
            sFullName = sFullName & "." & LCase$(Left$((Right$(.lpstrFilter, 4)), 3))
            sExtension = LCase$(Left$((Right$(.lpstrFilter, 4)), 3))
            .lpstrDefExt = sExtension
        End If
        ' Get the file name without the path name.
        sFileTitle = Left$(.lpstrFileTitle, _
          InStr(.lpstrFileTitle, vbNullChar) - 1)
        ' Get the path name.
        sPath = Left$(sFullName, .nFileOffset - 1)
        ' Get the extension.
        If .nFileExtension Then
          sExtension = Mid$(sFullName, .nFileExtension + 1)
        End If
        ' If sFileTitle is a string,
        ' we have a single selection.
        If Len(sFileTitle) Then
          ' Add to the collections.
          colFileTitles.Add _
            Mid$(sFullName, .nFileOffset + 1)
          colFileNames.Add sFullName
        Else  ' Tear multiple selection apart.
          Dim sTemp(), nCount
          sTemp = Split(sFullName, vbNullChar)
          ' If array contains no elements,
          ' UBound returns -1.
          If UBound(sTemp) > LBound(sTemp) Then
            ' We have more than one array element!
            ' Remove backslash if sPath is the root folder.
            If Len(sPath) = 3 Then _
              sPath = Left$(sPath, 2)
            ' Loop through the array, and create the
            ' collections; skip the first element
            ' (containing the path name), so start the
            ' counter at 1, not at 0.
            For nCount = 1 To UBound(sTemp)
              colFileTitles.Add sTemp(nCount)
              ' If the string already contains a backslash,
              ' the user must have selected a shortcut
              ' file, so we don't add the path.
               colFileNames.Add IIf(InStr(sTemp(nCount), sBackSlash), sTemp(nCount), sPath & sBackSlash & sTemp(nCount))
            Next
            ' Clear this variable.
            sFullName = vbNullString
          End If
        End If
        ' Add backslash if sPath is the root folder.
        If Len(sPath) = 2 Then _
          sPath = sPath & sBackSlash
      End If
    End If
  End With
End Function
 
Private Sub SetFlag(nValue, bTrue)
  ' Wrapper routine to set or clear bit flags.
  With OFN
    If bTrue Then
      .flags = .flags Or nValue
    Else
      .flags = .flags And Not nValue
    End If
  End With
End Sub
 
Private Sub Class_Initialize()
  ' This routine runs when the object is created.
  OFN.flags = OFN.flags Or OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
End Sub
