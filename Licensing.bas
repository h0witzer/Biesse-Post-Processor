Attribute VB_Name = "Licensing"
Option Explicit
Option Private Module

Public sLicName As String
Public sCurrName As String

'Written by Hector Henry .......
Public Function GetLicense() As Boolean
    GetLicense = False
    
    sLicName = ""
    
    '########## RESELLERS  ##############
    sLicName = " Vero Software US Direct Sales"
    sLicName = sLicName & " Planit Solutions Inc"
    sLicName = sLicName & " BIESSE AMERICA"
    sLicName = sLicName & " WB Systech LLC"
    sLicName = sLicName & " Vectorline"
    sLicName = sLicName & " CAM Solutions (NZ) Limited"
    sLicName = sLicName & " CAM Solutions Ltd"
    sLicName = sLicName & " Soft 2 Cam (Paul Corey)" '
'    sLicName = sLicName & " " '
'    sLicName = sLicName & " " '
'    sLicName = sLicName & " " '
'    sLicName = sLicName & " " '
   
    '########## CUTSTOMERS  ##########################
    sLicName = sLicName & sCustomerName
'    sLicName = sLicName & " " '



    sLicName = UCase(sLicName)
    sCurrName = UCase(App.License.GetCustomerName)
    
    If InStr(sLicName, sCurrName) > 0 Then
        GetLicense = True
    End If
    
End Function
Public Function PstExpirey() As Boolean
        
        Dim varExpireDate As Variant: varExpireDate = DateValue(strExpiryDate)
        Dim varCurrentDate As Variant: varCurrentDate = Date
        
        If varCurrentDate > varExpireDate Then
            PstExpirey = False
        Else
           PstExpirey = True
        End If

End Function
