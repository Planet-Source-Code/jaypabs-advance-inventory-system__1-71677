Attribute VB_Name = "modKurimaw"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function GetINI(strMain As String, strSub As String) As String
    Dim strBuffer As String
    Dim lngLen As Long
    Dim lngRet As Long
    
    strBuffer = Space(100)
    lngLen = Len(strBuffer)
    lngRet = GetPrivateProfileString(strMain, strSub, vbNullString, strBuffer, lngLen, App.Path & "\VTM.txt")
    GetINI = Left(strBuffer, lngRet)
End Function

Public Sub SetINI(strMain As String, strSub As String, strvalue As String)
    WritePrivateProfileString strMain, strSub, strvalue, App.Path & "\VTM.txt"
End Sub

Public Function LocationExist(ByVal Route As String, ByVal City As String) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = "SELECT Locations.LocationID " _
            & "FROM Routes RIGHT JOIN (Cities RIGHT JOIN Locations ON Cities.CityID = Locations.CityID) ON Routes.RouteID = Locations.RouteID " _
            & "WHERE (((Routes.Route)='" & Replace(Route, "'", "''") & "') AND " _
            & "((Cities.City)='" & Replace(City, "'", "''") & "'))"
    
    rs.Open sql, CN, adOpenDynamic, adLockOptimistic
    
    If Not rs.EOF Then
        LocationExist = True
    Else
        LocationExist = False
    End If
    
    rs.Close
    Set rs = Nothing
End Function


