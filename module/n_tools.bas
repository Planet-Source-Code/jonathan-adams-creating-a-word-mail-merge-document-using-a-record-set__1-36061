Attribute VB_Name = "n_tools"
Option Explicit

Declare Function apiGetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
 
Public Function get_tmp_file_name() As String
    On Error Resume Next
    Dim s_check_path As String
    Dim s_computer_name As String
 
    s_check_path = App.Path
    If Right(s_check_path, 1) <> "\" Then
        s_check_path = s_check_path & "\"
    End If
 
    s_computer_name = GetComputerName()
 
    s_check_path = s_check_path & s_computer_name
 
    get_tmp_file_name = s_check_path
get_tmp_html_file_name_Exit:
 

End Function
 
 
Function GetComputerName()
 
On Error GoTo Err_GetComputerName
 
    Dim UserString As String * 50 'needs to be bigger than max allowed name
    Dim StrLen As Long            'however long that is !
 
    StrLen = 50
    Call apiGetComputerName(UserString, StrLen)
    GetComputerName = Left(UserString, StrLen)
 
Exit_GetComputerName:
  Exit Function
 
Err_GetComputerName:
    Err.Raise Err.Number, "n_tools.Err_GetComputerName", Err.Description
    Resume Exit_GetComputerName
 
End Function

