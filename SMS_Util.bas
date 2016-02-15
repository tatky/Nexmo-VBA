Attribute VB_Name = "SMS_Util"
'--------------------------------------------------------------------------------
' SMS Util Module
' @author tat
'--------------------------------------------------------------------------------

Option Explicit

' encode to Utf8
' @param strSource(Ref) encode source String
' @author tat
Public Function encodeUtf8(ByRef strSource As String) As String
    Dim scriptControl As Object
    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "Jscript"
    encodeUtf8 = scriptControl.CodeObject.encodeURIComponent(strSource)
    Set scriptControl = Nothing
End Function

' Add CountryCode to localTelNo
' @param countryCode(Ref) countryNo
' @param telNo(Ref) localTelNo
' @author tat
Public Function getE164telNo(ByRef countryCode As String, ByRef telNo As String) As String
    Dim toTelNo As String
    toTelNo = telNo
    If Left(telNo, 1) = "0" Then
        toTelNo = Mid(telNo, 2)
    End If
    getE164telNo = countryCode & toTelNo
End Function
