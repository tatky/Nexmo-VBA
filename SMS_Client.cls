VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SMS_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' SMS Client Class
' for NEXMO SMS API
' @author tat
'--------------------------------------------------------------------------------

Option Explicit
Const BASE_URL = "https://rest.nexmo.com/sms/xml"
Const QUERY_STRING_BASE = "//mt-submission-response/messages/message/"

Private api_key As String
Private api_secret As String

Private smsClient As Object 'MSXML2.XMLHTTP

Private defaultRequestParamCollection As Collection

' Constructor
' @author tat
Private Sub Class_Initialize()
    Set smsClient = CreateObject("Msxml2.XMLHTTP.6.0")
End Sub

' Destructor
' @author tat
Private Sub Class_Terminate()
    Set smsClient = Nothing
End Sub

' Setter for api_key
' @author tat
Property Let letAPI_KEY(key As String)
    api_key = key
End Property

' Setter for api_secret
' @author tat
Property Let letAPI_SECRET(secret As String)
    api_secret = secret
End Property

' Setter for defaultRequestParamCollection
' @author tat
Property Set setDefaultRequestparam(ByRef requestParamCollection As Collection)
    Set defaultRequestParamCollection = requestParamCollection
End Property

' send Function
' @param from_(Ref) sms from string
' @param to_(Ref) sms to string
' @param text_(Ref) sms text string
' @author tat
Public Function send(ByRef from_ As String, ByRef to_ As String, ByRef text_ As String) As Collection
    If api_key = "" Or api_secret = "" Then
        MsgBox "require api_key and api_seqret", vbCritical, "SMS Client"
        Err.Raise 5
    End If
    Call smsClient.Open("POST", BASE_URL, False)
    Call smsClient.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    Const EQ As String = "="
    Const AMP As String = "&"
    Dim postData As Variant
    postData = SMS_RequestConst.KEY_API_KEY & EQ & api_key
    postData = postData & AMP & SMS_RequestConst.KEY_API_SECRET & EQ & api_secret
    postData = postData & AMP & SMS_RequestConst.KEY_FROM & EQ & from_
    postData = postData & AMP & SMS_RequestConst.KEY_TO & EQ & to_
    postData = postData & AMP & SMS_RequestConst.KEY_TEXT & EQ & text_
    Dim key As Variant
    On Error Resume Next
    For Each key In SMS_RequestConst.keys
        postData = postData & AMP & key & EQ & defaultRequestParamCollection.item(key)
    Next key
    On Error GoTo 0
    Call smsClient.send(postData)
    
    Dim ret As New Collection
    ret.Add smsClient.STATUS, SMS_ResponseConst.STATUS_CODE
    ret.Add smsClient.statusText, SMS_ResponseConst.STATUS_TEXT
    ret.Add smsClient.responseText, SMS_ResponseConst.RESPONSE_TEXT
    
    Dim dd As Object 'MSXML2.DOMDocument
    Set dd = CreateObject("MSXML2.DOMDocument")
    Dim loadRet As Boolean
    loadRet = dd.LoadXML(smsClient.responseText)
    If loadRet Then
        Dim key2 As Variant
        Dim node As Object 'IXMLDOMNode
        For Each key2 In SMS_ResponseConst.keys
            Set node = dd.SelectSingleNode(QUERY_STRING_BASE & key2)
            If Not node Is Nothing Then
                ret.Add node.text, key2
            End If
        Next key2
    End If
    Set send = ret
End Function

