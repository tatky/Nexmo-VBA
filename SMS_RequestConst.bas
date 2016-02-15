Attribute VB_Name = "SMS_RequestConst"
'--------------------------------------------------------------------------------
' SMS Response Const
' @author tat
'--------------------------------------------------------------------------------

Public Const KEY_FROM As String = "from"
Public Const KEY_TO As String = "to"
Public Const KEY_TYPE As String = "type"
Public Const KEY_TEXT As String = "text"
Public Const KEY_STATUS_REPORT_REQ As String = "status-report-req"
Public Const KEY_CLIENT_REF As String = "client-ref"
Public Const KEY_VCARD As String = "vcard"
Public Const KEY_VCAL As String = "vcal"
Public Const KEY_TTL As String = "ttl"
Public Const KEY_CALLBACK As String = "callback"
Public Const KEY_MESSAGE_CLASS As String = "message-class"
Public Const KEY_UDH As String = "udh"
Public Const KEY_PROTOCOL_ID As String = "protocol-id"
Public Const KEY_BODY As String = "body"
Public Const KEY_TITLE As String = "title"
Public Const KEY_URL As String = "url"
Public Const KEY_VALIDITY As String = "validity"
Public Const KEY_API_KEY As String = "api_key"
Public Const KEY_API_SECRET As String = "api_secret"

Public Const TYPE_TEXT As String = "text"
Public Const TYPE_BINARY As String = "binary"
Public Const TYPE_WAPPUSH As String = "wappush"
Public Const TYPE_UNICODE As String = "unicode"
Public Const TYPE_VCAL As String = "vcal"
Public Const TYPE_VCARD As String = "vcard"

' Return NEXMO SMS Request Keys
' @author tat
Function keys() As Variant
    Dim rtnArray(18) As String
    rtnArray(0) = KEY_FROM
    rtnArray(1) = KEY_TO
    rtnArray(2) = KEY_TYPE
    rtnArray(3) = KEY_TEXT
    rtnArray(4) = KEY_STATUS_REPORT_REQ
    rtnArray(5) = KEY_CLIENT_REF
    rtnArray(6) = KEY_VCARD
    rtnArray(7) = KEY_VCAL
    rtnArray(8) = KEY_TTL
    rtnArray(9) = KEY_CALLBACK
    rtnArray(10) = KEY_MESSAGE_CLASS
    rtnArray(11) = KEY_UDH
    rtnArray(12) = KEY_PROTOCOL_ID
    rtnArray(13) = KEY_BODY
    rtnArray(14) = KEY_TITLE
    rtnArray(15) = KEY_URL
    rtnArray(16) = KEY_VALIDITY
    rtnArray(17) = KEY_API_KEY
    rtnArray(18) = KEY_API_SECRET
    keys = rtnArray
End Function


