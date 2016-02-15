Attribute VB_Name = "SMS_ResponseConst"
'--------------------------------------------------------------------------------
' SMS Response Const
' @author tat
'--------------------------------------------------------------------------------

Public Const STATUS_CODE As String = "statusCode"
Public Const STATUS_TEXT  As String = "statusText"
Public Const RESPONSE_TEXT  As String = "responseText"

Public Const KEY_STATUS As String = "status"
Public Const KEY_MESSAGE_ID As String = "messageId"
Public Const KEY_TO As String = "to"
Public Const KEY_CLIENT_REF As String = "clientRef"
Public Const KEY_REMAINING_BALANCE As String = "remainingBalance"
Public Const KEY_MESSAGE_PRICE As String = "messagePrice"
Public Const KEY_NETWORK As String = "network"
Public Const KEY_ERROR_TEXT As String = "errorText"

' Return NEXMO SMS Respoonse Keys
' @author tat
Function keys() As Variant
    Dim rtnArray(7) As String
    rtnArray(0) = KEY_STATUS
    rtnArray(1) = KEY_MESSAGE_ID
    rtnArray(2) = KEY_TO
    rtnArray(3) = KEY_CLIENT_REF
    rtnArray(4) = KEY_REMAINING_BALANCE
    rtnArray(5) = KEY_MESSAGE_PRICE
    rtnArray(6) = KEY_NETWORK
    rtnArray(7) = KEY_ERROR_TEXT
    keys = rtnArray
End Function

