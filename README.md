# Nexmo-VBA
Nemo SMS API Utility for VBA

## Install
Import ALL Files.

## How to Use
### 1.Create SMS Client
    Dim client As New SMS_Client
    client.letAPI_KEY = {Nexmo API Key}
    client.letAPI_SECRET = {Nexmo API Secret}
### 2.Send
    Dim ret As Collection
    Set ret = client.send({From }, {To}, {Text})

## 3.Get response params
    Dim a as String
    a = ret.item(SMS_ResponseConst.KEY_ERROR_TEXT)

## Default request params
    Dim opts As New Collection
    Set client.setDefaultRequestparam = opts
    opts.Add SMS_RequestConst.TYPE_UNICODE, SMS_RequestConst.KEY_TYPE

## Utility

### EncodeUTF8
    
    Dim utf8Str as String
    utf8Str = encodeUtf8({baseString})
    
### E164
    Dim e164No as String
    e164No = getE164telNo({countryCode}, {telNo})
