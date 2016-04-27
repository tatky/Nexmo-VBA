Nexmo Client Library for Visual Basic for Application
===================================

[Installation](#Installation) |  [Usage](#Usage) |  [Examples](#Examples) | [Coverage](#API-Coverage) | [Contributing](#Contributing)  

This is the VBA client library for use Nexmo's API. To use this, you'll need a Nexmo account. Sign up [for free
at nexmo.com][signup].

Installation
------------

Import ALL Files to VBAProject.

Usage
-----

Specify your credentials let to client's field.
For example:

```VisualBasic
Dim client As New SMS_Client
client.letAPI_KEY = {Nexmo API Key}
client.letAPI_SECRET = {Nexmo API Secret}
```



Examples
--------

### Sending A Message

Use [Nexmo's SMS API][doc_sms] to send an SMS message. 

```VisualBasic
Dim ret As Collection
Set ret = client.send({From}, {To}, {Text})
```

### Get send request response

```VisualBasic
Dim hoge as String
hoge = ret.item(SMS_ResponseConst.KEY_ERROR_TEXT)
```

### Set request's common patameter

```VisualBasic
Dim opts As New Collection
Set client.setDefaultRequestparam = opts
opts.Add SMS_RequestConst.TYPE_UNICODE, SMS_RequestConst.KEY_TYPE
```

### Utility

#### EncodeUTF8
```VisualBasic
Dim utf8Str as String
utf8Str = encodeUtf8({baseString})
```

#### E164
```VisualBasic
Dim e164No as String
e164No = getE164telNo({countryCode}, {telNo})
```

API Coverage
------------

* Account
    * [ ] Balance
    * [ ] Pricing
    * [ ] Settings
    * [ ] Top Up
    * [ ] Numbers
* Number
    * [ ] Search
    * [ ] Buy
    * [ ] Cancel
    * [ ] Update
* NumberInsight
    * [ ] Request
    * [ ] Response
* NumberVerify
    * [ ] Verify
    * [ ] Check
    * [ ] Search
    * [ ] Control
* Search
    * [ ] Message
    * [ ] Messages
    * [ ] Rejections
* Short Code
    * [ ] 2FA
    * [ ] Alerts
    * [ ] Marketing
* SMS
    * [X] Send
    * [ ] Receipt
    * [ ] Inbound
* Voice
    * [ ] Call
    * [ ] TTS/TTS Prompt
    * [ ] SIP

Contributing
------------

License
-------

This library is released under the [MIT License][license]

[signup]: http://nexmo.com?src=vba-client-library
[doc_sms]: https://docs.nexmo.com/api-ref/sms-api
[license]: LICENSE.txt
