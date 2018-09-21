Attribute VB_Name = "mdlMoTrPAC_API"
Option Explicit

'API settings
Const API_UserName = "stas.rirak@mssm.edu" '"wakepass"
Const API_Password = "BB846687-038A-D5BE-0509801DBBD95124"
Const API_ContentType = "application/json"
Const API_Accept = "application/json"
Const API_URL = "https://www.motrpac.org/rest/motrpacapi/biospecimen/{MID}"

'Excel worksheet settings
Const Wks_Target = "API_Test"
Const Num_Rec = "B3"
Const Tbl_FirstCell = "A4"
Const Msgbox_title = "MoTrPAC API"
    


Sub Test_RESTapi()

    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String

    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    strUrl = "https://jsonplaceholder.typicode.com/posts/1"
    blnAsync = True

    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .ResponseText
    End With

    Debug.Print strResponse

End Sub

Public Sub Test_MoTrPAC_api(MID As String)
    
    
    
    'Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim jsonResponse As New Dictionary
    Dim i As Integer, cnt As Integer
    Dim d As Dictionary
    Dim r As Range
    Dim xmlhttp As New MSXML2.XMLHTTP60 'Requires to set reference (in Tools/Reference) to Microsoft XML, v6.0
    
    'Set objRequest = CreateObject("MSXML2.XMLHTTP60")
    '''strUrl = "https://wakepass@www.motrpac.org/rest/motrpacapi/biospecimen/99901"
    'strUrl = "https://www.motrpac.org/rest/motrpacapi/biospecimen/99901"
    strUrl = Replace(API_URL, "{MID}", MID) 'Worksheets("API_Test").Range("B1").Value
    blnAsync = True

    With xmlhttp 'objRequest
        .Open "GET", strUrl, blnAsync, API_UserName, API_Password 'user name = wakepass;
        .SetRequestHeader "Content-Type", API_ContentType '"application/json"
        .SetRequestHeader "Accept", API_Accept '"application/json"; "application/xml"; "text/csv"
        .Send
        'wait for for response
        While xmlhttp.readyState <> 4
            DoEvents
        Wend
        strResponse = .ResponseText
        
        If Len(Trim(strResponse)) > 0 Then
            'clean area where to data will be posted.
            
            With Worksheets(Wks_Target) '"API_Test"
                .Range(Num_Rec).Clear 'clean row count field
                
                'clean area of main table output
                Dim row_end As Integer
                If .UsedRange.Rows.Count > .Range(Tbl_FirstCell).Row Then
                    row_end = .UsedRange.Rows.Count 'table currently is not empty
                Else
                    row_end = .Range(Tbl_FirstCell).Row ' table is empty
                End If
                .Range(Tbl_FirstCell, .Cells(row_end, .UsedRange.Columns.Count).Address).Clear
            End With
        
        
            'jsonResponse.CompareMode = TextCompare
            
            Set jsonResponse = JsonConverter.ParseJson(strResponse)
            
            'Validate response
            If jsonResponse.Exists("errorCode") Then
                'Report error
                MsgBox "API service reported and error." & vbCrLf & _
                        "Error code: " & jsonResponse("errorCode") & vbCrLf & _
                        "Message: " & jsonResponse("message"), vbCritical, Msgbox_title
            Else
                'retrieve data
                
                Worksheets(Wks_Target).Range(Num_Rec).Value = jsonResponse.Items(0)("recordcount")
                cnt = 0
                
                For Each d In jsonResponse("data")
                    If cnt = 0 Then
                        'Print column headers
                        For i = 0 To d.Count - 1
                            Worksheets(Wks_Target).Range(Tbl_FirstCell).Offset(0, i).Value = d.Keys(i)
                        Next
                    End If
                    
                    For i = 0 To d.Count - 1
                        Worksheets(Wks_Target).Range(Tbl_FirstCell).Offset(cnt + 1, i).Value = d.Items(i)
        '                    Debug.Print d.Keys(i)
        '                    Debug.Print d.Items(i)
                    Next
                    'Debug.Print d("sampleTypeCode")
                    cnt = cnt + 1
                Next
                
                'MsgBox "Data for BID " & Worksheets(Wks_Target).Range("B1").Value & " was successfully received.", vbInformation, Msgbox_title
            End If
        
        Else
            'Report error
            MsgBox "API service returned no response. Please verify that the Internet is available and the API's URL (" & strUrl & ") is reachable.", vbCritical, Msgbox_title
        
        End If
        
    End With

    'Debug.Print strResponse

End Sub


'Test_MoTrPAC_api
'
'Print jsonResponse("meta")("recordcount")
'23
'Print jsonResponse("data").Count
'23
'23
'meta
'Set v = jsonResponse.Items(0).Item(0).Item(0)
'Print v.Count
'Print v("bid")
'99901
'
'Print v.Keys.Count
'Print strResponse

