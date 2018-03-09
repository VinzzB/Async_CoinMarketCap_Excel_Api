VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoinReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
'   Created by VinzzB (https://github.com/VinzzB)
'   Donations are welcome @ Litecoin address: LNaFvaSnedCNToBrYNm6Pz4UFEazhMNYnH

'   This library loads Crypto Currency data from the CoinMarketCap API into Excel Worsheets.
'   Async methods are being used from the WinHttp module to avoid Excel from freezing while loading web data.
'   Still, Excel could temporarly freeze if the 'auto-recovery' feature is enabled and an async call is being executed. You can disable 'auto-recovery' for the current workbook through the excel options.
'
'   The workbook will automatically recalculate when new data is loaded. You can trigger the 'ReadFromWeb' function manually or with the use of a timer.
'   You can retrieve the data for a coin through two methods: xxx.GetCoinForName(name) or xxx.GetCoinForTicker(ticker)
'
'   options:
'    - Convert currency option (EUR, AUD, CAD, GBP, ..) all supported currencies at https://coinmarketcap.com/api/
'    - Split requests.
'    - Timer
'
''


Option Explicit
Private coins As New Dictionary
Private WithEvents http As WinHttp.WinHttpRequest
Attribute http.VB_VarHelpID = -1
Private currentReadingPosition As Integer 'iterator for async calls.
Private plimit As Integer 'Max results (coin prices) per request from coinmarketcap API
Private pLoadCoinRank As Integer
Private IsLoaded As Boolean
Private isLoading As Boolean
Private pLastUpdate As Date
Private pTimerTime As String
Private pTimerCallBack As String
Private pConvertCurrency As String
Private pRefreshOn As Date

Private Sub Class_Initialize()
    SetupHttp
    'Default values
    Setup 1500, 0
    pTimerTime = "00:02:00"
End Sub

Public Property Get ConvertCurrency() As String
    ConvertCurrency = pConvertCurrency
End Property

Public Property Get TimerTime() As String
    TimerTime = pTimerTime
End Property

Private Property Get TimerValue() As Date
     TimerValue = Now + timeValue(pTimerTime)
End Property

Public Sub StartTimer(Optional callback As String = "", Optional time As String = "")
    If time <> "" Then pTimerTime = time
    If callback <> "" Then pTimerCallBack = callback
    If pTimerCallBack = "" Or TimerTime = "" Or pLoadCoinRank = 0 Then Exit Sub
    If coins.Count = 0 Then
        ReadFromWeb
        Exit Sub
    End If
    pRefreshOn = TimerValue
    Debug.Print "Next refresh @ " & pRefreshOn
    Application.OnTime pRefreshOn, pTimerCallBack
End Sub
Public Sub StopTimer()
On Error Resume Next
    Application.OnTime pRefreshOn, pTimerCallBack, , False
End Sub

Public Sub Setup(maxRank As Integer, Optional limitRequest As Integer = 0, Optional convertToCurrency As String = "usd")
    plimit = IIf(limitRequest = 0, maxRank, limitRequest)
    pLoadCoinRank = maxRank
    pConvertCurrency = LCase(convertToCurrency)
End Sub

Public Sub ReadFromWeb()
    ReadAt 0, plimit
End Sub
Private Sub SetupHttp()
    Set http = New WinHttp.WinHttpRequest
    http.SetTimeouts 10000, 30000, 30000, 120000
End Sub
Private Sub ReadAt(start As Integer, limit As Integer)
    If isLoading Or pLoadCoinRank = 0 Then Exit Sub
    SetupHttp
    Debug.Print currentReadingPosition & " Reading from " & start & " to " & limit + start
    Application.StatusBar = "Reading from " & start & " to " & limit + start
    isLoading = True
    plimit = limit
    If start = 0 Then currentReadingPosition = 0
    http.Open "GET", "https://api.coinmarketcap.com/v1/ticker/?start=" _
                    & start & "&limit=" & limit _
                    & IIf(pConvertCurrency <> "", "&convert=" & pConvertCurrency, ""), True
    http.Send
End Sub

Private Sub http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
On Error Resume Next
 Debug.Print ErrorDescription
 Application.StatusBar = ErrorDescription
 isLoading = False
 IsLoaded = True 'prevent iterations
 StartTimer
End Sub

Public Function GetCoinOnRank(rank_nr As Integer) As coin
    If rank_nr < 0 Then Exit Function
    If Not IsLoaded Then ReadFromWeb
    If Not IsEmpty(coins) Then
        Set GetCoinOnRank = coins.Items(rank_nr - 1)
        Exit Function
    End If
    Set GetCoinOnRank = New coin
End Function

Public Function GetCoinForName(coinName As String) As coin
Dim value As Dictionary
If IsEmpty(coinName) Then Exit Function
If Not IsLoaded Then ReadFromWeb
Dim coin As Variant
If Not IsEmpty(coins) Then

    For Each coin In coins.Items()
        If LCase(coin.Name) = LCase(coinName) Then
           Set GetCoinForName = coin
            Exit Function
        End If
    Next coin
End If
Set GetCoinForName = New coin
End Function
Public Function GetCoinForTicker(coinTicker As String) As coin
Dim value As Dictionary
If IsEmpty(coinTicker) Then Exit Function
If Not IsLoaded Then ReadFromWeb
Dim coin As Variant
If Not IsEmpty(coins) Then

    For Each coin In coins.Items()
        If LCase(coin.Ticker) = LCase(coinTicker) Then
            Set GetCoinForTicker = coin
            Exit Function
        End If
    Next coin
End If
Set GetCoinForTicker = New coin
End Function

Private Sub ReadJson(jsonString As String)
On Error GoTo ExitSub
Dim jsonData As Variant
Dim coin As coin
Dim value As Variant
Set jsonData = ParseJson(jsonString)
If Not IsEmpty(jsonData) Then
    For Each value In jsonData
        Set coin = New coin
        coin.SetData value, pConvertCurrency
         If coins.Exists(LCase(coin.Name)) Then
            Set coins(LCase(coin.Name)) = coin
         Else
            coins.Add LCase(coin.Name), coin
         End If
Next value
End If
ExitSub:
End Sub

Private Sub http_OnResponseFinished()
    On Error Resume Next
    Debug.Print currentReadingPosition & " parsing data..."
    currentReadingPosition = currentReadingPosition + 1
    ReadJson http.ResponseText
    pLastUpdate = Now
    Calculate
    isLoading = False
    If currentReadingPosition > (pLoadCoinRank / plimit) - 1 Then
        IsLoaded = True
        Application.StatusBar = False
        currentReadingPosition = 0
        StartTimer
    Else
        ReadAt plimit * currentReadingPosition, plimit
    End If
        
End Sub

Property Get LastUpdate() As Date
    LastUpdate = pLastUpdate
End Property
Property Get NextUpdate() As Date
    NextUpdate = pRefreshOn
End Property
