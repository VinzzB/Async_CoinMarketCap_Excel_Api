Attribute VB_Name = "CoinReaderHelper"
''
'   Created by VinzzB (https://github.com/VinzzB)
'   Donations are welcome @ Litecoin address: LNaFvaSnedCNToBrYNm6Pz4UFEazhMNYnH
'
Option Explicit
'''''''''''''''''''''''''''''''
'initialize coinReader class.
'''''''''''''''''''''''''''''''
Public reader As New CoinReader
'''''''''''''''''''''''''''''''
' EXCEL WORKSHEET FUNCTIONS
'-----------------------------
' use the 'NOW()' function for dummyDateTime. This forces the excel field to be volatile.(= not using cached data on refresh.)
'''''''''''''''''''''''''''''''

Public Function GetCurrencyFor(coinName As String, dummyDateTime As Date) As Double
   On Error Resume Next
   GetCurrencyFor = reader.GetCoinForName(coinName).Price
End Function

Public Function GetCurrencyForTicker(coinTicker As String, dummyDateTime As Date) As Double
   On Error Resume Next
   GetCurrencyForTicker = reader.GetCoinForTicker(coinTicker).Price
End Function

Public Function LastUpdate(dummyDateNow As Date) As Date
    LastUpdate = reader.LastUpdate
End Function
Public Function NextUpdate(dummyDateNow As Date) As Date
    NextUpdate = reader.NextUpdate
End Function

Private Function GetFieldFromCoin(coin As coin, getFieldName As String)
    If Not coin Is Nothing Then
        Select Case LCase(getFieldName)
        Case "id":               GetFieldFromCoin = coin.Id
        Case "circulatingsupply", "availablesupply"
                                 GetFieldFromCoin = coin.AvailableSupply
        Case "totalsupply":      GetFieldFromCoin = coin.TotalSupply
        Case "marketcap":        GetFieldFromCoin = coin.MarketCap
        Case "name":             GetFieldFromCoin = coin.Name
        Case "percentchange1h", "1hchange", "change1h"
                                 GetFieldFromCoin = coin.PercentChange1h
        Case "percentchange24h", "24hchange", "change24h"
                                 GetFieldFromCoin = coin.PercentChange24h
        Case "percentchange7d", "7dchange", "change7d"
                                 GetFieldFromCoin = coin.PercentChange7d
        Case "price":            GetFieldFromCoin = coin.Price
        Case "pricebtc", "price btc", "btcprice", "btc price"
                                 GetFieldFromCoin = coin.PriceBtc
        Case "rank":             GetFieldFromCoin = coin.Rank
        Case "ticker":           GetFieldFromCoin = coin.Ticker
        Case "volume24", "volume24h", "volume", "24hvolume"
                                 GetFieldFromCoin = coin.Volume24h
                                 
        Case Else:               GetFieldFromCoin = "#invalid field#"
        End Select
    End If
End Function

Public Function GetCoinOnRank(Rank As Integer, getFieldName As String, dummyDateTime As Date) As Variant
    GetCoinOnRank = GetFieldFromCoin(reader.GetCoinOnRank(Rank), getFieldName)
End Function

Public Function GetCoinForName(coinName As String, getFieldName As String, dummyDateTime As Date) As Variant
    GetCoinForName = GetFieldFromCoin(reader.GetCoinForName(coinName), getFieldName)
End Function

Public Function GetCoinForTicker(coinTicker As String, getFieldName As String, dummyDateTime As Date) As Variant
    GetCoinForTicker = GetFieldFromCoin(reader.GetCoinForTicker(coinTicker), getFieldName)
End Function

''''''''''''''''''''''''''''''''
' Get data from CMC API.
' You can attach this function to a button on the ribbon or on the quick access toolbar through the excel options.
' This function is also configured as the timer callback in VBA file 'ThisWorkbook'.
''''''''''''''''''''''''''''''''
Public Sub ReadApi()
    On Error GoTo Errhandler
    Debug.Print Now & " - refreshing data..."
    reader.ReadFromWeb 'this will reset the timer (if configured)
ExitHandler:
    Exit Sub
Errhandler:
    Application.StatusBar = "Failed to load crypto currencies!"
    Err.Clear
    Resume ExitHandler
End Sub


