VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
'   Created by VinzzB (https://github.com/VinzzB)
'   Donations are welcome @ Litecoin address: LNaFvaSnedCNToBrYNm6Pz4UFEazhMNYnH
''
Option Explicit

Private pId                 As String   'id                         "ethereum"
Private pName               As String   'name                       "Ethereum"
Private pTicker             As String   'Symbol                     "ETH"
Private pRank               As Integer  'rank                       2
Private pPrice              As Double   'price_usd / eur / ...      12,1844
Private pPriceBtc           As Double   'price_btc                  0,021262
Private p24hVolume          As Double   '24h_volume_usd / eur / ... 24085900,0
Private pMarketCap          As Double   'market_cap_usd / eur / ... 1018098455,0
Private pAvailableSupply    As Double   'available_supply           83557537,0
Private pTotalSupply        As Double   'total_supply               83557537,0
Private pPercentChange1h    As Single   'percent_change_1h          -0,58
Private pPercentChange24h   As Single   'percent_change_24h         6,38
Private pPercentChange7d    As Single   'percent_change_7d          8,58
Private pLastUpddate        As Date     'last_updated               1472762062  (js date)
Property Get Id() As String:                Id = pId:                               End Property
Property Get Name() As String:              Name = pName:                           End Property
Property Get Ticker() As String:            Ticker = pTicker:                       End Property
Property Get Rank() As Integer:             Rank = pRank:                           End Property
Property Get Price() As Double:             Price = pPrice:                         End Property
Property Get PriceBtc() As Double:          PriceBtc = pPriceBtc:                   End Property
Property Get Volume24h() As Double:         Volume24h = p24hVolume:                 End Property
Property Get MarketCap() As Double:         MarketCap = pMarketCap:                 End Property
Property Get AvailableSupply() As Double:   AvailableSupply = pAvailableSupply:     End Property
Property Get TotalSupply() As Double:       TotalSupply = pTotalSupply:             End Property
Property Get PercentChange1h() As Single:   PercentChange1h = pPercentChange1h:     End Property
Property Get PercentChange24h() As Single:  PercentChange24h = pPercentChange24h:   End Property
Property Get PercentChange7d() As Single:   PercentChange7d = pPercentChange7d:     End Property
Property Get LastUpdate() As Date:          LastUpdate = pLastUpddate:               End Property
Public Sub SetData(value As Variant, useCurrency As String)
    pId = value("id")
    pName = value("name")
    pTicker = value("symbol")
    pRank = CInt(value("rank"))
    pPrice = ParseToDouble(value, "price_" & useCurrency)
    pPriceBtc = ParseToDouble(value, "price_btc")
    p24hVolume = ParseToDouble(value, "24h_volume_" & useCurrency)
    pMarketCap = ParseToDouble(value, "market_cap_" & useCurrency)
    pAvailableSupply = ParseToDouble(value, "available_supply")
    pTotalSupply = ParseToDouble(value, "total_supply")
    pPercentChange1h = ParseToSingle(value, "percent_change_1h") / 100
    pPercentChange24h = ParseToSingle(value, "percent_change_24h") / 100
    pPercentChange7d = ParseToSingle(value, "percent_change_7d") / 100
    pLastUpddate = ParseToDate(value, "last_updated")
End Sub

Private Function ParseToDouble(values As Variant, extractField As String) As Double
    ParseToDouble = CDbl(Replace( _
                                 IIf(IsNull(values(extractField)), 0, values(extractField)), _
                                 ".", _
                                 Application.DecimalSeparator))
End Function
Private Function ParseToSingle(values As Variant, extractField As String) As Double
    ParseToSingle = CSng(Replace( _
                                 IIf(IsNull(values(extractField)), 0, values(extractField)), _
                                 ".", _
                                 Application.DecimalSeparator))
End Function
Private Function ParseToDate(values As Variant, extractField As String) As Date
    ParseToDate = DateAdd("s", CDbl(values(extractField)), #1/1/1970#)
End Function
