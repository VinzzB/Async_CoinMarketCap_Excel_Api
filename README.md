# Async CoinMarketCap Excel Api
This library loads Crypto Currency data from the CoinMarketCap API into Memory. All data can be exposed through Excel functions in a worksheet.
By default, data is refreshed every 2 minutes. 

The Excel workbook in this repository is of type 'xlsm'. This means that macros are being used in this Excel file. The library will not work without macro's enabled. 
All code can be examined through the vba editor (ALT+F11) for revision. This code is free from any virusses, trojans, malware,...
You can optionally supress the security warnings for this excel file by making it trusted. (See File menu -> Enable Content)

## How Does It Work

-  This library loads 1500 coins (by default) from CoinMarketCap.com API (CMC-API) in one request and without blocking the user interface. The amount of coins and requests can be adjusted in the Workbook_Open() method by calling the Setup() method in the Reader class. (see example)
-  Prices are shown in USD by default. You can choose to convert prices to another currency in the Workbook_Open() method by calling the Setup() method in the Reader class.
all supported currencies can be found at https://coinmarketcap.com/api/
- The refresh timer is also adjustable through the StartTimer() method.

Examples (open VBA in Excel with ALT + F11)
```vba
Private Sub Workbook_Open()
    
    'load all in one request. (could fail if api is under heavy load) and load prices in EURO.
    reader.Setup 1500, 0, "EUR"
    
    'Load 1500 coins in three requests
    'reader.Setup 500,3, "USD"
    
    'Reload data every 2 minutes.
    reader.StartTimer "ReadApi", "00:02:00"
End Sub
```

- These function are available in an excel worksheet:
  - **GetCurrencyFor**(coinName As String, dummyDateTime As Date) As Double
  - **GetCurrencyForTicker**(coinTicker As String, dummyDateTime As Date) As Double
  - **GetCoinOnRank**(Rank As Integer, getFieldName As String, dummyDateTime As Date)
  - **GetCoinForName**(coinName As String, getFieldName As String, dummyDateTime As Date) as Variant
  - **GetCoinForTicker**(coinTicker As String, getFieldName As String, dummyDateTime As Date) as Variant
  - **LastUpdate**(dummyDateNow As Date) As Date
  - **NextUpdate**(dummyDateNow As Date) As Date
  
  All functions uses  a 'dummyDateNow' parameter. Use the NOW() function, this will force Excel to update the cell on every refresh.
  examples for an excel cell formula:
  ```vba
  
  GetCurrencyFor("bitcoin", now())
  
  GetCurrencyForTicker("btc", now())
  
  GetCoinForTicker("btc", "24hchange", now())
  
  ```
  GetCoinOnRank, GetCoinForName, and GetCoinForTicker uses the 'getFieldName' argument to get specific data from a coin.
  These are all valid fieldnames:
  
  - *id*
  - *circulatingsupply* or *availablesupply*
  - *totalsupply*
  - *marketcap*
  - *name*
  - *percentchange1h* or *1hchange* or *change1h*
  - *percentchange24h* or *24hchange* or *change24h*
  - *percentchange7d* or *7dchange* or *change7d*
  - *price*
  - *pricebtc* or *btcprice*
  - *rank*
  - *ticker*
  - *volume24* or *volume24h* or *volume* or *24hvolume*
  
- You can also attach the ReadApi() method to a custom ribbon or quick toolbar button which allows you to load (or refresh) the CMC data manually. 
  A button is provided in the excel workbook.
  
  
  If you have problems with the WinHttpRequest library. please check https://stackoverflow.com/questions/3119207/sending-http-requests-with-vba-from-word
  
  Donations are always welcome @ Litecoin address: LNaFvaSnedCNToBrYNm6Pz4UFEazhMNYnH