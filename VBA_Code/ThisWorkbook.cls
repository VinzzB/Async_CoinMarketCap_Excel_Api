VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''
'   Created by VinzzB (https://github.com/VinzzB)
'   Donations are welcome @ Litecoin address: LNaFvaSnedCNToBrYNm6Pz4UFEazhMNYnH
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    reader.StopTimer
End Sub

Private Sub Workbook_Open()
    'Bigger requests can result in timeouts or JSON read failures! Especially when the API is under heavy load.
    'eg: reader.Setup 1500
    'Too many request could lead to a (temp) ban.
    'eg: reader.Setup 100, 15
    
    'load all in one request. (could fail if api is under heavy load) and load prices in EURO.
    'reader.Setup 1500, 0, "EUR"
    
    'Load 1500 coins in three requests
    'reader.Setup 500,3, "USD"
    
    'Reload data every 2 minutes.
    reader.StartTimer "ReadApi", "00:02:00"
End Sub
