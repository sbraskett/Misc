VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestList 
   Caption         =   "TestForm"
   ClientHeight    =   9000.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   OleObjectBlob   =   "TestList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents pDesignListBox As clsDesignListBox
Attribute pDesignListBox.VB_VarHelpID = -1
Private pDeactivateEvents As Boolean


Private Sub CommandButton1_Click()
    Test_2D_Array
End Sub

Private Sub UserForm_Initialize()

    pDeactivateEvents = True

    Dim arr As Variant
    arr = GetMarketPriceArr

    Set pDesignListBox = New clsDesignListBox

    With pDesignListBox
        .Headers = True
        .ColumnWidths = "30;50;110;90"
        .Create Me, 6, 6, 330, 300, arr
        .Frame.Left = (Me.Width - .Frame.Width) / 2
        .Frame.Top = (Me.Height - .Frame.Height) / 2 - 20

    End With
    
    Dim Labl
    For Each Labl In pDesignListBox.AllLabels
        Labl.Font.Name = "Segoe UI"
        Labl.Font.Size = 10
    Next Labl
    
    
    pDesignListBox.ClearFormats
    pDesignListBox.AddFormat 1, "=", "eos", vbRed, False, False
    pDesignListBox.AddFormat 1, "=", "dash", vbYellow, False, False
    pDesignListBox.AddFormat 1, "=", "omg", vbGreen, False, False
    pDesignListBox.ApplyFormats
    
    pDeactivateEvents = False
    FilterColBox.Visible = False
    FilterOpBox.Visible = False
    FilterColBox.AddItem "Rank"
    FilterColBox.AddItem "Symbol"
    FilterColBox.AddItem "Name"
    FilterColBox.AddItem "Price"

    FilterColBox.ListIndex = 0

    FilterOpBox.AddItem "contains"
    FilterOpBox.AddItem "="
    FilterOpBox.AddItem ">"
    FilterOpBox.AddItem "<"

    FilterOpBox.ListIndex = 0
    'Me.KeyPreview = True
    pDesignListBox.RowHeight = 18
    
    

End Sub

Private Sub Test_Array()

Dim rows As Variant
rows = pDesignListBox.SelectedRowArray

Dim i As Long
For i = LBound(rows) To UBound(rows)
    Debug.Print rows(i)
Next i


End Sub
Private Sub Test_2D_Array()

Dim data As Variant
data = pDesignListBox.SelectedDataArray

If Not IsEmpty(data) Then
    Sheet1.Range("A1").Resize( _
        UBound(data, 1) + 1, _
        UBound(data, 2) + 1 _
    ).Value = data
End If

End Sub

Private Sub ApplyFilterBtn_Click()

    Dim c As Long
    c = CLng(Me.FilterColBox.Value) - 1 'convert to 0-based column

    pDesignListBox.ClearFilters
    pDesignListBox.AddFilter c, Me.FilterOpBox.Value, Me.FilterTextBox.Value, False
    pDesignListBox.ApplyFilters

End Sub

Private Sub ClearFilterBtn_Click()
    pDesignListBox.ResetFilters
End Sub

Private Function GetMarketPriceArr() As Variant

    Dim data As Worksheet
    Set data = ThisWorkbook.Worksheets("Data")

    Dim LastRowMarketPrices As Long
    LastRowMarketPrices = data.Range("A1").CurrentRegion.rows.Count

    Dim MarketPriceArr As Variant
    MarketPriceArr = data.Range("A1:D" & LastRowMarketPrices).Value

    Dim i As Long
    Dim colPrice As Long
    colPrice = LBound(MarketPriceArr, 2) + 3   'A=+0, B=+1, C=+2, D=+3

    For i = LBound(MarketPriceArr, 1) + 1 To UBound(MarketPriceArr, 1)
        MarketPriceArr(i, colPrice) = Round(CDbl(MarketPriceArr(i, colPrice)), 4)
    Next i

    GetMarketPriceArr = MarketPriceArr

End Function
' needed for keypress events


Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Forward to your instance
    DesignListBox1_HandleKey CLng(KeyCode), CLng(Shift)  'rename as needed
End Sub

Public Sub DesignListBox1_HandleKey(ByVal KeyCode As Long, ByVal Shift As Long)
    Me.dlb.HandleKey KeyCode, Shift   'dlb is your clsDesignListBox instance
End Sub
