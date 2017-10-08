Attribute VB_Name = "Module1"
Option Explicit

Private Sub CrawlProxy()

    On Error Resume Next

    Dim WinHttp As WinHttpRequest
    Set WinHttp = New WinHttpRequest
    
    Dim HtmlDoc As MSHTML.HTMLDocument
    Set HtmlDoc = New MSHTML.HTMLDocument
    Dim ProxyTable As Object
    Dim TableRows As Object
    Dim i As Integer
    
    With Worksheets("Proxy")
        .Cells.Clear
        .Range("a1").Select
    End With
    
    
    WinHttp.Open "GET", "http://www.xicidaili.com/nn"
    WinHttp.send
    HtmlDoc.body.innerHTML = WinHttp.responseText


    Set ProxyTable = HtmlDoc.getElementById("ip_list")
    Debug.Print ProxyTable.innerText
  
    Set TableRows = ProxyTable.getElementsByTagName("tr")
    Debug.Print TableRows.Length
    
    For i = 1 To TableRows.Length - 1
        ActiveCell.Value = TableRows(i).getElementsByTagName("td")(1).innerText
        ActiveCell.Offset(0, 1).Value = TableRows(i).getElementsByTagName("td")(2).innerText
        ActiveCell.Offset(1, 0).Select
    Next i

    Set HtmlDoc = Nothing
    Set WinHttp = Nothing
End Sub


