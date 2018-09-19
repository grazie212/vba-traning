Option Explicit

Sub main04()
    '// add declarations
    dim siteUrl as String
    siteUrl = "http://traininfo.jreast.co.jp/train_info/shinkansen.aspx"
    call getTableData(siteUrl)
    On Error GoTo catchError
exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub

Function getTableData(byref siteUrl as variant)
    dim objIE as internetExplorer
    set objIE = CreateObject("Internetexplorer.Application")

    objIE.Visible = True
    objIE.navigate siteUrl

    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument
    Set htmlDoc = objIE.document

    Dim colTR, colTH, colTD, colImg  As IHTMLElementCollection
    Set colTR = htmlDoc.getElementsByTagName("tr")

    Dim el As IHTMLElement
    For Each el In colTR
        Set colTH = el.getElementsByTagName("th")
        Set colTD = el.getElementsByTagName("td")
        Set colImg = el.getElementsByTagName("img")
        Debug.Print colTH(0).innerText & "|" & colImg(0).alt & "|" & colTD(1).innerText
    next el

end Function