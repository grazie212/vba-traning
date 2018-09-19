Option Explicit

Sub main03()
    '// add declarations
    dim siteUrl as String
    siteUrl = "https://finance.yahoo.co.jp/"
    getLinkAll(siteUrl)
    On Error GoTo catchError

exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub

function getLinkAll(byref siteUrl As String)
    dim objIE as internetExplorer
    set objIE = CreateObject("Internetexplorer.Application")

    objIE.Visible = True
    objIE.navigate siteUrl

    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument
    Set htmlDoc = objIE.document

    dim el as IHTMLElement
    for each el in htmlDoc.links
        debug.print el.href
    next el

end function