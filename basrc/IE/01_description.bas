' HTMLドキュメント.getElementsByName(name属性名)
' Dim コレクション名 As IHTMLElementCollection

sub main01()
    dim siteUrl as variant
    siteUrl = array ("http://www.yahoo.co.jp/","https://news.yahoo.co.jp/")
    call getDescription(siteUrl(0))
    call getDescription(siteUrl(1))
    MsgBox "end", vbButtonType, "msgTitle"
end sub


Function getDescription(byref siteUrl as variant)
    dim objIE as internetExplorer
    set objIE = CreateObject("Internetexplorer.Application")

    objIE.Visible = True
    objIE.navigate siteUrl

    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument
    Set htmlDoc = objIE.document

    Dim colDes As IHTMLElementCollection 
    Set colDes = htmlDoc.getElementsByName("description")

    Debug.Print colDes(0).Content

end Function