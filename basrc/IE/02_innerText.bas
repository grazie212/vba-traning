sub main02()
    dim siteUrl as string
    siteUrl="https://news.yahoo.co.jp/"
    call getInnerText(siteUrl)
    MsgBox "02 end", vbButtonType, "msgTitle"
end sub


Function getInnerText(byref siteUrl as string)
    dim objIE as internetExplorer
    set objIE = CreateObject("Internetexplorer.Application")

    objIE.Visible = True
    objIE.navigate siteUrl

    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument
    Set htmlDoc = objIE.document

    Dim colH1 As IHTMLElementCollection 
    dim colH2 As IHTMLElementCollection 
    dim colH3 As IHTMLElementCollection

    Set colH1 = htmlDoc.getElementsByTagName("h1")
    ' Set colH2 = htmlDoc.getElementsByTagName("h2")
    ' Set colH3 = htmlDoc.getElementsByTagName("h3")

    Dim el As IHTMLElement

    Debug.Print "h1"
    For Each el In colH1
    
        Debug.Print el.innerText
        
    Next el
    
    ' Debug.Print "h2"
    ' For Each el In colH2
    
    '     Debug.Print el.innerText
        
    ' Next el
    
    ' Debug.Print "h3"
    ' For Each el In colH3
    
    '     Debug.Print el.innerText
        
    ' Next el

end Function
