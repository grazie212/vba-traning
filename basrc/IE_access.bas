Attribute VB_Name = "IE_access"
Option Explicit
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
   (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long
 public const sheetName as string = "work"

sub fileDownloadGooglePic()
    Dim strURL As String
    Dim strPath As String
     
    strPath = "save path and savefilename"
    strURL = "https://www.google.co.jp/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png"
    call fileDownload(strPath,strURL)
    MsgBox "main end"
end sub

sub yahoo()
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True

    dim siteUrl as string
    dim search as string

    siteUrl = "https://www.yahoo.co.jp/"

    objIE.Navigate siteUrl

    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE 
        DoEvents
    Loop
    
    Dim htmlDoc As HTMLDocument
    Dim objLink As IHTMLelement
    
    Set htmlDoc = objIE.document

    For Each objLink In htmlDoc.Links
        siteUrl = objLink.href
        objIE.navigate siteUrl
        Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE 
            DoEvents
        Loop
    Next objLink
    objIE.quit
    Set objIE = Nothing
    
end sub

Function IE_access(ByRef siteUrl As String,ByRef moji as string)
    ' ?C???^?[?l?b?g????????u???E?U??J??
    dim objIE as InternetExplorer
    set objIE = CreateObject("InternetExplorer.Application")

    objIE.Visible = True

    objIE.Navigate siteUrl
    call IEwait(objIE)
    call waitfor(3)

    dim objtag,objsubmit as object

    For Each objtag In objIE.Document.getElementsByTagName("input")
      If InStr(objtag.outerHTML, """yschsp""") > 0 Then
            objtag.Value = moji
            Exit For
      End If
    Next

    For Each objsubmit In objIE.Document.getElementsByTagName("input")
      If InStr(objsubmit.outerHTML, """?? ??""") > 0 Then
            objsubmit.Click
            Call WaitFor(3)
            Exit For
      End If
    Next

    Dim objtsugi As Object
    For Each objtsugi In objIE.Document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "????y?[?W") > 0 Then
            objtsugi.Click
            Call WaitFor(3)
            Exit For
        End If
    Next

End Function

Function IEWait(ByRef objIE As Object)
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
    Loop
End Function

Function WaitFor(ByVal second As Integer)
    Dim futureTime As Date
 
    futureTime = DateAdd("s", second, Now)
 
    While Now < futureTime
        DoEvents
    Wend
End Function

Function fileDownload(ByRef savePath as string,ByRef dlUrl as string)
    dim checkNum as Integer
    checkNum = URLDownloadToFile(0, dlUrl, savePath, 0, 0)
    If checkNum = 0 Then
        MsgBox "complate?I"
    Else
        MsgBox "NG"
    End If
end Function


Function getLink(ByRef objIE As InternetExplorer,ByRef siteUrl as string)

    objIE.Navigate siteUrl
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE 
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument 
    Set htmlDoc = objIE.document 

    sheets(sheetName).cells(1,1).value = htmlDoc.Title 
    dim elinks as IHTMLElement
    dim cnt as Integer
    cnt = 2
    For Each elinks In htmlDoc.Links
        objIE.navigate elinks.href
        WaitFor(5)
        cnt = cnt + 1
    Next elinks
 
End Function

Function yahoo_search(ByRef objIE as Object,ByRef siteUrl As String,ByRef search as string)

    objIE.Navigate siteUrl
    call IEwait(objIE)
    call waitfor(3)
    
    dim objtag,objsubmit as object
    For Each objtag In objIE.Document.getElementsByTagName("input")
      If InStr(objtag.outerHTML, """srchtxt""") > 0 Then
            objtag.Value = search
            Exit For
      End If
    Next

    For Each objsubmit In objIE.Document.getElementsByTagName("input")
      If InStr(objsubmit.outerHTML, """srchbtn""") > 0 Then
            objsubmit.Click
            Call WaitFor(3)
            Exit For
      End If
    Next    

    ' ' close
    ' objIE.quit
    ' Set objIE = Nothing

end Function


Function IELinkClick(ByRef objIE As Object, ByVal anchorText As String)
    Dim objLink As Object
 
    For Each objLink In objIE.Document.getElementsByTagName("A")
        If objLink.innerText = anchorText Then
            objIE.navigate objLink.href
            Exit For
        End If
    Next
End Function

Function IELinkClickAll(ByRef objIE As Object)
    Dim objLink As Object
 
    For Each objLink In objIE.Document.getElementsByTagName("A")
        objIE.navigate objLink.href
        Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE 
            DoEvents
        Loop
    Next
End Function