Attribute VB_Name = "IE_access"
Option Explicit
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
   (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
    ) As Long

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

    ' ?{?^???N???b?N????J??
    Dim objtsugi As Object
    For Each objtsugi In objIE.Document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "????y?[?W") > 0 Then
            objtsugi.Click
            Call WaitFor(3)
            Exit For
        End If
    Next

    ' close
    objIE.quit
    Set objIE = Nothing

End Sub

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

sub main()
    Dim strURL As String
    Dim strPath As String
     
    strPath = "save path and savefilename"
    strURL = "https://www.google.co.jp/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png"
    call fileDownload(strPath,strURL)
    MsgBox "main end"
end sub




Function fileDownload(ByRef savePath as string,ByRef dlUrl as string)
    dim checkNum as Integer
    checkNum = URLDownloadToFile(0, dlUrl, savePath, 0, 0)
    If checkNum = 0 Then
        MsgBox "complateI"
    Else
        MsgBox "NG"
    End If
end Function


Function getUrl(ByRef url as string)
    dim sheetName as string
    sheetName = "work"

    Dim objIE As InternetExplorer 
    Set objIE = CreateObject("Internetexplorer.Application") 
    
    ' objIE.Visible = True 
    objIE.Visible = false
     
    objIE.navigate url 'IE??URL??J??

    ' HTML?????????m??
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '????????
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument 'HTML?h?L???????g?I?u?W?F?N?g???
    Set htmlDoc = objIE.document 'objIE????????????HTML?h?L???????g??Z?b?g

    sheets(sheetName).cells(1,1).value = htmlDoc.Title 'HTML?h?L???????g??^?C?g????\??
    dim elinks as IHTMLElement 
    dim cnt as Integer
    cnt = 2
    For Each elinks In htmlDoc.Links
        sheets(sheetName).cells(cnt,2).value = elinks.href
        cnt = cnt + 1
    Next elinks
 
End Function

