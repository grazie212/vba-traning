Attribute VB_Name = "IE_access"

Option Explicit
sub IE_access()
    ' インターネットに接続してブラウザを開く
    dim objIE as InternetExplorer
    set objIE = CreateObject("InternetExplorer.Application")
    '可視化
    objIE.Visible = True

    ' 指定のページを開く
    dim siteUrl As String
    siteUrl = "http://auctions.yahoo.co.jp/"
    objIE.Navigate siteUrl
    call IEwait(objIE)
    call waitfor(3)

    ' IEに自動文字入力
    dim moji as String
    moji = "colnago"
    dim objtag,objsubmit as object

    For Each objtag In objIE.Document.getElementsByTagName("input")
      If InStr(objtag.outerHTML, """yschsp""") > 0 Then
            objtag.Value = moji
            Exit For
      End If
    Next

    For Each objsubmit In objIE.Document.getElementsByTagName("input")
      If InStr(objsubmit.outerHTML, """検 索""") > 0 Then
            objsubmit.Click
            Call WaitFor(3)
            Exit For
      End If
    Next

    ' ボタンクリックで画面遷移
    Dim objtsugi As Object
    For Each objtsugi In objIE.Document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "次のページ") > 0 Then
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


Sub testIE()
    dim sheetName as string
    sheetName = "work"

    Dim objIE As InternetExplorer 'IEオブジェクトを準備
    Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
    
    ' objIE.Visible = True 'IEを表示
    objIE.Visible = false
    dim url as string
    url = "https://tonari-it.com/vba-ie-links/"
    objIE.navigate url 'IEでURLを開く

    ' HTML読み込み時間確保
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument 'HTMLドキュメントオブジェクトを準備
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット

    sheets(sheetName).cells(1,1).value = htmlDoc.Title 'HTMLドキュメントのタイトルを表示
    dim elinks as IHTMLElement 
    dim cnt as Integer
    cnt = 2
    For Each elinks In htmlDoc.Links
        sheets(sheetName).cells(cnt,2).value = elinks.href
        cnt = cnt + 1
    Next elinks
 
End Sub