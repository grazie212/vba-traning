Attribute VB_Name = "IE_access"

Option Explicit
sub IE_access()
    ' �C���^�[�l�b�g�ɐڑ����ău���E�U���J��
    dim objIE as InternetExplorer
    set objIE = CreateObject("InternetExplorer.Application")
    '����
    objIE.Visible = True

    ' �w��̃y�[�W���J��
    dim siteUrl As String
    siteUrl = "http://auctions.yahoo.co.jp/"
    objIE.Navigate siteUrl
    call IEwait(objIE)
    call waitfor(3)

    ' IE�Ɏ�����������
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
      If InStr(objsubmit.outerHTML, """�� ��""") > 0 Then
            objsubmit.Click
            Call WaitFor(3)
            Exit For
      End If
    Next

    ' �{�^���N���b�N�ŉ�ʑJ��
    Dim objtsugi As Object
    For Each objtsugi In objIE.Document.getElementsByTagName("a")
        If InStr(objtsugi.outerHTML, "���̃y�[�W") > 0 Then
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

    Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
    Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
    
    ' objIE.Visible = True 'IE��\��
    objIE.Visible = false
    dim url as string
    url = "https://tonari-it.com/vba-ie-links/"
    objIE.navigate url 'IE��URL���J��

    ' HTML�ǂݍ��ݎ��Ԋm��
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop

    Dim htmlDoc As HTMLDocument 'HTML�h�L�������g�I�u�W�F�N�g������
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g

    sheets(sheetName).cells(1,1).value = htmlDoc.Title 'HTML�h�L�������g�̃^�C�g����\��
    dim elinks as IHTMLElement 
    dim cnt as Integer
    cnt = 2
    For Each elinks In htmlDoc.Links
        sheets(sheetName).cells(cnt,2).value = elinks.href
        cnt = cnt + 1
    Next elinks
 
End Sub