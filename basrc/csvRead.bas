Attribute VB_Name = "csvFileRead"

Option Explicit

' 作業用シート
const sheetName as string = "work"

sub csvFileRead()
    dim bookPath as string
    dim workPath as string
    dim fileName as string
    dim fileName2 as string
    
    bookPath = ThisWorkbook.Path
    workPath = "\work\"
    fileName = bookPath & workpath & "test.csv"
    fileName2 = bookPath & workpath & "testLf.csv"

    ' call csvFileReadFunc(fileName)
    ' call csvFileReadLfFunc(fileName2)
    MsgBox "csvFileRead END"
end sub

' crlfの読み込み
function csvFileReadFunc(ByRef filename as string)
    dim buf as string
    dim splt as variant
    dim i as integer
    dim j as integer

    ' (1)テキストファイルを開く(Openステートメント)
    open filename for input as #1
    j = 1
    Do Until EOF(1)
        ' (2)1行分のデータを読み込む(Line Inputステートメント)
        line input #1,buf
        
        ' (3-1)読み込んだデータをセルに代入する
        ' sheets(sheetName).cells(j,1).value = buf

        ' (3-2)読み込んだデータをカンマ毎に分割する
        splt = split(buf,",")
        for i=0 to UBOUND(splt)
            sheets(sheetName).cells(j,i + 1).value = splt(i)
        next i
        j = j + 1
    Loop
    
    ' (4)開いたファイルを閉じる(Closeステートメント)
    Close #1
end function

' LFの読み込み
function csvFileReadLfFunc(ByRef filename as string)
    dim buf as variant
    dim splt as variant
    dim splt2 as variant
    dim i as integer
    dim j as integer

    ' ファイルの読み込み
    open filename for input as #1
        line input #1,buf
    Close #1

    splt = split(buf,vblf)
    for i = 0 to UBOUND(splt)
        splt2 = split(splt(i),",")
        for j = 0 to UBOUND(splt2)
            sheets(sheetName).cells(i + 1,j + 1).value = splt2(j)
        next j
    next i 
end function