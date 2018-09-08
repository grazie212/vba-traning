Attribute VB_Name = "csvFileRead"

Option Explicit

' ��Ɨp�V�[�g
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

' crlf�̓ǂݍ���
function csvFileReadFunc(ByRef filename as string)
    dim buf as string
    dim splt as variant
    dim i as integer
    dim j as integer

    ' (1)�e�L�X�g�t�@�C�����J��(Open�X�e�[�g�����g)
    open filename for input as #1
    j = 1
    Do Until EOF(1)
        ' (2)1�s���̃f�[�^��ǂݍ���(Line Input�X�e�[�g�����g)
        line input #1,buf
        
        ' (3-1)�ǂݍ��񂾃f�[�^���Z���ɑ������
        ' sheets(sheetName).cells(j,1).value = buf

        ' (3-2)�ǂݍ��񂾃f�[�^���J���}���ɕ�������
        splt = split(buf,",")
        for i=0 to UBOUND(splt)
            sheets(sheetName).cells(j,i + 1).value = splt(i)
        next i
        j = j + 1
    Loop
    
    ' (4)�J�����t�@�C�������(Close�X�e�[�g�����g)
    Close #1
end function

' LF�̓ǂݍ���
function csvFileReadLfFunc(ByRef filename as string)
    dim buf as variant
    dim splt as variant
    dim splt2 as variant
    dim i as integer
    dim j as integer

    ' �t�@�C���̓ǂݍ���
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