Attribute VB_Name = "sort"

Option Explicit
sub bubbleSort()

    dim num() as variant
    dim i as integer
    num = array(1,4,3,5,6,7,8,2,9)
    call bubbleSortFunc(num())
    for i =0 to UBOUND(num)
        MsgBox num(i)
    next i

end sub

function bubbleSortFunc(byref numArr() as variant)
    dim tmp as variant
    dim i as integer,j as integer

    for i =LBOUND(numArr) to UBOUND(numArr)
        for j=UBOUND(numArr) to i step - 1
            if numArr(i) > numArr(j) then
                tmp = numArr(i)
                numArr(i) = numArr(j)
                numArr(j) = tmp
            end if
        next j
    next i 
end function