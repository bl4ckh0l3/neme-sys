<%
Dim prt
prt = Array("this", "array", "organized", "is", "not")
print_array(prt)
arr_sort prt
print_array(prt)

Sub arr_sort (arr)
    Call QuickSort(arr, 0, ubound(arr, 1))
End Sub

Sub SwapRows (ary,row1,row2)
  Dim tempvar
  tempvar = ary(row1)
  ary(row1) = ary(row2)
  ary(row2) = tempvar
End Sub  'SwapRows

Sub QuickSort (vec,loBound,hiBound)
  '==--------------------------------------------------------==
  '== Sort a 1 dimensional array                             ==
  '==                                                        ==
  '== This procedure is adapted from the algorithm given in: ==
  '==    ~ Data Abstractions & Structures using C++ by ~     ==
  '==    ~ Mark Headington and David Riley, pg. 586    ~     ==
  '== Quicksort is the fastest array sorting routine For     ==
  '== unordered arrays.  Its big O is  n log n               ==
  '==                                                        ==
  '== Parameters:                                            ==
  '== vec       - array to be sorted                         ==
  '== loBound and hiBound are simply the upper and lower     ==
  '==   bounds of the array's 1st dimension.  It's probably  ==
  '==   easiest to use the LBound and UBound functions to    ==
  '==   Set these.                                           ==
  '==--------------------------------------------------------==

  Dim pivot,loSwap,hiSwap,temp,counter
  '== Two items to sort
  if hiBound - loBound = 1 then
    if vec(loBound) > vec(hiBound) then
      Call SwapRows(vec,hiBound,loBound)
    End If
  End If

  '== Three or more items to sort
    pivot = vec(int((loBound + hiBound) / 2))
    vec(int((loBound + hiBound) / 2)) = vec(loBound)
    vec(loBound) = pivot

  loSwap = loBound + 1
  hiSwap = hiBound
  Do
    '== Find the right loSwap
    while loSwap < hiSwap and vec(loSwap) <= pivot
      loSwap = loSwap + 1
    wend
    '== Find the right hiSwap
    while vec(hiSwap) > pivot
      hiSwap = hiSwap - 1
    wend
    '== Swap values if loSwap is less then hiSwap
    if loSwap < hiSwap then Call SwapRows(vec,loSwap,hiSwap)


  Loop While loSwap < hiSwap

    vec(loBound) = vec(hiSwap)
    vec(hiSwap) = pivot

  '== Recursively call function .. the beauty of Quicksort
    '== 2 or more items in first section
    if loBound < (hiSwap - 1) then Call QuickSort(vec,loBound,hiSwap-1)
    '== 2 or more items in second section
    if hiSwap + 1 < hibound then Call QuickSort(vec,hiSwap+1,hiBound)

End Sub  'QuickSort

public sub print_array (var)
    call print_r_depth(var, 0)
end sub

public sub print_r_depth (var, depth)
    if depth=0 then
        response.write("<pre>" & Tab(depth))
        response.write(typename(var))
    end if
    if isarray(var) then
        response.write(Tab(depth) & " (<br />")
        dim x
        for x=0 to uBound(var)
            response.write(Tab(depth+1) & "("&x&")")
            call print_r_depth(var(x), depth+2) 
            response.write("<br />")
        next
        response.write(Tab(depth) & ")")
    end if
    select case vartype(var)
    case VBEmpty: 'Uninitialized
    case VBNull: 'Contains no valid data
    case VBDataObject: 'Data access object
    case VBError:
    case VBArray:
    case VBObject:
    case VBVariant:
    case else:
        if vartype(var) < 16 then
            response.write(" => " & var)
        else
            response.write(" - vartype:" & vartype(var) & " depth:" & depth)
        end if
    end select
    if depth=0 then response.write("</pre>") end if
end sub

public function Tab (spaces)
    dim val, x
    val = ""
    for x=1 to spaces
        val=val & "    "
    next
    Tab = val
end function
%>