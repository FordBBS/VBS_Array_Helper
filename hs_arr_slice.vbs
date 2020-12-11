Option Explicit

Function hs_arr_slice(ByVal arrBase, ByVal idxStart, ByVal idxEnd)
	'*** History ***********************************************************************************
	' 2020/12/11, BBS:	- First Release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Array helper, Get sub array of 'arrBase' based on 'idxStart' and 'idxEnd' where both of them
	' can be provided in negative value which means backward counting
	' Example, arrBase = {0, 1, 2, 3, 4, 5}
	'	(idxStart, idxEnd, Result) -> (0, 2, {0, 1, 2}), (3, 7, {3, 4, 5}), (3, 4, {3, 4})
	'								  (-1, -2, {5, 4}), (-2, -1, {4, 5}), (-2, -4, {4, 3, 2})
	'
	' Possible Return Value
	'	<array> Sub array from 'arrBase'
	'	<Null>  If 'arrBase' isn't Array or 'idxStart' and 'idxEnd' are both invalid
	'  
	'***********************************************************************************************

	On Error Resume Next
	hs_arr_slice = Empty

	'*** Pre-Validation ****************************************************************************
	If Not IsArray(arrBase) Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim n_size, cnt, thisStep, arrRet, arrIdx(1)
	n_size 	  = UBound(arrBase)
	arrIdx(0) = idxStart
	arrIdx(1) = idxEnd

	'*** Operations ********************************************************************************
	'--- Conditioning, Indices ---------------------------------------------------------------------
	For cnt = 0 to UBound(arrIdx)
		If arrIdx(cnt) < 0 Then
			If Abs(arrIdx(cnt)) > (1 + n_size) Then
				arrIdx(cnt) = -1*n_size
			End If

			arrIdx(cnt) = n_size + arrIdx(cnt) + 1
		End If

		If arrIdx(cnt) > n_size Then
			arrIdx(cnt) = n_size
		End If
	Next

	'--- Slicing -----------------------------------------------------------------------------------
	n_size = arrIdx(0) - arrIdx(1)
	Redim arrRet(Abs(n_size))

	If n_size > 0 Then
		thisStep = -1
	Else
		thisStep = 1
	End If

	For cnt = arrIdx(0) to arrIdx(1) Step thisStep
		If IsObject(arrBase(cnt)) Then
			Set arrRet(Abs(cnt - arrIdx(0))) = arrBase(cnt)
		Else
			arrRet(Abs(cnt - arrIdx(0))) = arrBase(cnt)
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	hs_arr_slice = arrRet

	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
