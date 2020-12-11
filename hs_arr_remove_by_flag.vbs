Option Explicit

Function hs_arr_remove_by_flag(ByRef arrBase, ByVal arrFlag)
	'*** History ***********************************************************************************
	' 2020/12/11, BBS:	- First Release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Array helper, Reduce 'arrBase' based on 'arrFlag' which contains either 0 or 1
	' data
	' e.g.  arrBase = {"DataStr1", 0, "DataStr2", "1", {"SubData", "0"}, "-1", "0.42", 0.99}
	' 		arrFlag = {0, 1, 0, 1, 0, 1, 1, 1}
	'		Reduced = {0, "1", "-1", "0.42", 0.99}
	'
	'***********************************************************************************************

	On Error Resume Next

	'*** Initialization ****************************************************************************
	Dim idx, thisFlag, n_size, n_cnt, arrRes()

	n_size = -1

	'*** Operations ********************************************************************************
	'--- Conditioning ------------------------------------------------------------------------------
	If Not IsArray(arrBase) Then
		arrBase = Array(arrBase)
	End If

	If Not IsArray(arrFlag) Then
		arrFlag = Array(arrFlag)
	End If

	'--- Determine the size of result array --------------------------------------------------------
	For idx = 0 to UBound(arrBase)
		If idx <= UBound(arrFlag) Then
			If IsNumeric(CStr(arrFlag(idx))) Then
				arrFlag(idx) = CInt(CStr(arrFlag(idx)))
			ElseIf LCase(CStr(arrFlag(idx))) = "true" Then
				arrFlag(idx) = 1
			Else
				arrFlag(idx) = 0
			End If
			n_size = n_size + arrFlag(idx)
		Else
			n_size = n_size + UBound(arrBase) - idx
			Exit For
		End If
	Next

	'--- Create Result -----------------------------------------------------------------------------
	Redim Preserve arrRes(n_size)
	n_cnt = 0

	For idx = 0 to UBound(arrBase)
		If idx > UBound(arrFlag) Then
			thisFlag = 1
		Else
			thisFlag = arrFlag(idx)
		End If

		If thisFlag = 1 Then
			arrRes(n_cnt) = arrBase(idx)
			n_cnt = n_cnt + 1
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	arrBase = arrRes
	hs_arr_remove_by_flag = Err.Number 
	Err.Clear
End Function
