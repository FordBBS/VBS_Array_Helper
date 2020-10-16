Function hs_arr_to_str(ByVal arrInput)
	'*** History ***********************************************************************************
	' 2020/10/15, BBS:	- First Release, completely join elements in array into string
	'
	'***********************************************************************************************

	On Error Resume Next
	hs_arr_to_str = "<invalid>"

	'*** Pre-Validation ****************************************************************************
	If Not IsArray(arrInput) Then
		hs_arr_to_str = CStr(arrInput)
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim idx, strCombined, tmpVal

	strCombined = ""

	'*** Operations ********************************************************************************
	'--- Join all elements -------------------------------------------------------------------------
	For idx = 0 to UBound(arrInput)
		tmpVal = hs_arr_to_str(arrInput(idx))

		If strCombined = "" Then
			strCombined = tmpVal
		Else
			strCombined = strCombined & ";" & tmpVal
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	hs_arr_to_str = "[" & strCombined & "]"

	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function