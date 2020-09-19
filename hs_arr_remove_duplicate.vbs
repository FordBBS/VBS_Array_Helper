Function hs_arr_remove_duplicate(ByVal arrInput)
	'*** History ***********************************************************************************
	' 2020/09/19, BBS:	- First release, only one column array is acceptable
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Remove duplicate elements out from a single column Array 'arrInput'
	'
	'	Argument(s)
	'	<Array>  arrInput, One column array to be executed
	'
	'***********************************************************************************************
	
	On Error Resume Next
	hs_arr_remove_duplicate = arrInput

	'*** Pre-Validation ****************************************************************************
	If Not IsArray(arrInput) Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim arrRes, strFoundElement, strElem, thisElem, flg_append

	strFoundElement = ""

	'*** Operations ********************************************************************************
	'--- Go through each element -------------------------------------------------------------------
	For Each thisElem in arrInput
		flg_append = True

		If Not IsArray(thisElem) Then
			strElem = CStr(thisElem)

			If InStr(strFoundElement, "%" & strElem & "%") > 0 Then
				flg_append = False
			ElseIf strFoundElement = "" Then
				strFoundElement = "%" & strElem & "%"
			Else
				strFoundElement = strFoundElement & ";%" & strElem & "%" 
			End If
		End If

		If flg_append Then
			Call hs_arr_append(arrRes, thisElem)
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	hs_arr_remove_duplicate = arrRes

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
