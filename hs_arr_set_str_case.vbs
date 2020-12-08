Option Explicit

Sub hs_arr_set_str_case(ByRef arrInput, ByVal caseType)
	'*** History ***********************************************************************************
	' 2020/12/08, BBS:	- First release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Array helper, Change all strings found to target case either lower or upper
	' Arguments
	'	<arr> arrInput:	Target array to be manipulated
	'	<int> caseType: 0: Lower case, 1: Upper case
	'
	'***********************************************************************************************

	On Error Resume Next

	'*** Initialization ****************************************************************************
	Dim idx1, idx2, size_1, size_2

	'*** Operations ********************************************************************************
	'--- Conditioning, 'caseType' ------------------------------------------------------------------
	If Not IsNumeric(CStr(caseType)) Then
		caseType = 0
	Else
		caseType = CInt(CStr(caseType))
	End If
	
	If caseType <> 0 Then
		caseType = 1
	End If

	'--- Case Modification -------------------------------------------------------------------------
	If IsArray(arrInput) Then
		size_1 = UBound(arrInput, 1)
		size_2 = UBound(arrInput, 2)
		Err.Clear

		For idx1 = 0 to size_1
			If IsEmpty(size_2) Then
				If VarType(arrInput(idx1)) = 8 Then
					If caseType = 0 Then
						arrInput(idx1) = LCase(arrInput(idx1))
					Else
						arrInput(idx1) = UCase(arrInput(idx1))
					End If
				ElseIf IsArray(arrInput(idx1)) Then
					Call hs_arr_set_str_case(arrInput(idx1), caseType)
				End If
			Else
				For idx2 = 0 to size_2
					If VarType(arrInput(idx1, idx2)) = 8 Then
						If caseType = 0 Then
							arrInput(idx1, idx2) = LCase(arrInput(idx1, idx2))
						Else
							arrInput(idx1, idx2) = UCase(arrInput(idx1, idx2))
						End If
					ElseIf IsArray(arrInput(idx1, idx2)) Then
						Call hs_arr_set_str_case(arrInput(idx1, idx2), caseType)
					End If
				Next
			End If
		Next

	ElseIf VarType(arrInput) = 8 Then
		If caseType = 0 Then
			arrInput = LCase(arrInput)
		Else
			arrInput = UCase(arrInput)
		End If
	End If

	'--- Release -----------------------------------------------------------------------------------
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Sub
