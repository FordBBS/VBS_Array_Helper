Function hs_arr_append(ByRef arrInput, ByVal tarValue)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First release
	' 2020/08/25, BBS:  - Implemented handler for Non-Array 'arrInput'
	' 2020/09/19, BBS: 	- Improved
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Append 'tarValue' to target array provided as 'arrInput', 'arrInput' can be only a single
	'	column array only
	'
	'	Argument(s)
	'	<Array>  arrInput, Base array to be appended 'tarValue'
	'	<Any> 	 tarValue, Desire value to be appended to 'arrInput'
	'
	'***********************************************************************************************
	
	On Error Resume Next

	'*** Initialization ****************************************************************************
	' Nothing to be initialized

	'*** Operations ********************************************************************************
	'--- Ensure 'arrInput' is Array type before doing appending ------------------------------------
	If Not IsArray(arrInput) Then
		arrInput = Array(arrInput)
	End If

	'--- Appending ---------------------------------------------------------------------------------
	If Not (UBound(arrInput) = 0 and LCase(TypeName(arrInput(0))) = "empty") Then
		Redim Preserve arrInput(UBound(arrInput) + 1)
	End If

	arrInput(UBound(arrInput)) = tarValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function