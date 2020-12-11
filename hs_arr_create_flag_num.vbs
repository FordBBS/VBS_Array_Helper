Option Explicit

Function hs_arr_create_flag_num(ByVal arrBase)
	'*** History ***********************************************************************************
	' 2020/12/11, BBS:	- First Release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Array helper, Create an array flag based on 'arrBase' to determine which position is a numerical
	' data
	' e.g.  arrBase = {"DataStr1", 0, "DataStr2", "1", {"SubData", "0"}, "-1", "0.42", 0.99}
	' 		Return  = {0, 1, 0, 1, 0, 1, 1, 1}
	'
	'***********************************************************************************************

	On Error Resume Next

	'*** Initialization ****************************************************************************
	Dim idx, flg_val, arrRes()
	Redim Preserve arrRes(UBound(arrBase))

	'*** Operations ********************************************************************************
	'--- Conditioning, 'arrBase' -------------------------------------------------------------------
	If Not IsArray(arrBase) Then
		arrBase = Array(arrBase)
	End If

	'--- Validation --------------------------------------------------------------------------------
	For idx = 0 to UBound(arrRes)
		If IsArray(arrBase(idx)) Then
			flg_val = 0
		ElseIf Not IsNumeric(CStr(arrBase(idx))) or LCase(CStr(arrBase(idx))) = "true" _
			or LCase(CStr(arrBase(idx))) = "false" Then
			flg_val = 0
		Else
			flg_val = 1
		End If
		arrRes(idx) = flg_val
	Next

	'--- Release -----------------------------------------------------------------------------------
	hs_arr_create_flag_num = arrRes
	
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
