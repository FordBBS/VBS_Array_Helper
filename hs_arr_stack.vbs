Function hs_arr_stack(ByRef tarValue, ByVal intLevel)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Stack 'tarValue' inside out for 'intLevel' additional level
	'	e.g. tarValue = ("SULEV"), intLevel = 2
	' 		 return ((("SULEV")))
	'
	'		 tarValue = "CONT_BAG", intLevel = 2
	'		 return (("CONT_BAG"))
	'
	'	Argument(s)
	'	<Any>  tarValue, Any type of value to be stack
	'	<Long> intLevel, Amount of level
	'
	'***********************************************************************************************
	
	On Error Resume Next
	hs_arr_stack = tarValue

	'*** Pre-Validation ****************************************************************************
	If Not IsNumeric(intLevel) Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, arrRes(), arrTmp()
	Redim Preserve arrRes(0)

	intLevel = CInt(intLevel)

	'*** Operations ********************************************************************************
	If intLevel > 0 Then arrRes(0) = tarValue

	For cnt1 = 2 to intLevel
		Erase arrTmp
		Redim Preserve arrTmp(0)
		
		arrTmp(0) = arrRes(0)
		arrRes(0) = arrTmp
	Next

	tarValue = arrRes

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
