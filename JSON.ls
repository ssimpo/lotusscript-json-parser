Private Class JSONScalar
%REM
	Class to contain scalar values processed by JSON class.
	
	@private
	@author		Stephen Simpson <me@simpo.org>
	@version	1.0.0	
%END REM
	Public data As Variant
	Public valueType As String
	
	Public Sub New(value As Variant)
		On Error GoTo ERR_HDL
		
		Select Case TypeName(value)
			Case "STRING":
				Me.valueType = "STRING"
			Case "BOOLEAN":
				Me.valueType = "BOOLEAN"
			Case "NULL":
				Me.valueType = "NULL"
			Case "INTEGER":
				Me.valueType = "INTEGER"
			Case "LONG":
				Me.valueType = "INTEGER"
			Case "SINGLE":
				Me.valueType = "FLOAT"
			Case "DOUBLE":
				Me.valueType = "FLOAT"
			Case Else:
				Me.valueType = "STRING"
		End Select
		
		Me.data = value
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Sub
	
	Public Sub converToNumber()
	%REM
		Convert the current data property to a number. Detect whether
		to convert to float or integer type.  Floats are equivilant to
		Lotusscript Double and integers to Lotusscript Longs.
	%END REM
		If (contains(Me.data, ".")) Then
			Me.data = CDbl(Me.data)
			Me.valueType = "FLOAT"
		Else
			Me.data = CLng(Me.data)
			Me.valueType = "INTEGER"
		End If
	End Sub
	
	Private Function contains (txt1 As Variant,txt2 As Variant) As Boolean
	%REM
		Is one item contained within another
	
		Convert both paramters into strings, lower-case them, then compare to
		see if txt2 can be found within txt2.
	
		@public
		@param Variant txt1	Item to compare against.
		@param Variant txt2	Item to compare with.
		@return Boolean		Is txt2 present within txt1?
	%END REM
		On Error GoTo ERR_HDL
		
		Dim s1 As String, s2 As String
		
		s1 = LCase(CStr(txt1))
		s2 = LCase(CStr(txt2))
		
		If (InStr(s1,s2) > 0) Then contains=True Else contains=False
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Sub reportError(subName As String, errNum As Long, lineNo As Long)
	%REM
		Report an error.  This method is designed for editing to whatever error
		processing is required.  In basic form, it will just fire a message box.
		
		@private
		@param String subName	The method name that fired the error.
		@param errNum Long		The error number.
		@param lineNo Long		The line number that the error occured at.
	%END REM
		MsgBox "Error [" & CStr(errNum) & ": Line no. " & CStr(lineNo) & "] in Class Instance (JSONScalar) - " & subName & Chr$(13) & Error$(errNum)
	End Sub
End Class
Private Class JSONObject
%REM
	Class to contain object values processed via JSON class.
	
	@private
	@author		Stephen Simpson <me@simpo.org>
	@version	1.0.0
%END REM
	Private parser As JSON
	Public valueType As String
	Public data As Variant
	
	Public Sub New(txt As String)
		On Error GoTo ERR_HDL
		
		Me.valueType = "OBJECT"
		Set Me.parser = New JSON(txt)
		Me.data = Me.parser.data
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Sub
	
	Private Sub reportError(subName As String, errNum As Long, lineNo As Long)
	%REM
		Report an error.  This method is designed for editing to whatever error
		processing is required.  In basic form, it will just fire a message box.
		
		@private
		@param String subName	The method name that fired the error.
		@param errNum Long		The error number.
		@param lineNo Long		The line number that the error occured at.
	%END REM
		MsgBox "Error [" & CStr(errNum) & ": Line no. " & CStr(lineNo) & "] in Class Instance (JSONObject) - " & subName & Chr$(13) & Error$(errNum)
	End Sub
End Class
Private Class JSONArray
%REM
	Class to contain array values processed by JSON class.
	
	@private
	@author		Stephen Simpson <me@simpo.org>
	@version	1.0.0	
%END REM
	Public data() As Variant
	Private parser As JSON
	Public valueType As String
	
	Public Sub New(txt As String)
		On Error GoTo ERR_HDL
		
		Dim count As long
		
		Me.valueType = "ARRAY"
		Set Me.parser = New JSON(txt)
		
		count = 0
		ForAll value In Me.parser.data
			count = count + 1
		End ForAll
		ReDim Me.data(1 To count)
		count = 1
		ForAll value In Me.parser.data
			Set Me.data(count) = value
			count = count + 1
		End ForAll
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Sub
	
	Private Sub reportError(subName As String, errNum As Long, lineNo As Long)
	%REM
		Report an error.  This method is designed for editing to whatever error
		processing is required.  In basic form, it will just fire a message box.
		
		@private
		@param String subName	The method name that fired the error.
		@param errNum Long		The error number.
		@param lineNo Long		The line number that the error occured at.
	%END REM
		MsgBox "Error [" & CStr(errNum) & ": Line no. " & CStr(lineNo) & "] in Class Instance (JSONArray) - " & subName & Chr$(13) & Error$(errNum)
	End Sub
End Class
Public Class JSONBase
%REM
	Base Class containing generic functions used in the JSON class.  This
	class is used as a store/collection-point for these methods.
	
	@public
	@author		Stephen Simpson <me@simpo.org>
	@version	1.0.0
%END REM

	Private Function rndLong (x As Long, y As Long) As Long
	%REM
		Generate a random whole number between a range.
	
		@private
		@param Long x	Start of range.
		@param Long y	End of range.
		@return Long	Random number.
	%END REM
		On Error GoTo ERR_HDL
		
		rndLong = Int(( y - x )*Rnd() + x)
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Sub reportError(subName As String, errNum As Long, lineNo As Long)
	%REM
		Report an error.  This method is designed for editing to whatever error
		processing is required.  In basic form, it will just fire a message box.
		
		@private
		@param String subName	The method name that fired the error.
		@param errNum Long		The error number.
		@param lineNo Long		The line number that the error occured at.
	%END REM
		MsgBox "Error [" & CStr(errNum) & ": Line no. " & CStr(lineNo) & "] in Class Instance (JSONBase) - " & subName & Chr$(13) & Error$(errNum)
	End Sub
	
	Private Function isEqual(txt1 As Variant,txt2 As Variant) As Boolean
	%REM
		Are two items equal/equivilant?
	
		Convert both paramters into strings, trim and lower-case them, then
		compare them to see if they are equal.
	
		@private
		@param Variant txt1	Item to compare against.
		@param Variant txt2	Item to compare with.
		@return Boolean		Are they equal/equivilant or not?
	%END REM
		On Error GoTo ERR_HDL
		
		If LCT(txt1) = LCT(txt2) Then isEqual = True Else isEqual = False
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Public Function LCT(txt As Variant) As String
	%REM
		Lower-case, trim and convert to a string the supplied argument.
	
		@private
		@param Variant txt	The text to parse.
		@return String		The converted text.
	%END REM
		On Error GoTo ERR_HDL
		
		LCT = LCase(Trim(CStr(txt)))
		
ERR_HDL:
		If Err<>0 Then
			LCT = ""
			Resume Next
		End If
	End Function
	
	Private Function createRandomCharString(length As Integer) As String
	%REM
		Create a random character string of length defined via parametres. The
		String will contain only uppercase letters or numbers.
		
		@private
		@param Integer length	The length of the random character string to create.
		@return String			The random string created.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim i As Integer, cRndNum As Long
		
		For i = 1 To length
			cRndNum = rndLong(0,36)
			If(cRndNum > 9) Then 'Ascii uppercase latin
				cRndNum = cRndNum - 10 + 65
			Else ' Numerals
				cRndNum = cRndNum + 48
			End If
			
			createRandomCharString = createRandomCharString & Chr$(cRndNum)
		Next
		
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Function LRChop(txt As String) As String
	%REM
		Chop off the first and last characters of string and return the
		modified string.
		
		@private
		@param String txt	String to modify.
		@return String		Modified String.
	%END REM
		LRChop = LChop(RChop(txt))
	End Function
	
	Private Function LChop(txt As String) As String
	%REM
		Chop off the first and character of string and return the
		modified string.
		
		@private
		@param String txt	String to modify.
		@return String		Modified String.
	%END REM
		LChop = Right$(txt, Len(txt)-1)
	End Function
	
	Private Function RChop(txt As String) As String
	%REM
		Chop off the last character of string and return the
		modified string.
		
		@private
		@param String txt	String to modify.
		@return String		Modified String.
	%END REM
		RChop = Left$(txt, Len(txt)-1)
	End Function
	
	Private Function testChar(char As String, tests As String) As Boolean
	%REM
		Test the supplied character for any of the charachers in a test string.
		
		@private
		@param char String	Character to test (NB: should be length: 1).
		@paran tests String	Characters to test for.
		@return Boolean		Did "char" match any of the characters in "tests"?
	%END REM
		On Error GoTo ERR_HDL
		
		Dim chars() As String
		ReDim chars(1 To Len(tests))
		Dim i As Integer
		
		For i = 1 To Len(tests)
			chars(i) = Mid$(tests,i, 1)
		Next
		
		testChar = False
		For i = LBound(chars) To UBound(chars)
			If isEqual(chars(i),char) Then
				testChar = True
			End If
		Next
		
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
End Class
Public Class JSON As JSONBase
%REM
	JSON Parsing class for Lotusscript.  Will take a JSON formatted string and output
	as LS objects/arrays.  Processed data is added to the "data" property, which is
	a Lotusscript List.  The keys are the JSON object keys and the values are either:
	JSONObject, JSONArray or JSONScalar types.  Each of these special types have a "data"
	property containing the actual value.

	The valueType property indicates the actual type as: STRING|OBJECT|ARRAY|FLOAT|
	INTEGER|NULL|BOOLEAN.  If the type is "ARRAY" then the data attribute is a normal
	Lotusscript Array.  If it "OBJECT" then it is a Lotusscript List.  All other items
	are native Lotusscript types of STRING|INTEGER|BOOLEAN|NULL|DOUBLE.
	
	
	@public
	@author		Stephen Simpson <me@simpo.org>
	@version	1.0.0
	
	@param String txt	The JSON string to format.
	
	
	@todo:		Add code to check for badly formatted JSON and report errors
	@todo:		Refector some of the longer methods into muliple
				single-purpose methods.
	@todo:		Re-work the slash processing code so it works with all slashed
				content (even replacing \n for example).
	@todo:		Add stringify methods. 
	@todo:		Class will process JSON if the base object is an object but not
				if it is an array.  This needes fixing.
%END REM

	Public data List As Variant
	
	Private pos As Long
	Private mode As String
	Private count As Long
	
	Private DQRP As String
	Private SQRP As String
	Private CBRP1 As String
	Private SBRP1 As String
	Private CBRP2 As String
	Private SBRP2 As String
	
	Private Function addPlaceHolds(txt As String) As String
	%REM
		Replace certain slashed text with specfic random strings so the
		text can be parsed without functional text interferring. Replaced
		text can be converted back to original format via removePlaceHolds().
		
		@private
		@todo	Refactor to create variables for slashed content found and not to use
				class properties (or be more efficient in doing so). 
		@param String txt	The text to do a search and replace on.
		@return String		Supplied string with content replaced.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim searcher As Variant, replacer As Variant
		
		searcher = Split(|\" \' \[ \{ \] \}|)
		replacer = Split(Me.DQRP & " " & Me.SQRP & " " & Me.SBRP1 & " " & Me.CBRP1 & " " & Me.SBRP2 & " " & Me.CBRP2)
		
		addPlaceHolds = Replace(txt, searcher, replacer)
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Function removePlaceHolds(txt As String) As String
	%REM
		Replace certain character strings with slashed text. This is the
		opposite to addPlaceHolds() and is meant to reverse this process.
		
		@private
		@todo	Refactor to not use class properties (or be more efficient in doing so). 
		@param String txt	The text to do a search and replace on.
		@return String		Supplied string with content replaced.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim searcher As Variant, replacer As Variant
		
		searcher = Split(Me.DQRP & " " & Me.SQRP & " " & Me.SBRP1 & " " & Me.CBRP1 & " " & Me.SBRP2 & " " & Me.CBRP2)
		replacer = Split(|\" \' \[ \{ \] \}|)
		
		removePlaceHolds = Replace(txt, searcher, replacer)
			
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Function removeSlashStrings(txt As String) As String
	%REM
		Remove certain slashed content, replacing with non-slashed
		equivilant text.
		
		@private
		@todo	Change so all slashed text is unslashed. 
		@param String txt	Text to unslash.
		@return String		Unslashed text.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim searcher As Variant, replacer As Variant
		
		searcher = Split(|\" \' \[ \{ \] \}|)
		replacer = Split(|" ' [ { ] }|)
		
		
		removeSlashStrings = Replace(txt, searcher, replacer)
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private sub setMode(txt)
	%REM
		Set the current parsing mode for this instance of the parser.  If we are
		parsing an object we are looking for key/value pairs.  However, if it is
		an array then we are looking for a series of comma seperated values.  Hence,
		the parsing is different for these two types.  Method detects this type and
		set's it on the class property: "mode".  Mode is set to
		either "ARRAY" or "OBJECT".
		
		@private
		@param String txt	The current text being parsed.
	%END REM
		If Left$(Trim(txt), 1) = "[" Then Me.mode = "ARRAY" Else Me.mode = "OBJECT"
	End Sub
	
	Public Sub New(txt As String)
		On Error GoTo ERR_HDL
		
		Dim breakout As Long
	
		Me.DQRP = createRandomCharString(14)
		Me.SQRP = createRandomCharString(14)
		Me.SBRP1 = createRandomCharString(14)
		Me.CBRP1 = createRandomCharString(14)
		Me.SBRP2 = createRandomCharString(14)
		Me.CBRP2 = createRandomCharString(14)
		
		txt = Me.addPlaceHolds(txt)
	
		Me.pos = 1
		Me.count = 1
		breakout = 10
		
		Call setMode(txt)
		If Me.mode = "ARRAY" Then txt = LRChop(Trim(txt))
		While ((Me.pos <= Len(txt)) And (breakout > 0))
			Call parseNextItem(txt)
			breakout = breakout - 1
		Wend
		MsgBox "THE END."
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Sub
	
	Private Function getItemId(txt As String) As String
	%REM
		Get the next item Id in a JSON string, this will be the text that
		appears before the colon.  The ID may or may not be quoted in either
		a single or double quote.
		
		@note	Method is part of the workflow so the class
				property value: pos, will be advance before returning the ID.
		
		@private
		@param String txt	JSON string to parse for next ID.
		@return String		The unquoted ID.		
	%END REM
		On Error GoTo ERR_HDL
		
		Dim char As String
		Dim i As Long
		Dim quoteChar As String
		Dim fieldTokenStart As Long, fieldTokenEnd As Long
		Dim valueTokenStart As Long, valueTokenEnd As Long
		
		i = Me.pos
		Do
			char = Mid$(txt,i, 1)
			i = i + 1
		Loop Until ((Not testChar(char, |{[, |+Chr$(9)+Chr$(13)+Chr$(10))) or ((i > Len(txt))))
		If (i > Len(txt)) Then
			getItemId = ""
			Me.pos = i
			Exit function
		End If
		
		If Not testChar(char, |"'|) Then
			quoteChar = ":"
			fieldTokenStart = i -1
		Else
			fieldTokenStart = i
			quoteChar = char
		End If
		
		Do
			char = Mid$(txt,i, 1)
			i = i + 1
		Loop Until ((testChar(char, quoteChar)) Or ((i > Len(txt))))
		fieldTokenEnd = i - 1
		
		getItemId = Mid$(txt, fieldTokenStart, fieldTokenEnd-fieldTokenStart)
		If Not testChar(quoteChar, |"'|) Then getItemId = Trim(getItemId)
		getItemId = Me.removePlaceHolds(getItemId)
		getItemId = Me.removeSlashStrings(getItemId)
		
		Me.pos = i
		
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Sub parseNextItem(txt As String)
	%REM
		Parse the next item in the supplied JSON string.  Parsing 
		procedure will depend on whether the parser is in ARRAY or OBJECT
		mode.  The class property: data is set as a result of parsing.
		
		@note	Method is part of the workflow so the class
				property value: pos, will be advance before returning the item.
		
		@private
		@param String txt	JSON string to parse for next item.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim id As Variant
		Dim value As Variant
		
		If Me.mode = "OBJECT" Then
			id = getItemId(txt)
			If (Me.pos > Len(txt)) Then Exit Sub
			Set value = getItemValue(txt)
		Else
			id = Me.count
			Set value = getItemValue(txt)
			If (Me.pos > Len(txt)) Then Exit Sub
			Me.count = Me.count + 1
		End If
		Set Me.data(id) = value
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Sub
	
	Private Function getItemValue(txt As String) As Variant
	%REM
		Parse for the next item value in the string.  Will detect what
		type the value is and parse accordingly.  Return value will vary
		depending on detected type.
		
		@note	Method is part of the workflow so the class
				property value: pos, may be advanced before returning the value.
				
		@private
		@param String txt	JSON string to parse for next value.
		@return JSONObject|JSONArray|JSONScalar	Parsed data.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim char As String, longChar4 As String, longChar5 As String
		Dim i As Long
		Dim quoteChar As String
		Dim fieldTokenStart As Long, fieldTokenEnd As Long
		Dim valueTokenStart As Long, valueTokenEnd As Long
		
		i = Me.pos
		Do
			char = Mid$(txt,i, 1)
			longChar4 = LCase(Mid$(txt,i, 4))
			longChar5 = LCase(Mid$(txt,i, 5))
			i = i + 1
		Loop Until testChar(char, |"'{[|) Or IsNumeric(char) Or (i > Len(txt)) Or _
			longChar4 = "true" Or longChar4 = "null" Or longChar5 = "false"
		
		If (i > Len(txt)) Then
			getItemValue = ""
			Me.pos = i
			Exit Function
		End If
		
		Me.pos = i - 1
		Select Case longChar5
			Case "false":
				Me.pos = Me.pos + 5
				Set getItemValue = getItemBoolean(false)
				MsgBox getItemValue.data
			Case Else:
				Select Case longChar4
					Case "true":
					Me.pos = Me.pos + 4
						Set getItemValue = getItemBoolean(true)
						MsgBox getItemValue.data
					Case "null":
						Me.pos = Me.pos + 4
						Set getItemValue = getItemNull(Null)
						MsgBox "Null"
					Case Else:
						Select Case char
							Case |'|, |"|:
								Set getItemValue = getItemString(txt)
								getItemValue.data = Me.removePlaceHolds(CStr(getItemValue.data))
								getItemValue.data = Me.removeSlashStrings(CStr(getItemValue.data))
								MsgBox getItemValue.data
							Case |{|: Set getItemValue = getItemObject(txt)
							Case |[|: Set getItemValue = getItemArray(txt)
							Case Else:
								If IsNumeric(char) Then
									Set getItemValue = getItemNumber(txt)
									MsgBox getItemValue.data
								End If
						End Select
				End Select
		End Select
		
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Function getItemArray(txt As String) As Variant
	%REM
		Get the next JSON Array from the supplied text and return as a
		JSONArray object.
		
		@note	Method is part of the workflow so the class
				property value: pos, will be advance before returning the array.
		
		@private
		@param String txt	JSON string to parse for next array.
		@return JSONArray	The array.
	%END REM	
		On Error GoTo ERR_HDL
		
		Dim i As Long, objectLevel As Integer
		Dim char As String
		Dim value As String
		
		i = pos
		Do 
			char = Mid$(txt,i, 1)
			If char = "[" Then objectLevel = objectLevel + 1
			If char = "]" Then objectLevel = objectLevel - 1
			i = i + 1
		Loop While ((objectLevel > 0) And i <= Len(Txt))
		value = removePlaceHolds(Mid$(txt, Me.pos, i - Me.pos))
		Set getItemArray = New JSONArray(value)
		Me.pos = i + 1
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Function getItemObject(txt As String) As Variant
	%REM
		Get the next JSON Object from the supplied text and return as a
		JSONObject object.
		
		@note	Method is part of the workflow so the class
				property value: pos, will be advance before returning the object.
		
		@private
		@param String txt	JSON string to parse for next object.
		@return JSONObject	The object.
	%END REM	
		On Error GoTo ERR_HDL
		
		Dim i As Long, objectLevel As Integer
		Dim char As String
		Dim value As String
		
		i = Me.pos
		Do 
			char = Mid$(txt,i, 1)
			If char = "{" Then objectLevel = objectLevel + 1
			If char = "}" Then objectLevel = objectLevel - 1
			i = i + 1
		Loop While ((objectLevel > 0) And i <= Len(Txt))
		value = removePlaceHolds(Mid$(txt, Me.pos, i - Me.pos))
		Set getItemObject = New JSONObject(value)
		Me.pos = i + 1
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Function getItemNull(value As Variant) As Variant
	%REM
		Convert the supplied value to JSONScaler of value = Null.
		
		@private
		@param value Variant	This assumes that it is a text value of "null".
		@return JSONScalar		A JSONSclar of value = Null.
	%END REM
		On Error GoTo ERR_HDL
		
		Set getItemNull = New JSONScalar(value)
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Function getItemBoolean(value As Boolean) As Variant
	%REM
		Convert the supplied value to JSONScaler of type boolean.
		
		@private
		@param value Variant	This assumes that it is a text value of "true" or "false".
		@return JSONScalar		A JSONSclar of type boolean.
	%END REM
		On Error GoTo ERR_HDL
		
		Set getItemBoolean = New JSONScalar(value)
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function

	Private Function getItemNumber(txt As String) As Variant
	%REM
		Convert the supplied value to JSONScaler of type float or interger.
		Will dectect actual type based on presence of a decimal point.
		
		@private
		@param value Variant	This assumes that it is a numeric text value.
		@return JSONScalar		A JSONSclar of type float or integer.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim i As Long
		Dim char As String
		Dim value As String
		
		i = Me.pos
		Do
			char = Mid$(txt,i, 1)
			i = i + 1
		Loop Until ((testChar(char, | ,}]|+Chr$(9)+Chr$(10)+Chr$(13))) or (i > Len(txt)))
		value = Mid$(txt, Me.pos, i - Me.pos - 1)
		Set getItemNumber = New JSONScalar(value)
		Call getItemNumber.converToNumber()
		Me.pos = i
		
ERR_HDL:
		If Err<>0 Then Call reportError(GetThreadInfo(1),Err(),Erl())
	End Function

	Private Function getItemString(txt As String) As Variant
	%REM
		Convert the supplied value to JSONScaler of type string.
		
		@private
		@param value Variant	This assumes that it is a non-specfic text value.
		@return JSONScalar		A JSONSclar of type string.
	%END REM
		On Error GoTo ERR_HDL
		
		Dim char As String
		Dim i As Long
		Dim quoteChar As String
		Dim fieldTokenStart As Long, fieldTokenEnd As Long
		Dim valueTokenStart As Long, valueTokenEnd As Long
		Dim value As String
		
		i = Me.pos
		Do
			char = Mid$(txt,i, 1)
			i = i + 1
		Loop Until ((testChar(char, |"'|)) Or ((i > Len(txt))))
		fieldTokenStart = i
		quoteChar = char
		
		Do
			char = Mid$(txt,i, 1)
			i = i + 1
		Loop Until ((testChar(char, quoteChar)) Or (i > Len(txt)))
		fieldTokenEnd = i - 1
		
		value = Mid$(txt, fieldTokenStart, fieldTokenEnd-fieldTokenStart)
		Set getItemString = New JSONScalar(value)
		Me.pos = i
		
		
ERR_HDL:
		If Err<>0 Then call reportError(GetThreadInfo(1),Err(),Erl())
	End Function
	
	Private Sub reportError(subName As String, errNum As Long, lineNo As Long)
	%REM
		Report an error.  This method is designed for editing to whatever error
		processing is required.  In basic form, it will just fire a message box.
		
		@private
		@param String subName	The method name that fired the error.
		@param errNum Long		The error number.
		@param lineNo Long		The line number that the error occured at.
	%END REM
		MsgBox "Error [" & CStr(errNum) & ": Line no. " & CStr(lineNo) & "] in Class Instance (JSON) - " & subName & Chr$(13) & Error$(errNum)
	End Sub
End Class


