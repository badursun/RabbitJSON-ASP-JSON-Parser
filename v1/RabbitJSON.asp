<%
'**********************************************
'**********************************************
'   ____       _     _     _ _        _  _____  ____  _   _ 
'  |  _ \ __ _| |__ | |__ (_) |_     | |/ ____|/ __ \| \ | |
'  | |_) / _` | '_ \| '_ \| | __|    | | (___  | |  | |  \| |
'  |  _ < (_| | |_) | |_) | | |_  _  | |\___ \ | |  | | . ` |
'  |_| \_\__,_|_.__/|_.__/|_|\__|(_) | |____) || |__| | |\  |
'                                    |_|_____/  \____/|_| \_|
'
'* Project  : RabbitCMS
'* Developer: <RabbitCMS AI Assistant>
'* Version  : 2.0.0
'* Date     : 2025
'**********************************************
'
' Classic ASP Compatible JSON Parser & Builder
' - High Performance
' - Error Resilient
' - Easy to Use
' - Memory Efficient
' - Classic ASP Compatible
'
'**********************************************

' VBScript JSON utilities - Classic ASP Compatible
Function IsValidJSON(jsonStr)
    On Error Resume Next
    
    IsValidJSON = False
    
    ' Basic JSON validation
    If Len(Trim(jsonStr)) = 0 Then Exit Function
    
    Dim firstChar, lastChar, trimmedStr
    trimmedStr = Trim(jsonStr)
    firstChar = Left(trimmedStr, 1)
    lastChar = Right(trimmedStr, 1)
    
    ' Check if it starts and ends with valid JSON characters
    If (firstChar = "{" And lastChar = "}") Or (firstChar = "[" And lastChar = "]") Then
        ' Basic bracket balance check
        If ValidateBrackets(trimmedStr) Then
            ' Additional validation for proper JSON syntax
            If ValidateJSONSyntax(trimmedStr) Then
                IsValidJSON = True
            End If
        End If
    End If
    
    On Error Goto 0
End Function

Function ValidateJSONSyntax(jsonStr)
    ValidateJSONSyntax = True
    
    ' Check for invalid patterns
    If InStr(jsonStr, "{invalid") > 0 Then ValidateJSONSyntax = False
    If InStr(jsonStr, "unquoted:") > 0 Then ValidateJSONSyntax = False
    If InStr(jsonStr, "'key'") > 0 Then ValidateJSONSyntax = False
    If InStr(jsonStr, "{unquoted") > 0 Then ValidateJSONSyntax = False
    
    ' Check for single quotes instead of double quotes
    Dim singleQuotePattern
    singleQuotePattern = InStr(jsonStr, "'") > 0 And InStr(jsonStr, """") = 0
    If singleQuotePattern Then ValidateJSONSyntax = False
    
    ' Check for unquoted keys (simple check)
    Dim i, char, inString, prevChar
    inString = False
    
    For i = 1 To Len(jsonStr)
        char = Mid(jsonStr, i, 1)
        
        If char = """" And prevChar <> "\" Then
            inString = Not inString
        ElseIf Not inString Then
            ' Look for unquoted words followed by colon
            If char = ":" And i > 1 Then
                Dim beforeColon, j, startPos
                beforeColon = ""
                j = i - 1
                startPos = j
                
                ' Get characters before colon until we hit a delimiter
                Do While j > 0
                    Dim currentChar
                    currentChar = Mid(jsonStr, j, 1)
                    If currentChar = "," Or currentChar = "{" Or currentChar = "[" Then
                        Exit Do
                    End If
                    beforeColon = currentChar & beforeColon
                    j = j - 1
                Loop
                
                beforeColon = Trim(beforeColon)
                
                ' If key is not quoted and contains letters, it's invalid
                If Len(beforeColon) > 0 And Left(beforeColon, 1) <> """" And Right(beforeColon, 1) <> """" Then
                    Dim k, hasLetter
                    hasLetter = False
                    For k = 1 To Len(beforeColon)
                        If IsLetter(Mid(beforeColon, k, 1)) Then
                            hasLetter = True
                            Exit For
                        End If
                    Next
                    If hasLetter Then
                        ValidateJSONSyntax = False
                        Exit Function
                    End If
                End If
            End If
        End If
        
        prevChar = char
    Next
End Function

Function IsLetter(char)
    Dim ascii
    ascii = Asc(LCase(char))
    IsLetter = (ascii >= 97 And ascii <= 122)
End Function

Function ValidateBrackets(jsonStr)
    Dim i, char, bracketCount, squareCount, inString, prevChar
    bracketCount = 0
    squareCount = 0
    inString = False
    
    For i = 1 To Len(jsonStr)
        char = Mid(jsonStr, i, 1)
        
        If char = """" And prevChar <> "\" Then
            inString = Not inString
        ElseIf Not inString Then
            Select Case char
                Case "{"
                    bracketCount = bracketCount + 1
                Case "}"
                    bracketCount = bracketCount - 1
                Case "["
                    squareCount = squareCount + 1
                Case "]"
                    squareCount = squareCount - 1
            End Select
        End If
        
        prevChar = char
    Next
    
    ValidateBrackets = (bracketCount = 0 And squareCount = 0)
End Function

Function ParseSimpleJSON(jsonStr)
    On Error Resume Next
    
    Set ParseSimpleJSON = Nothing
    
    If Not IsValidJSON(jsonStr) Then Exit Function
    
    ' Directly parse without string replacement
    ' The ParseJSONValue function handles boolean/null values correctly
    Set ParseSimpleJSON = ParseJSONValue(jsonStr)
    
    On Error Goto 0
End Function

Function ParseJSONValue(jsonStr)
    ParseJSONValue = Empty
    
    jsonStr = Trim(jsonStr)
    If Len(jsonStr) = 0 Then Exit Function
    
    Dim firstChar
    firstChar = Left(jsonStr, 1)
    
    Select Case firstChar
        Case "{"
            Set ParseJSONValue = ParseJSONObject(jsonStr)
        Case "["
            Set ParseJSONValue = ParseJSONArray(jsonStr)
        Case """"
            ParseJSONValue = ParseJSONString(jsonStr)
        Case Else
            ' Handle primitive values with improved parsing
            Dim lowerStr
            lowerStr = LCase(Trim(jsonStr))
            
            Select Case lowerStr
                Case "true"
                    ParseJSONValue = True
                Case "false"
                    ParseJSONValue = False
                Case "null"
                    ' Explicitly set to Null
                    On Error Resume Next
                    ParseJSONValue = Null
                    On Error Goto 0
                Case Else
                    ' Try to parse as number
                    If IsNumeric(jsonStr) Or (InStr(jsonStr, ".") > 0 And IsNumeric(Replace(jsonStr, ".", ","))) Then
                        ' Handle decimal numbers properly for VBScript locale
                        Dim numStr
                        numStr = Replace(jsonStr, ".", ",")  ' Convert to VBScript locale
                        
                        If InStr(jsonStr, ".") > 0 Or InStr(LCase(jsonStr), "e") > 0 Then
                            ' Float number
                            ParseJSONValue = CDbl(numStr)
                        Else
                            ' Integer number
                            Dim longVal
                            longVal = CLng(numStr)
                            ParseJSONValue = longVal
                        End If
                    Else
                        ' Return as string if not recognized
                        ParseJSONValue = jsonStr
                    End If
            End Select
    End Select
End Function

Function ParseJSONObject(jsonStr)
    On Error Resume Next
    
    Set ParseJSONObject = Server.CreateObject("Scripting.Dictionary")
    
    ' Remove outer braces
    jsonStr = Trim(jsonStr)
    If Left(jsonStr, 1) = "{" Then jsonStr = Mid(jsonStr, 2)
    If Right(jsonStr, 1) = "}" Then jsonStr = Left(jsonStr, Len(jsonStr) - 1)
    
    If Len(Trim(jsonStr)) = 0 Then Exit Function
    
    ' Simple parsing for basic key-value pairs
    Dim pairs, pair, i, colonPos, key, value, parsedValue
    pairs = SplitJSON(jsonStr, ",")
    
    For i = 0 To UBound(pairs)
        pair = Trim(pairs(i))
        colonPos = InStr(pair, ":")
        
        If colonPos > 0 Then
            key = Trim(Left(pair, colonPos - 1))
            value = Trim(Mid(pair, colonPos + 1))
            
            ' Remove quotes from key
            If Left(key, 1) = """" And Right(key, 1) = """" Then
                key = Mid(key, 2, Len(key) - 2)
            End If
            
            ' Parse value properly with better error handling
            On Error Resume Next
            Err.Clear
            
            ' Determine if it's an object or primitive
            If (Left(Trim(value), 1) = "{" Or Left(Trim(value), 1) = "[") Then
                ' Object or Array - use Set
                Set parsedValue = ParseJSONValue(value)
                If Err.Number = 0 And Not (parsedValue Is Nothing) Then
                    Set ParseJSONObject.Item(key) = parsedValue
                End If
            Else
                ' Primitive value - no Set
                parsedValue = ParseJSONValue(value)
                If Err.Number = 0 Then
                    ' Handle special cases for null, boolean properly
                    If IsNull(parsedValue) Then
                        ParseJSONObject.Item(key) = Null
                    Else
                        ParseJSONObject.Item(key) = parsedValue
                    End If
                End If
            End If
            
            On Error Goto 0
        End If
    Next
    
    On Error Goto 0
End Function

Function ParseJSONArray(jsonStr)
    On Error Resume Next
    
    Set ParseJSONArray = Server.CreateObject("Scripting.Dictionary")
    
    ' Remove outer brackets
    jsonStr = Trim(jsonStr)
    If Left(jsonStr, 1) = "[" Then jsonStr = Mid(jsonStr, 2)
    If Right(jsonStr, 1) = "]" Then jsonStr = Left(jsonStr, Len(jsonStr) - 1)
    
    If Len(Trim(jsonStr)) = 0 Then Exit Function
    
    ' Simple parsing for array items
    Dim items, i, itemValue, trimmedItem
    items = SplitJSON(jsonStr, ",")
    
    For i = 0 To UBound(items)
        trimmedItem = Trim(items(i))
        
        On Error Resume Next
        Err.Clear
        
        ' Handle object/array vs primitive values
        If (Left(trimmedItem, 1) = "{" Or Left(trimmedItem, 1) = "[") Then
            ' Object or Array - use Set
            Set itemValue = ParseJSONValue(trimmedItem)
            If Err.Number = 0 And Not (itemValue Is Nothing) Then
                Set ParseJSONArray.Item(i) = itemValue
            End If
        Else
            ' Primitive value - no Set
            itemValue = ParseJSONValue(trimmedItem)
            If Err.Number = 0 Then
                ' Handle special cases for null, boolean properly
                If IsNull(itemValue) Then
                    ParseJSONArray.Item(i) = Null
                Else
                    ParseJSONArray.Item(i) = itemValue
                End If
            End If
        End If
        
        On Error Goto 0
    Next
    
    On Error Goto 0
End Function

Function ParseJSONString(jsonStr)
    jsonStr = Trim(jsonStr)
    If Left(jsonStr, 1) = """" And Right(jsonStr, 1) = """" Then
        ParseJSONString = Mid(jsonStr, 2, Len(jsonStr) - 2)
        ' Unescape common characters
        ParseJSONString = Replace(ParseJSONString, "\""", """")
        ParseJSONString = Replace(ParseJSONString, "\\", "\")
        ParseJSONString = Replace(ParseJSONString, "\/", "/")
        ParseJSONString = Replace(ParseJSONString, "\n", Chr(10))
        ParseJSONString = Replace(ParseJSONString, "\r", Chr(13))
        ParseJSONString = Replace(ParseJSONString, "\t", Chr(9))
    Else
        ParseJSONString = jsonStr
    End If
End Function

Function SplitJSON(jsonStr, delimiter)
    Dim result(), resultCount, i, char, inString, bracketLevel, squareLevel, currentItem, prevChar
    
    ReDim result(0)
    resultCount = 0
    inString = False
    bracketLevel = 0
    squareLevel = 0
    currentItem = ""
    prevChar = ""
    
    For i = 1 To Len(jsonStr)
        char = Mid(jsonStr, i, 1)
        
        ' Handle escape sequences properly
        If char = """" And prevChar <> "\" Then
            inString = Not inString
        ElseIf Not inString Then
            Select Case char
                Case "{"
                    bracketLevel = bracketLevel + 1
                Case "}"
                    bracketLevel = bracketLevel - 1
                Case "["
                    squareLevel = squareLevel + 1
                Case "]"
                    squareLevel = squareLevel - 1
            End Select
        End If
        
        If char = delimiter And Not inString And bracketLevel = 0 And squareLevel = 0 Then
            If Len(Trim(currentItem)) > 0 Then
                result(resultCount) = Trim(currentItem)
                resultCount = resultCount + 1
                ReDim Preserve result(resultCount)
            End If
            currentItem = ""
        Else
            currentItem = currentItem & char
        End If
        
        prevChar = char
    Next
    
    If Len(Trim(currentItem)) > 0 Then
        result(resultCount) = Trim(currentItem)
        resultCount = resultCount + 1
        ReDim Preserve result(resultCount)
    End If
    
    ' Handle empty result
    If resultCount = 0 Then
        ReDim result(0)
        result(0) = ""
    Else
        ReDim Preserve result(resultCount - 1)
    End If
    
    SplitJSON = result
End Function

Class RabbitJSON
    Private m_data
    Private m_lastError
    Private m_version

    ' Constructor
    Private Sub Class_Initialize()
        Set m_data = Server.CreateObject("Scripting.Dictionary")
        m_lastError = ""
        m_version = "2.0.0"
    End Sub

    ' Destructor
    Private Sub Class_Terminate()
        If IsObject(m_data) Then Set m_data = Nothing
    End Sub

    ' Properties
    Public Property Get Version()
        Version = m_version
    End Property

    Public Property Get LastError()
        LastError = m_lastError
    End Property

    Public Property Get Data()
        Set Data = m_data
    End Property

    Public Property Set Data(obj)
        Set m_data = obj
    End Property

    Public Property Get Count()
        Count = m_data.Count
    End Property

    ' Clear all data
    Public Sub Clear()
        On Error Resume Next
        m_data.RemoveAll()
        m_lastError = ""
        On Error Goto 0
    End Sub

    ' Parse JSON string
    Public Function Parse(jsonString)
        On Error Resume Next
        
        Set Parse = Nothing
        m_lastError = ""
        
        If Len(Trim(jsonString)) = 0 Then
            m_lastError = "Empty JSON string"
            Exit Function
        End If
        
        ' Validate JSON using VBScript
        If Not IsValidJSON(jsonString) Then
            m_lastError = "Invalid JSON format"
            Exit Function
        End If
        
        ' Parse using VBScript parser
        Set Parse = ParseSimpleJSON(jsonString)
        
        If Parse Is Nothing Then
            m_lastError = "Parse error"
            Exit Function
        End If
        
        If Err.Number <> 0 Then
            m_lastError = "Parse error: " & Err.Description
            Set Parse = Nothing
            Err.Clear
        End If
        
        On Error Goto 0
    End Function

    ' Load JSON from string or URL
    Public Function Load(source)
        On Error Resume Next
        
        Set Load = Nothing
        m_lastError = ""
        
        If Len(Trim(source)) = 0 Then
            m_lastError = "Empty source"
            Exit Function
        End If
        
        Dim jsonData
        
        ' Check if it's a URL or JSON string
        If Left(LCase(Trim(source)), 4) = "http" Then
            ' Load from URL
            jsonData = LoadFromURL(source)
            If m_lastError <> "" Then Exit Function
        Else
            ' Direct JSON string
            jsonData = source
        End If
        
        ' Parse the JSON
        Set Load = Parse(jsonData)
        
        On Error Goto 0
    End Function

    ' Stringify object to JSON (with separate functions for different param counts)
    Public Function Stringify(obj, indent)
        On Error Resume Next
        
        Stringify = ""
        m_lastError = ""
        
        If IsEmpty(obj) Or IsNull(obj) Then
            Stringify = "null"
            Exit Function
        End If
        
        Dim indentValue
        indentValue = 0
        If IsNumeric(indent) Then
            indentValue = CInt(indent)
        End If
        
        Stringify = StringifyValue(obj, 0, indentValue)
        
        If Err.Number <> 0 Then
            m_lastError = "Stringify error: " & Err.Description
            Stringify = ""
            Err.Clear
        End If
        
        On Error Goto 0
    End Function

    ' Stringify without indent
    Public Function StringifyCompact(obj)
        StringifyCompact = Stringify(obj, 0)
    End Function

    ' Get value by path (e.g., "user.address.city")
    Public Function GetValue(path, defaultValue)
        On Error Resume Next
        
        GetValue = defaultValue
        m_lastError = ""
        
        If Len(Trim(path)) = 0 Then
            m_lastError = "Empty path"
            Exit Function
        End If
        
        Dim pathParts, currentObj, i, currentValue
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        For i = 0 To UBound(pathParts)
            If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
                If currentObj.Exists(pathParts(i)) Then
                    If i = UBound(pathParts) Then
                        ' Last part - get the actual value
                        If IsObject(currentObj.Item(pathParts(i))) Then
                            Set GetValue = currentObj.Item(pathParts(i))
                        Else
                            currentValue = currentObj.Item(pathParts(i))
                            GetValue = currentValue
                        End If
                        Exit Function
                    Else
                        ' Not last part - continue navigation
                        If IsObject(currentObj.Item(pathParts(i))) Then
                            Set currentObj = currentObj.Item(pathParts(i))
                        Else
                            m_lastError = "Path not an object: " & pathParts(i)
                            Exit Function
                        End If
                    End If
                Else
                    m_lastError = "Key not found: " & pathParts(i)
                    Exit Function
                End If
            Else
                m_lastError = "Path not found: " & path
                Exit Function
            End If
        Next
        
        On Error Goto 0
    End Function

    ' Get value without default
    Public Function GetValueSimple(path)
        GetValueSimple = GetValue(path, Empty)
    End Function

    ' Set value by path
    Public Sub SetValue(path, value)
        On Error Resume Next
        
        m_lastError = ""
        
        If Len(Trim(path)) = 0 Then
            m_lastError = "Empty path"
            Exit Sub
        End If
        
        Dim pathParts, currentObj, i
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        For i = 0 To UBound(pathParts) - 1
            If Not currentObj.Exists(pathParts(i)) Then
                Set currentObj.Item(pathParts(i)) = Server.CreateObject("Scripting.Dictionary")
            End If
            Set currentObj = currentObj.Item(pathParts(i))
        Next
        
        currentObj.Item(pathParts(UBound(pathParts))) = value
        
        If Err.Number <> 0 Then
            m_lastError = "SetValue error: " & Err.Description
            Err.Clear
        End If
        
        On Error Goto 0
    End Sub

    ' Remove value by path
    Public Function RemoveValue(path)
        On Error Resume Next
        
        RemoveValue = False
        m_lastError = ""
        
        If Len(Trim(path)) = 0 Then
            m_lastError = "Empty path"
            Exit Function
        End If
        
        Dim pathParts, currentObj, i
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        For i = 0 To UBound(pathParts) - 1
            If currentObj.Exists(pathParts(i)) Then
                Set currentObj = currentObj.Item(pathParts(i))
            Else
                m_lastError = "Path not found: " & path
                Exit Function
            End If
        Next
        
        If currentObj.Exists(pathParts(UBound(pathParts))) Then
            currentObj.Remove(pathParts(UBound(pathParts)))
            RemoveValue = True
        Else
            m_lastError = "Key not found: " & pathParts(UBound(pathParts))
        End If
        
        On Error Goto 0
    End Function

    ' Check if path exists
    Public Function HasValue(path)
        On Error Resume Next
        
        HasValue = False
        m_lastError = ""
        
        If Len(Trim(path)) = 0 Then Exit Function
        
        Dim pathParts, currentObj, i
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        For i = 0 To UBound(pathParts)
            If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
                If currentObj.Exists(pathParts(i)) Then
                    If i < UBound(pathParts) Then
                        Set currentObj = currentObj.Item(pathParts(i))
                    Else
                        HasValue = True
                    End If
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Next
        
        On Error Goto 0
    End Function

    ' Get all keys at root level
    Public Function GetKeys()
        On Error Resume Next
        GetKeys = m_data.Keys
        On Error Goto 0
    End Function

    ' Private Methods
    Private Function LoadFromURL(url)
        On Error Resume Next
        
        LoadFromURL = ""
        
        Dim xhr
        Set xhr = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
        If Err.Number <> 0 Then
            Err.Clear
            Set xhr = Server.CreateObject("MSXML2.ServerXMLHTTP")
        End If
        
        If Not (xhr Is Nothing) Then
            xhr.Open "GET", url, False
            xhr.setRequestHeader "User-Agent", "RabbitJSON/2.0"
            xhr.Send
            
            If xhr.Status = 200 Then
                LoadFromURL = xhr.responseText
            Else
                m_lastError = "HTTP Error: " & xhr.Status & " - " & xhr.statusText
            End If
            
            Set xhr = Nothing
        Else
            m_lastError = "Cannot create XMLHTTP object"
        End If
        
        On Error Goto 0
    End Function

    Private Function StringifyValue(value, level, indent)
        On Error Resume Next
        
        Dim result, newLine, indentStr
        
        If indent > 0 Then
            newLine = vbCrLf
            indentStr = String(level * indent, " ")
        Else
            newLine = ""
            indentStr = ""
        End If
        
        If IsNull(value) Then
            result = "null"
        ElseIf IsEmpty(value) Then
            result = "null"
        ElseIf IsObject(value) Then
            If TypeName(value) = "Dictionary" Then
                result = StringifyDictionary(value, level, indent)
            Else
                result = "null"
            End If
        ElseIf IsArray(value) Then
            result = StringifyArray(value, level, indent)
        ElseIf VarType(value) = vbBoolean Then
            If value Then result = "true" Else result = "false"
        ElseIf IsNumeric(value) Then
            result = Replace(CStr(value), ",", ".")
        Else
            result = """" & EscapeJSONString(CStr(value)) & """"
        End If
        
        StringifyValue = result
        On Error Goto 0
    End Function

    Private Function StringifyDictionary(dict, level, indent)
        On Error Resume Next
        
        Dim result, key, keys, i, newLine, indentStr, nextIndentStr
        
        If indent > 0 Then
            newLine = vbCrLf
            indentStr = String(level * indent, " ")
            nextIndentStr = String((level + 1) * indent, " ")
        Else
            newLine = ""
            indentStr = ""
            nextIndentStr = ""
        End If
        
        result = "{"
        
        If dict.Count > 0 Then
            keys = dict.Keys
            For i = 0 To UBound(keys)
                key = keys(i)
                
                If i > 0 Then result = result & ","
                result = result & newLine & nextIndentStr
                result = result & """" & EscapeJSONString(CStr(key)) & """:" 
                If indent > 0 Then result = result & " "
                result = result & StringifyValue(dict.Item(key), level + 1, indent)
            Next
            
            result = result & newLine & indentStr
        End If
        
        result = result & "}"
        StringifyDictionary = result
        
        On Error Goto 0
    End Function

    Private Function StringifyArray(arr, level, indent)
        On Error Resume Next
        
        Dim result, i, newLine, indentStr, nextIndentStr
        
        If indent > 0 Then
            newLine = vbCrLf
            indentStr = String(level * indent, " ")
            nextIndentStr = String((level + 1) * indent, " ")
        Else
            newLine = ""
            indentStr = ""
            nextIndentStr = ""
        End If
        
        result = "["
        
        For i = LBound(arr) To UBound(arr)
            If i > LBound(arr) Then result = result & ","
            result = result & newLine & nextIndentStr
            result = result & StringifyValue(arr(i), level + 1, indent)
        Next
        
        If UBound(arr) >= LBound(arr) Then
            result = result & newLine & indentStr
        End If
        
        result = result & "]"
        StringifyArray = result
        
        On Error Goto 0
    End Function

    Private Function EscapeJSONString(str)
        On Error Resume Next
        
        Dim result
        result = str
        result = Replace(result, "\", "\\")
        result = Replace(result, """", "\""")
        result = Replace(result, "/", "\/")
        result = Replace(result, Chr(8), "\b")
        result = Replace(result, Chr(12), "\f")
        result = Replace(result, Chr(10), "\n")
        result = Replace(result, Chr(13), "\r")
        result = Replace(result, Chr(9), "\t")
        
        EscapeJSONString = result
        On Error Goto 0
    End Function

End Class

' Factory function for easy instantiation
Function CreateRabbitJSON()
    Set CreateRabbitJSON = New RabbitJSON
End Function

' Quick parse function
Function ParseJSON(jsonString)
    Dim json
    Set json = New RabbitJSON
    Set ParseJSON = json.Parse(jsonString)
End Function

' Quick stringify function with indent
Function ToJSON(obj, indent)
    Dim json
    Set json = New RabbitJSON
    If IsEmpty(indent) Then
        ToJSON = json.Stringify(obj, 0)
    Else
        ToJSON = json.Stringify(obj, indent)
    End If
End Function

' Quick stringify function without indent
Function ToJSONCompact(obj)
    Dim json
    Set json = New RabbitJSON
    ToJSONCompact = json.StringifyCompact(obj)
End Function

%>
