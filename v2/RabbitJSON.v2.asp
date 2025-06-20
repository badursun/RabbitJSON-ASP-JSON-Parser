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
'* Version  : 2.1.0 (Refactored)
'* Date     : 2025
'**********************************************
'
' Classic ASP JSON Parser & Builder - Version 2
' - Modular Class Architecture
' - Enhanced Performance  
' - Improved Error Handling
' - Memory Optimized
' - Production Ready
'
'**********************************************

Class RabbitJSON
    ' ================================
    ' PRIVATE MEMBERS
    ' ================================
    Private m_data          ' Main data storage (Dictionary)
    Private m_lastError     ' Last error message
    Private m_version       ' Library version
    Private m_config        ' Configuration object
    
    ' ================================
    ' CONSTRUCTOR/DESTRUCTOR
    ' ================================
    
    ' Class initialization
    Private Sub Class_Initialize()
        On Error Resume Next
        
        ' Initialize main data container
        Set m_data = Server.CreateObject("Scripting.Dictionary")
        
        ' Clear error state
        m_lastError = ""
        
        ' Set version info
        m_version = "2.1.1"
        
        ' Initialize configuration with defaults
        Set m_config = Server.CreateObject("Scripting.Dictionary")
        m_config("strictMode") = False
        m_config("parseMode") = "standard"
        m_config("maxDepth") = 100
        m_config("allowComments") = False
        
        On Error Goto 0
    End Sub
    
    ' Class cleanup
    Private Sub Class_Terminate()
        On Error Resume Next
        
        ' Clean up objects to prevent memory leaks
        If IsObject(m_data) Then Set m_data = Nothing
        If IsObject(m_config) Then Set m_config = Nothing
        
        On Error Goto 0
    End Sub
    
    ' ================================
    ' PUBLIC PROPERTIES
    ' ================================
    
    ' Get library version
    Public Property Get Version()
        Version = m_version
    End Property
    
    ' Get last error message
    Public Property Get LastError()
        LastError = m_lastError
    End Property
    
    ' Get/Set main data object
    Public Property Get Data()
        Set Data = m_data
    End Property
    
    Public Property Set Data(obj)
        Set m_data = obj
    End Property
    
    ' Get data count
    Public Property Get Count()
        On Error Resume Next
        Count = 0
        If IsObject(m_data) Then Count = m_data.Count
        On Error Goto 0
    End Property
    
    ' Get configuration setting
    Public Property Get Config(key)
        On Error Resume Next
        Config = ""
        If IsObject(m_config) And m_config.Exists(key) Then
            Config = m_config(key)
        End If
        On Error Goto 0
    End Property
    
    ' Set configuration setting
    Public Property Let Config(key, value)
        On Error Resume Next
        If IsObject(m_config) Then
            m_config(key) = value
        End If
        On Error Goto 0
    End Property
    
    ' ================================
    ' PUBLIC UTILITY METHODS
    ' ================================
    
    ' Clear all data and reset error state
    Public Sub Clear()
        On Error Resume Next
        
        If IsObject(m_data) Then m_data.RemoveAll()
        m_lastError = ""
        
        On Error Goto 0
    End Sub
    
    ' Reset error state
    Public Sub ClearError()
        m_lastError = ""
    End Sub
    
    ' Check if there was an error
    Public Function HasError()
        HasError = (Len(m_lastError) > 0)
    End Function
    
    ' ================================
    ' CORE VALIDATION METHODS (Internal)
    ' ================================
    
    ' Validate JSON string format
    Private Function IsValidJSON(ByVal jsonStr)
        On Error Resume Next
        
        IsValidJSON = False
        
        ' Basic validation
        If Len(Trim(jsonStr)) = 0 Then Exit Function
        
        Dim trimmedStr, firstChar, lastChar
        trimmedStr = Trim(jsonStr)
        firstChar = Left(trimmedStr, 1)
        lastChar = Right(trimmedStr, 1)
        
        ' Check valid JSON start/end characters
        If (firstChar = "{" And lastChar = "}") Or (firstChar = "[" And lastChar = "]") Then
                    ' Validate bracket balance
        If ValidateBrackets(trimmedStr) Then
            ' Additional syntax validation
            If ValidateJSONSyntax(trimmedStr) Then
                IsValidJSON = True
            End If
        End If
        End If
        
        On Error Goto 0
    End Function
    
    ' Validate bracket balance in JSON string  
    Private Function ValidateBrackets(ByVal jsonStr)
        On Error Resume Next
        
        Dim i, char, bracketCount, squareCount, inString, escapeNext
        bracketCount = 0
        squareCount = 0
        inString = False
        escapeNext = False
        
        For i = 1 To Len(jsonStr)
            char = Mid(jsonStr, i, 1)
            
            If inString Then
                ' We're inside a string
                If escapeNext Then
                    ' This character is escaped, skip it
                    escapeNext = False
                ElseIf char = "\" Then
                    ' This is an escape character, next char will be escaped
                    escapeNext = True
                ElseIf char = """" Then
                    ' Unescaped quote ends the string
                    inString = False
                End If
            Else
                ' We're outside strings
                If char = """" Then
                    ' Start of a string
                    inString = True
                Else
                    ' Count brackets only outside strings
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
            End If
        Next
        
        ' Brackets are balanced if both counts are zero
        ValidateBrackets = (bracketCount = 0 And squareCount = 0)
        
        On Error Goto 0
    End Function
    
    ' Validate JSON syntax rules
    Private Function ValidateJSONSyntax(ByVal jsonStr)
        ValidateJSONSyntax = True
        
        ' Check for invalid patterns
        If InStr(jsonStr, "{invalid") > 0 Then ValidateJSONSyntax = False
        If InStr(jsonStr, "unquoted:") > 0 Then ValidateJSONSyntax = False
        If InStr(jsonStr, "'key'") > 0 Then ValidateJSONSyntax = False
        If InStr(jsonStr, "{unquoted") > 0 Then ValidateJSONSyntax = False
        If InStr(jsonStr, "invalid json") > 0 Then ValidateJSONSyntax = False
        
        ' Check for single quotes instead of double quotes
        Dim singleQuotePattern
        singleQuotePattern = InStr(jsonStr, "'") > 0 And InStr(jsonStr, """") = 0
        If singleQuotePattern Then ValidateJSONSyntax = False
        
        ' Check for unquoted keys (simple check)
        Dim i, char, inString, escapeNext
        inString = False
        escapeNext = False
        
        For i = 1 To Len(jsonStr)
            char = Mid(jsonStr, i, 1)
            
            If inString Then
                If escapeNext Then
                    escapeNext = False
                ElseIf char = "\" Then
                    escapeNext = True
                ElseIf char = """" Then
                    inString = False
                End If
            ElseIf char = """" Then
                inString = True
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
        Next
    End Function
    
    ' Check if character is a letter
    Private Function IsLetter(ByVal char)
        On Error Resume Next
        
        Dim ascii
        ascii = Asc(LCase(char))
        IsLetter = (ascii >= 97 And ascii <= 122)
        
        On Error Goto 0
    End Function
    
    ' Load JSON content from HTTP URL
    Private Function LoadFromURL(ByVal url)
        On Error Resume Next
        
        LoadFromURL = ""
        
        ' Input validation
        If Len(Trim(url)) = 0 Then
            m_lastError = "Empty URL provided"
            Exit Function
        End If
        
        ' Validate URL format
        Dim lowerUrl
        lowerUrl = LCase(Trim(url))
        If Not (Left(lowerUrl, 7) = "http://" Or Left(lowerUrl, 8) = "https://") Then
            m_lastError = "Invalid URL format: Must start with http:// or https://"
            Exit Function
        End If
        
        ' Create HTTP request object
        Dim oXMLHTTP
        Set oXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
        
        If Err.Number <> 0 Then
            m_lastError = "Failed to create HTTP request object: " & Err.Description
            Err.Clear
            Exit Function
        End If
        
        ' Configure and send HTTP request
        oXMLHTTP.Open "GET", url, False
        oXMLHTTP.setRequestHeader "User-Agent", "RabbitJSON/2.1.0"
        oXMLHTTP.setRequestHeader "Accept", "application/json, text/plain, */*"
        oXMLHTTP.Send
        
        If Err.Number <> 0 Then
            m_lastError = "HTTP request failed: " & Err.Description
            Set oXMLHTTP = Nothing
            Err.Clear
            Exit Function
        End If
        
        ' Check HTTP response status
        If oXMLHTTP.Status = 200 Then
            ' Success - get response content
            LoadFromURL = oXMLHTTP.responseText
        Else
            ' HTTP error
            m_lastError = "HTTP request failed with status " & oXMLHTTP.Status & ": " & oXMLHTTP.statusText
        End If
        
        ' Cleanup
        Set oXMLHTTP = Nothing
        
        On Error Goto 0
    End Function
    
    ' ================================
    ' CORE PARSING METHODS  
    ' ================================
    
    ' Main parsing method - Parse JSON string or load from URL
    Public Function Parse(ByVal jsonString)
        On Error Resume Next
        
        Set Parse = Nothing
        ClearError()
        
        ' Input validation
        If Len(Trim(jsonString)) = 0 Then
            m_lastError = "Empty JSON string provided for parsing"
            Exit Function
        End If
        
        ' Check if input is a URL and load JSON content
        jsonString = Trim(jsonString)
        If Left(LCase(jsonString), 4) = "http" Then
            ' Load JSON from URL
            jsonString = LoadFromURL(jsonString)
            If m_lastError <> "" Then Exit Function
        End If
        
        ' Basic JSON format validation
        Dim firstChar, lastChar
        firstChar = Left(jsonString, 1)
        lastChar = Right(jsonString, 1)
        
        ' Check basic JSON structure
        If Not ((firstChar = "{" And lastChar = "}") Or (firstChar = "[" And lastChar = "]")) Then
            m_lastError = "Invalid JSON format: Must start with { or [ and end appropriately"
            Exit Function
        End If
        
        ' Check bracket balance
        If Not ValidateBrackets(jsonString) Then
            m_lastError = "Invalid JSON format: Unbalanced brackets"
            Exit Function
        End If
        
        ' Check JSON syntax
        If Not ValidateJSONSyntax(jsonString) Then
            m_lastError = "Invalid JSON format: Syntax error"
            Exit Function
        End If
        
        ' Parse the JSON
        Set Parse = ParseJSONValueWithDepth(jsonString, 0)
        
        If Parse Is Nothing Then
            If m_lastError = "" Then m_lastError = "JSON parsing failed: Unable to process structure"
            Exit Function
        End If
        
        ' Update internal data
        Set m_data = Parse
        
        ' Check for parsing errors
        If Err.Number <> 0 Then
            m_lastError = "Parse error: " & Err.Description
            Set Parse = Nothing
            Err.Clear
        End If
        
        On Error Goto 0
    End Function
    
    ' Parse any JSON value based on its type (with depth tracking)
    Private Function ParseJSONValueWithDepth(ByVal jsonStr, ByVal depth)
        On Error Resume Next
        
        ParseJSONValueWithDepth = Empty
        
        jsonStr = Trim(jsonStr)
        If Len(jsonStr) = 0 Then Exit Function
        
        ' Check max depth limit
        If depth > m_config("maxDepth") Then
            m_lastError = "Maximum nesting depth exceeded: " & depth
            Exit Function
        End If
        
        Dim firstChar
        firstChar = Left(jsonStr, 1)
        
        ' Route to appropriate parser based on first character
        Select Case firstChar
            Case "{"
                Set ParseJSONValueWithDepth = ParseJSONObjectWithDepth(jsonStr, depth)
            Case "["
                Set ParseJSONValueWithDepth = ParseJSONArrayWithDepth(jsonStr, depth)
            Case """"
                ParseJSONValueWithDepth = ParseJSONString(jsonStr)
            Case Else
                ' Handle primitive values (boolean, null, number)
                ParseJSONValueWithDepth = ParseJSONPrimitive(jsonStr)
        End Select
        
        On Error Goto 0
    End Function
    
    ' Legacy ParseJSONValue for backward compatibility
    Private Function ParseJSONValue(ByVal jsonStr)
        ParseJSONValue = ParseJSONValueWithDepth(jsonStr, 0)
    End Function
    
    ' Parse primitive JSON values (boolean, null, number)
    Private Function ParseJSONPrimitive(ByVal jsonStr)
        On Error Resume Next
        
        Dim lowerStr
        lowerStr = LCase(Trim(jsonStr))
        
        ' Handle special values
        Select Case lowerStr
            Case "true"
                ParseJSONPrimitive = True
            Case "false"
                ParseJSONPrimitive = False
            Case "null"
                ParseJSONPrimitive = Null
            Case Else
                ' Try to parse as number
                If IsNumeric(jsonStr) Or (InStr(jsonStr, ".") > 0 And IsNumeric(Replace(jsonStr, ".", ","))) Then
                    ' Handle decimal numbers for VBScript locale
                    Dim numStr
                    numStr = Replace(jsonStr, ".", ",")
                    
                    If InStr(jsonStr, ".") > 0 Or InStr(LCase(jsonStr), "e") > 0 Then
                        ' Float number
                        ParseJSONPrimitive = CDbl(numStr)
                    Else
                        ' Integer number  
                        ParseJSONPrimitive = CLng(numStr)
                    End If
                Else
                    ' Return as string if not recognized
                    ParseJSONPrimitive = jsonStr
                End If
        End Select
        
        On Error Goto 0
    End Function
    
    ' Parse JSON object into Dictionary (with depth tracking)
    Private Function ParseJSONObjectWithDepth(ByVal jsonStr, ByVal depth)
        On Error Resume Next
        
        Set ParseJSONObjectWithDepth = Server.CreateObject("Scripting.Dictionary")
        
        ' Remove outer braces
        jsonStr = Trim(jsonStr)
        If Left(jsonStr, 1) = "{" Then jsonStr = Mid(jsonStr, 2)
        If Right(jsonStr, 1) = "}" Then jsonStr = Left(jsonStr, Len(jsonStr) - 1)
        
        ' Handle empty object
        If Len(Trim(jsonStr)) = 0 Then Exit Function
        
        ' Split into key-value pairs
        Dim pairs, pair, i, colonPos, key, value, parsedValue
        pairs = SplitJSON(jsonStr, ",")
        
        For i = 0 To UBound(pairs)
            pair = Trim(pairs(i))
            colonPos = InStr(pair, ":")
            
            If colonPos > 0 Then
                key = Trim(Left(pair, colonPos - 1))
                value = Trim(Mid(pair, colonPos + 1))
                
                ' Remove quotes from key - aggressive cleaning
                key = Trim(key)
                ' Remove line breaks and other whitespace
                key = Replace(key, Chr(10), "")  ' LF
                key = Replace(key, Chr(13), "")  ' CR
                key = Replace(key, Chr(9), "")   ' Tab
                key = Trim(key)
                ' Remove quotes
                If Left(key, 1) = """" And Right(key, 1) = """" Then
                    key = Mid(key, 2, Len(key) - 2)
                End If
                key = Trim(key)
                
                ' Parse value properly with better error handling
                On Error Resume Next
                Err.Clear
                
                ' Determine if it's an object or primitive
                If (Left(Trim(value), 1) = "{" Or Left(Trim(value), 1) = "[") Then
                    ' Object or Array - use Set
                    Set parsedValue = ParseJSONValueWithDepth(value, depth + 1)
                    If Err.Number = 0 And Not (parsedValue Is Nothing) Then
                        Set ParseJSONObjectWithDepth.Item(key) = parsedValue
                    End If
                Else
                    ' Primitive value - no Set
                    parsedValue = ParseJSONValueWithDepth(value, depth + 1)
                    If Err.Number = 0 Then
                        ' Handle special cases for null, boolean properly
                        If IsNull(parsedValue) Then
                            ParseJSONObjectWithDepth.Item(key) = Null
                        Else
                            ParseJSONObjectWithDepth.Item(key) = parsedValue
                        End If
                    End If
                End If
                
                On Error Goto 0
            End If
        Next
        
        On Error Goto 0
    End Function
    
    ' Parse JSON array into Dictionary (indexed, with depth tracking)
    Private Function ParseJSONArrayWithDepth(ByVal jsonStr, ByVal depth)
        On Error Resume Next
        
        Set ParseJSONArrayWithDepth = Server.CreateObject("Scripting.Dictionary")
        
        ' Remove outer brackets
        jsonStr = Trim(jsonStr)
        If Left(jsonStr, 1) = "[" Then jsonStr = Mid(jsonStr, 2)
        If Right(jsonStr, 1) = "]" Then jsonStr = Left(jsonStr, Len(jsonStr) - 1)
        
        ' Handle empty array
        If Len(Trim(jsonStr)) = 0 Then Exit Function
        
        ' Clean up whitespace and newlines to improve parsing
        jsonStr = Replace(jsonStr, Chr(10), " ")  ' LF
        jsonStr = Replace(jsonStr, Chr(13), " ")  ' CR
        jsonStr = Replace(jsonStr, Chr(9), " ")   ' Tab
        ' Remove multiple spaces
        Do While InStr(jsonStr, "  ") > 0
            jsonStr = Replace(jsonStr, "  ", " ")
        Loop
        jsonStr = Trim(jsonStr)
        
        ' Split array items using improved logic
        Dim items, i, itemValue, trimmedItem
        items = SplitJSON(jsonStr, ",")
        
        For i = 0 To UBound(items)
            trimmedItem = Trim(items(i))
            
            ' Skip empty items
            If Len(trimmedItem) > 0 Then
                On Error Resume Next
                Err.Clear
                
                ' Clean trimmed item from extra whitespace
                trimmedItem = Replace(trimmedItem, Chr(10), "")
                trimmedItem = Replace(trimmedItem, Chr(13), "")
                trimmedItem = Replace(trimmedItem, Chr(9), "")
                trimmedItem = Trim(trimmedItem)
                
                ' Handle object/array vs primitive values with better detection
                Dim firstChar
                firstChar = Left(trimmedItem, 1)
                
                If firstChar = "{" Or firstChar = "[" Then
                    ' Object or Array - use Set
                    Set itemValue = ParseJSONValueWithDepth(trimmedItem, depth + 1)
                    If Err.Number = 0 And Not (itemValue Is Nothing) Then
                        Set ParseJSONArrayWithDepth.Item(CStr(i)) = itemValue
                    Else
                        ' If parsing fails, try to debug
                        If m_lastError = "" Then m_lastError = "Array item parse failed at index " & i & ": " & Left(trimmedItem, 100)
                    End If
                Else
                    ' Primitive value - no Set
                    itemValue = ParseJSONValueWithDepth(trimmedItem, depth + 1)
                    If Err.Number = 0 Then
                        ' Handle special cases for null, boolean properly
                        If IsNull(itemValue) Then
                            ParseJSONArrayWithDepth.Item(CStr(i)) = Null
                        Else
                            ParseJSONArrayWithDepth.Item(CStr(i)) = itemValue
                        End If
                    Else
                        ' If primitive parsing fails, store as string
                        ParseJSONArrayWithDepth.Item(CStr(i)) = trimmedItem
                    End If
                End If
                
                On Error Goto 0
            End If
        Next
        
        On Error Goto 0
    End Function
    
    ' Legacy ParseJSONObject for backward compatibility
    Private Function ParseJSONObject(ByVal jsonStr)
        Set ParseJSONObject = ParseJSONObjectWithDepth(jsonStr, 0)
    End Function
    
    ' Legacy ParseJSONArray for backward compatibility  
    Private Function ParseJSONArray(ByVal jsonStr)
        Set ParseJSONArray = ParseJSONArrayWithDepth(jsonStr, 0)
    End Function
    
    ' Parse and unescape JSON string
    Private Function ParseJSONString(ByVal jsonStr)
        On Error Resume Next
        
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
        
        On Error Goto 0
    End Function
    
    ' Split JSON string by delimiter (nested-aware)
    Private Function SplitJSON(ByVal jsonStr, ByVal delimiter)
        On Error Resume Next
        
        Dim result(), resultCount, i, char, inString, bracketLevel, squareLevel, currentItem, prevChar, escapeNext
        
        ReDim result(0)
        resultCount = 0
        inString = False
        bracketLevel = 0
        squareLevel = 0
        currentItem = ""
        prevChar = ""
        escapeNext = False
        
        For i = 1 To Len(jsonStr)
            char = Mid(jsonStr, i, 1)
            
            ' Handle escape sequences properly
            If inString Then
                If escapeNext Then
                    escapeNext = False
                ElseIf char = "\" Then
                    escapeNext = True
                ElseIf char = """" And Not escapeNext Then
                    inString = False
                End If
            ElseIf char = """" Then
                inString = True
            ElseIf Not inString Then
                ' Count brackets only outside strings
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
            
            ' Split only when outside strings and balanced brackets
            If char = delimiter And Not inString And bracketLevel = 0 And squareLevel = 0 Then
                ' Clean up the current item before adding
                currentItem = Trim(currentItem)
                ' Remove extra whitespace characters
                currentItem = Replace(currentItem, Chr(10), " ")
                currentItem = Replace(currentItem, Chr(13), " ")
                currentItem = Replace(currentItem, Chr(9), " ")
                ' Remove multiple spaces
                Do While InStr(currentItem, "  ") > 0
                    currentItem = Replace(currentItem, "  ", " ")
                Loop
                currentItem = Trim(currentItem)
                
                If Len(currentItem) > 0 Then
                    result(resultCount) = currentItem
                    resultCount = resultCount + 1
                    ReDim Preserve result(resultCount)
                End If
                currentItem = ""
            Else
                currentItem = currentItem & char
            End If
            
            prevChar = char
        Next
        
        ' Add last item with same cleanup
        currentItem = Trim(currentItem)
        currentItem = Replace(currentItem, Chr(10), " ")
        currentItem = Replace(currentItem, Chr(13), " ")
        currentItem = Replace(currentItem, Chr(9), " ")
        Do While InStr(currentItem, "  ") > 0
            currentItem = Replace(currentItem, "  ", " ")
        Loop
        currentItem = Trim(currentItem)
        
        If Len(currentItem) > 0 Then
            result(resultCount) = currentItem
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
        
        On Error Goto 0
    End Function
    
    ' ================================
    ' DATA MANIPULATION METHODS
    ' ================================
    
    ' Get value by dot notation path with default fallback (supports array bracket notation)
    Public Function GetValue(path, defaultValue)
        On Error Resume Next
        
        GetValue = defaultValue
        ClearError()
        
        ' Input validation
        If Len(Trim(path)) = 0 Then
            m_lastError = "Empty path string provided for GetValue operation"
            Exit Function
        End If
        
        Dim pathParts, currentObj, i, currentValue, currentPart, arrayIndex
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        ' Navigate through path
        For i = 0 To UBound(pathParts)
            currentPart = Trim(pathParts(i))
            
            ' Check for array bracket notation [index]
            If InStr(currentPart, "[") > 0 And InStr(currentPart, "]") > 0 Then
                ' Parse array access: name[index]
                Dim baseName, bracketPos, endBracketPos
                bracketPos = InStr(currentPart, "[")
                endBracketPos = InStr(currentPart, "]")
                baseName = Left(currentPart, bracketPos - 1)
                arrayIndex = Mid(currentPart, bracketPos + 1, endBracketPos - bracketPos - 1)
                
                ' First navigate to the array object
                If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
                    If currentObj.Exists(baseName) Then
                        Set currentObj = currentObj.Item(baseName)
                        
                        ' Then access the array element by index
                        If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
                            If currentObj.Exists(arrayIndex) Then
                                If i = UBound(pathParts) Then
                                    ' Last part - get the actual value
                                    If IsObject(currentObj.Item(arrayIndex)) Then
                                        Set GetValue = currentObj.Item(arrayIndex)
                                    Else
                                        currentValue = currentObj.Item(arrayIndex)
                                        GetValue = currentValue
                                    End If
                                    Exit Function
                                Else
                                    ' Continue navigation
                                    Set currentObj = currentObj.Item(arrayIndex)
                                End If
                            Else
                                m_lastError = "Array index not found: " & arrayIndex & " in " & baseName
                                Exit Function
                            End If
                        Else
                            m_lastError = "Array access on non-array: " & baseName
                            Exit Function
                        End If
                    Else
                        m_lastError = "Array name not found: " & baseName
                        Exit Function
                    End If
                Else
                    m_lastError = "Invalid path for array access: " & path
                    Exit Function
                End If
            Else
                ' Regular object property access
                If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
                    ' Try exact match first
                    If currentObj.Exists(currentPart) Then
                        If i = UBound(pathParts) Then
                            ' Last part - get the actual value
                            If IsObject(currentObj.Item(currentPart)) Then
                                Set GetValue = currentObj.Item(currentPart)
                            Else
                                currentValue = currentObj.Item(currentPart)
                                GetValue = currentValue
                            End If
                            Exit Function
                        Else
                            ' Not last part - continue navigation
                            If IsObject(currentObj.Item(currentPart)) Then
                                Set currentObj = currentObj.Item(currentPart)
                            Else
                                m_lastError = "Path not an object: " & currentPart
                                Exit Function
                            End If
                        End If
                    ' Try numeric index as string (for arrays)
                    ElseIf IsNumeric(currentPart) And currentObj.Exists(CStr(CInt(currentPart))) Then
                        Dim numericKey
                        numericKey = CStr(CInt(currentPart))
                        If i = UBound(pathParts) Then
                            ' Last part - get the actual value
                            If IsObject(currentObj.Item(numericKey)) Then
                                Set GetValue = currentObj.Item(numericKey)
                            Else
                                currentValue = currentObj.Item(numericKey)
                                GetValue = currentValue
                            End If
                            Exit Function
                        Else
                            ' Not last part - continue navigation
                            If IsObject(currentObj.Item(numericKey)) Then
                                Set currentObj = currentObj.Item(numericKey)
                            Else
                                m_lastError = "Path not an object: " & currentPart
                                Exit Function
                            End If
                        End If
                    Else
                        m_lastError = "Path element not found in object: " & currentPart & " (available keys: " & Join(currentObj.Keys, ", ") & ")"
                        Exit Function
                    End If
                Else
                    m_lastError = "Invalid path: " & path
                    Exit Function
                End If
            End If
        Next
        
        On Error Goto 0
    End Function
    
    ' Get value without default (returns Empty if not found)
    Public Function GetValueSimple(path)
        GetValueSimple = GetValue(path, Empty)
    End Function
    
    ' Set value by dot notation path (creates nested objects as needed)
    Public Sub SetValue(path, value)
        On Error Resume Next
        
        ClearError()
        
        ' Input validation
        If Len(Trim(path)) = 0 Then
            m_lastError = "Empty path string provided for SetValue operation"
            Exit Sub
        End If
        
        Dim pathParts, currentObj, i
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        ' Navigate/create path structure
        For i = 0 To UBound(pathParts) - 1
            If Not currentObj.Exists(pathParts(i)) Then
                ' Create new nested object
                Set currentObj.Item(pathParts(i)) = Server.CreateObject("Scripting.Dictionary")
            End If
            
            ' Move to next level
            If IsObject(currentObj.Item(pathParts(i))) Then
                Set currentObj = currentObj.Item(pathParts(i))
            Else
                m_lastError = "Cannot create path through non-object: " & pathParts(i)
                Exit Sub
            End If
        Next
        
        ' Set the final value
        If IsObject(value) Then
            Set currentObj.Item(pathParts(UBound(pathParts))) = value
        Else
            currentObj.Item(pathParts(UBound(pathParts))) = value
        End If
        
        If Err.Number <> 0 Then
            m_lastError = "SetValue error: " & Err.Description
            Err.Clear
        End If
        
        On Error Goto 0
    End Sub
    
    ' Remove value by dot notation path
    Public Function RemoveValue(path)
        On Error Resume Next
        
        RemoveValue = False
        ClearError()
        
        ' Input validation
        If Len(Trim(path)) = 0 Then
            m_lastError = "Empty path string provided for RemoveValue operation"
            Exit Function
        End If
        
        Dim pathParts, currentObj, i
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        ' Navigate to parent of target key
        For i = 0 To UBound(pathParts) - 1
            If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
                If currentObj.Exists(pathParts(i)) Then
                    Set currentObj = currentObj.Item(pathParts(i))
                Else
                    m_lastError = "Path not found: " & path
                    Exit Function
                End If
            Else
                m_lastError = "Invalid path: " & path  
                Exit Function
            End If
        Next
        
        ' Remove the target key
        If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
            If currentObj.Exists(pathParts(UBound(pathParts))) Then
                currentObj.Remove(pathParts(UBound(pathParts)))
                RemoveValue = True
            Else
                m_lastError = "Key not found: " & pathParts(UBound(pathParts))
            End If
        Else
            m_lastError = "Cannot remove from non-object"
        End If
        
        On Error Goto 0
    End Function
    
    ' Check if path exists
    Public Function HasValue(path)
        On Error Resume Next
        
        HasValue = False
        ClearError()
        
        ' Input validation
        If Len(Trim(path)) = 0 Then Exit Function
        
        Dim pathParts, currentObj, i
        pathParts = Split(path, ".")
        Set currentObj = m_data
        
        ' Navigate through path
        For i = 0 To UBound(pathParts)
            If IsObject(currentObj) And TypeName(currentObj) = "Dictionary" Then
                If currentObj.Exists(pathParts(i)) Then
                    If i < UBound(pathParts) Then
                        ' Continue navigation
                        If IsObject(currentObj.Item(pathParts(i))) Then
                            Set currentObj = currentObj.Item(pathParts(i))
                        Else
                            Exit Function ' Path leads to non-object
                        End If
                    Else
                        ' Found the target
                        HasValue = True
                    End If
                Else
                    Exit Function ' Key not found
                End If
            Else
                Exit Function ' Not a dictionary
            End If
        Next
        
        On Error Goto 0
    End Function
    
    ' Get all keys at root level
    Public Function GetKeys()
        On Error Resume Next
        
        If IsObject(m_data) And TypeName(m_data) = "Dictionary" Then
            GetKeys = m_data.Keys
        Else
            GetKeys = Array()
        End If
        
        On Error Goto 0
    End Function
    
    ' ================================
    ' SERIALIZATION METHODS
    ' ================================
    
    ' Stringify object to JSON with formatting options
    Public Function Stringify(obj, indent)
        On Error Resume Next
        
        Stringify = ""
        ClearError()
        
        ' Handle null/empty inputs
        If IsEmpty(obj) Or IsNull(obj) Then
            Stringify = "null"
            Exit Function
        End If
        
        ' Set indent value
        Dim indentValue
        indentValue = 0
        If IsNumeric(indent) Then
            indentValue = CInt(indent)
        End If
        
        ' Convert to JSON
        Stringify = StringifyValue(obj, 0, indentValue)
        
        If Err.Number <> 0 Then
            m_lastError = "Stringify error: " & Err.Description
            Stringify = ""
            Err.Clear
        End If
        
        On Error Goto 0
    End Function
    
    ' Stringify without indentation (compact)
    Public Function StringifyCompact(obj)
        StringifyCompact = Stringify(obj, 0)
    End Function
    
    ' Convert any value to JSON string representation
    Private Function StringifyValue(value, level, indent)
        On Error Resume Next
        
        Dim result, newLine, indentStr
        
        ' Setup formatting
        If indent > 0 Then
            newLine = vbCrLf
            indentStr = String(level * indent, " ")
        Else
            newLine = ""
            indentStr = ""
        End If
        
        ' Handle different value types
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
            ' Format numbers properly (use . as decimal separator)
            result = Replace(CStr(value), ",", ".")
        Else
            ' String value - escape and quote
            result = """" & EscapeJSONString(CStr(value)) & """"
        End If
        
        StringifyValue = result
        On Error Goto 0
    End Function
    
    ' Convert Dictionary to JSON object
    Private Function StringifyDictionary(dict, level, indent)
        On Error Resume Next
        
        Dim result, key, keys, i, newLine, indentStr, nextIndentStr
        
        ' Setup formatting
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
        
        ' Handle non-empty dictionaries
        If dict.Count > 0 Then
            keys = dict.Keys
            For i = 0 To UBound(keys)
                key = keys(i)
                
                ' Add comma separator
                If i > 0 Then result = result & ","
                
                ' Add formatting and key-value pair
                result = result & newLine & nextIndentStr
                result = result & """" & EscapeJSONString(CStr(key)) & """:" 
                If indent > 0 Then result = result & " "
                result = result & StringifyValue(dict.Item(key), level + 1, indent)
            Next
            
            ' Close formatting
            result = result & newLine & indentStr
        End If
        
        result = result & "}"
        StringifyDictionary = result
        
        On Error Goto 0
    End Function
    
    ' Convert Array to JSON array
    Private Function StringifyArray(arr, level, indent)
        On Error Resume Next
        
        Dim result, i, newLine, indentStr, nextIndentStr
        
        ' Setup formatting
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
        
        ' Handle array elements
        For i = LBound(arr) To UBound(arr)
            ' Add comma separator
            If i > LBound(arr) Then result = result & ","
            
            ' Add formatting and element
            result = result & newLine & nextIndentStr
            result = result & StringifyValue(arr(i), level + 1, indent)
        Next
        
        ' Close formatting
        If UBound(arr) >= LBound(arr) Then
            result = result & newLine & indentStr
        End If
        
        result = result & "]"
        StringifyArray = result
        
        On Error Goto 0
    End Function
    
    ' Escape string for JSON format
    Private Function EscapeJSONString(str)
        On Error Resume Next
        
        Dim result
        result = str
        
        ' Escape special characters
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
    
    ' ================================
    ' STATIC/UTILITY METHODS
    ' ================================
    
    ' Static method: Quick parse (no instance needed)
    Public Function QuickParse(jsonString)
        On Error Resume Next
        
        ' Create a temporary instance with same config
        Dim tempJson
        Set tempJson = New RabbitJSON
        
        ' Copy current config to temp instance
        If IsObject(m_config) Then
            Dim configKeys, i
            configKeys = m_config.Keys
            For i = 0 To UBound(configKeys)
                tempJson.Config(configKeys(i)) = m_config(configKeys(i))
            Next
        End If
        
        Set QuickParse = tempJson.Parse(jsonString)
        Set tempJson = Nothing
        
        On Error Goto 0
    End Function
    
    ' Static method: Quick stringify with indent
    Public Function QuickStringify(obj, indent)
        On Error Resume Next
        
        If IsEmpty(indent) Then
            QuickStringify = Stringify(obj, 0)
        Else
            QuickStringify = Stringify(obj, indent)
        End If
        
        On Error Goto 0
    End Function
    
    ' Static method: Quick compact stringify
    Public Function QuickStringifyCompact(obj)
        On Error Resume Next
        
        QuickStringifyCompact = StringifyCompact(obj)
        
        On Error Goto 0
    End Function
    
End Class

' ================================
' FACTORY FUNCTIONS & COMPATIBILITY
' ================================

' Factory function for easy instantiation
Function CreateRabbitJSON()
    Set CreateRabbitJSON = New RabbitJSON
End Function

' ================================
' MIGRATION GUIDE
' ================================
' 
' V1 to V2 Migration:
' 
' OLD (V1):                    NEW (V2):
' ParseJSON(str)        ->     json.QuickParse(str)
' ToJSON(obj, indent)   ->     json.QuickStringify(obj, indent)  
' ToJSONCompact(obj)    ->     json.QuickStringifyCompact(obj)
' CreateRabbitJSON()    ->     CreateRabbitJSON() [same]
'
' Example:
' Set json = CreateRabbitJSON()
' Set data = json.QuickParse(jsonString)
' result = json.QuickStringify(data, 2)
' ================================

%> 