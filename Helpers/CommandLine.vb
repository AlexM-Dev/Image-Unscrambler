Public Class CommandLine
#Region "Fields"

    Private _ParametersString As String = Nothing
    Private _CaseSensitive As Boolean = False
    Private _Parsed As Boolean = False
    Private _Success As Boolean = False
    Private _Offset As Integer = -1
    Private _CurrentName As String = Nothing
    Private _CurrentValue As String = Nothing
    Private _Parameters As Dictionary(Of String, String) = Nothing

#End Region


#Region "Properties"

    ''' <summary>
    ''' Gets or sets the parameters string.
    ''' </summary>
    ''' <value>The parameters string.</value>
    Public Property ParametersString() As String
        Get
            Return _ParametersString
        End Get
        Set
            _ParametersString = Value

            Parse()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value indicating whether parameters name are case sensitive.
    ''' </summary>
    ''' <value><c>true</c> if parameters name are case sensitive; otherwise, <c>false</c>.</value>
    Public ReadOnly Property CaseSensitive() As Boolean
        Get
            Return _CaseSensitive
        End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating whether this <see cref="CommandLine"/> has been parsed.
    ''' </summary>
    ''' <value><c>true</c> if parsed; otherwise, <c>false</c>.</value>
    Public ReadOnly Property Parsed() As Boolean
        Get
            Return _Parsed
        End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating whether this <see cref="CommandLine"/> has been parsed successfully.
    ''' </summary>
    ''' <value><c>true</c> if parsed successfully; otherwise, <c>false</c>.</value>
    Public ReadOnly Property Success() As Boolean
        Get
            Return _Success
        End Get
    End Property

    ''' <summary>
    ''' Gets the offset where the parsing blocked.
    ''' </summary>
    ''' <value>The error offset.</value>
    Public ReadOnly Property ErrorOffset() As Integer
        Get
            Return _Offset
        End Get
    End Property

    ''' <summary>
    ''' Gets the parameters dictionary.
    ''' </summary>
    ''' <value>The parameters dictionary.</value>
    Public ReadOnly Property Parameters() As Dictionary(Of String, String)
        Get
            If Not _Parsed Then
                Parse()
            End If

            Return _Parameters
        End Get
    End Property

    ''' <summary>
    ''' Gets the count of parameters currently parsed.
    ''' </summary>
    ''' <value>The count.</value>
    Public ReadOnly Property Count() As Integer
        Get
            Return Parameters.Count
        End Get
    End Property

    ''' <summary>
    ''' Gets the <see cref="System.String"/> parameter value with the specified parameter name.
    ''' </summary>
    ''' <value></value>
    Default Public ReadOnly Property Item(parameterName As String) As String
        Get
            If Not _CaseSensitive Then
                parameterName = parameterName.ToLower()
            End If

            Dim parms As Dictionary(Of String, String) = Parameters

            If parms.ContainsKey(parameterName) Then
                Return parms(parameterName)
            End If

            Return Nothing
        End Get
    End Property

#End Region


#Region "Construction"

    Public Sub New()
        Me.New(False)
    End Sub

    Public Sub New(caseSensitive As Boolean)
        Me.New(String.Empty, caseSensitive)
    End Sub

    Public Sub New(commandLine__1 As String)
        Me.New(commandLine__1, False)
    End Sub

    Public Sub New(parameters As String())
        Me.New(parameters, False)
    End Sub

    Public Sub New(parameters As String(), caseSensitive As Boolean)
        Me.New(String.Join(" ", parameters), caseSensitive)
    End Sub

    Public Sub New(commandLine__1 As String, caseSensitive As Boolean)
        _CaseSensitive = caseSensitive
        _Parameters = New Dictionary(Of String, String)()
        ParametersString = commandLine__1
    End Sub

#End Region


#Region "Methods"

    ''' <summary>
    ''' Clears the parameters dictionary.
    ''' </summary>
    Public Sub Clear()
        _Parameters.Clear()
    End Sub

    ''' <summary>
    ''' Parses the ParametersString and fill the parameters dictionary.
    ''' </summary>
    ''' <returns>true if parsing was succesful</returns>
    ''' <remarks>
    ''' 
    ''' Grammar
    ''' 
    ''' parameters : parameter*
    '''            ;
    ''' 
    ''' parameter : '/' parameter-struct
    '''           | '--' parameter-struct
    '''           | '-' parameter-struct
    '''           ;
    ''' 
    ''' parameter-struct : parameter-name (parameter-separator parameter-value)? ;
    ''' 
    ''' parameter-name : parameter-name-char+ ;
    ''' 
    ''' parameter-separator : ':'
    '''                     | '='
    '''                     | ' '
    '''                     ;
    ''' 
    ''' parameter-value : '\'' (ANY LESS '\'')* '\''
    '''                 | '"' (ANY LESS '"')* '"'
    '''                 | (parameter-value-first-char parameter-value-char*)?
    '''                 ;
    ''' 
    ''' parameter-name-char : ANY LESS ' ', ':', '=' ;
    ''' 
    ''' parameter-value-first-char : ANY LESS ' ', '/', ':', '=' ;
    ''' 
    ''' parameter-value-char : ANY LESS ' ' ;
    ''' 
    ''' 
    ''' Matches:
    ''' 
    ''' Application.exe /new /parm: "A parameter can be '/parm: '/value:''"   /parm2:  value2   /parm3: "value3: 'the value'"
    ''' 
    ''' </remarks>
    Public Function Parse() As Boolean
        Try
            Clear()

            _Offset = 0
            _Parsed = True
            _Success = False

            If Not String.IsNullOrEmpty(_ParametersString) Then
                If MatchParameters(_Offset) Then
                    _Success = True
                End If
            Else
                _Success = True
            End If
        Catch generatedExceptionName As ArgumentOutOfRangeException
            _Success = False
        Catch generatedExceptionName As IndexOutOfRangeException
            _Success = False
        End Try

        Return _Success
    End Function

    Private Function MatchParameters(pos As Integer) As Boolean
        While pos < _ParametersString.Length AndAlso MatchParameter(pos)
            pos = _Offset

            While pos < _ParametersString.Length AndAlso MatchSpace(pos)
                pos += 1
            End While
        End While

        If pos = _ParametersString.Length Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchParameter(pos As Integer) As Boolean
        If MatchSlash(pos) Then
            pos += 1
        ElseIf MatchMinus(pos) Then
            pos += 1

            If MatchMinus(pos) Then
                pos += 1
            End If
        Else
            Return False
        End If

        If MatchParameterStruct(pos) Then
            If Not _CaseSensitive Then
                _CurrentName = _CurrentName.ToLower()
            End If

            _Parameters(_CurrentName) = _CurrentValue

            pos = _Offset

            ' Error? (pos < (_ParametersString.Length-1))
            If pos < _ParametersString.Length AndAlso Not MatchSpace(pos) Then
                Return False
            End If

            Return True
        End If

        Return False
    End Function

    Private Function MatchParameterStruct(pos As Integer) As Boolean
        If MatchParameterName(pos) Then
            _CurrentName = _ParametersString.Substring(pos, _Offset - pos)
            _CurrentValue = "true"

            pos = _Offset

            If pos < _ParametersString.Length Then
                If MatchParameterSeparator(pos) Then
                    While MatchSpace(pos + 1)
                        pos += 1
                    End While

                    If MatchParameterValue(pos + 1) Then
                        pos += 1

                        _CurrentValue = _ParametersString.Substring(pos, _Offset - pos)

                        If String.IsNullOrEmpty(_CurrentValue) Then
                            _CurrentValue = "true"
                        ElseIf _CurrentValue(0) = "'"c OrElse _CurrentValue(0) = """"c Then
                            _CurrentValue = _CurrentValue.Substring(1, _CurrentValue.Length - 2)
                        End If

                        pos = _Offset
                    End If
                End If
            End If

            _Offset = pos

            Return True
        End If

        Return False
    End Function

    Private Function MatchParameterName(pos As Integer) As Boolean
        Dim pos2 As Integer = pos

        While pos < _ParametersString.Length AndAlso MatchParameterNameChar(pos)
            pos += 1
        End While

        If pos > pos2 Then
            _Offset = pos

            Return True
        End If

        Return False
    End Function

    Private Function MatchParameterSeparator(pos As Integer) As Boolean
        If Me._ParametersString(pos) = ":"c OrElse Me._ParametersString(pos) = "="c OrElse Me._ParametersString(pos) = " "c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchParameterValue(pos As Integer) As Boolean
        If MatchQuote(pos) Then
            pos += 1

            While pos < _ParametersString.Length AndAlso Not MatchQuote(pos)
                pos += 1
            End While

            If MatchQuote(pos) Then
                pos += 1

                _Offset = pos

                Return True
            End If
        ElseIf MatchDoubleQuote(pos) Then
            pos += 1

            While pos < _ParametersString.Length AndAlso Not MatchDoubleQuote(pos)
                pos += 1
            End While

            If MatchDoubleQuote(pos) Then
                pos += 1

                _Offset = pos

                Return True
            End If
        ElseIf MatchParameterValueFirstChar(pos) Then
            pos += 1

            While pos < _ParametersString.Length AndAlso MatchParameterValueChar(pos)
                pos += 1
            End While

            _Offset = pos

            Return True
        End If


        Return False
    End Function

    Private Function MatchParameterNameChar(pos As Integer) As Boolean
        If _ParametersString(pos) <> " "c AndAlso _ParametersString(pos) <> ":"c AndAlso _ParametersString(pos) <> "="c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchParameterValueChar(pos As Integer) As Boolean
        If _ParametersString(pos) <> " "c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchParameterValueFirstChar(pos As Integer) As Boolean
        If _ParametersString(pos) <> " "c AndAlso _ParametersString(pos) <> "/"c AndAlso _ParametersString(pos) <> ":"c AndAlso _ParametersString(pos) <> "="c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchSlash(pos As Integer) As Boolean
        If Me._ParametersString(pos) = "/"c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchMinus(pos As Integer) As Boolean
        If Me._ParametersString(pos) = "-"c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchColon(pos As Integer) As Boolean
        If Me._ParametersString(pos) = ":"c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchSpace(pos As Integer) As Boolean
        If Me._ParametersString(pos) = " "c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchQuote(pos As Integer) As Boolean
        If Me._ParametersString(pos) = "'"c Then
            Return True
        End If

        Return False
    End Function

    Private Function MatchDoubleQuote(pos As Integer) As Boolean
        If Me._ParametersString(pos) = """"c Then
            Return True
        End If

        Return False
    End Function

#End Region
End Class