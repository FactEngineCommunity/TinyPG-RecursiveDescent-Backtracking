Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Namespace TinyPG.Highlighter
    Public Class Scanner
        ' Methods
        Public Sub New()
            Me.SkipList.Add(TokenType.WHITESPACE)
            Dim regex As New Regex("\s+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.WHITESPACE, regex)
            Me.Tokens.Add(TokenType.WHITESPACE)
            regex = New Regex("^$", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.EOF, regex)
            Me.Tokens.Add(TokenType.EOF)
            regex = New Regex("//[^\n]*\n?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.GRAMMARCOMMENTLINE, regex)
            Me.Tokens.Add(TokenType.GRAMMARCOMMENTLINE)
            regex = New Regex("/\*([^*]+|\*[^/])+(\*/)?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.GRAMMARCOMMENTBLOCK, regex)
            Me.Tokens.Add(TokenType.GRAMMARCOMMENTBLOCK)
            regex = New Regex("@?\""(\""\""|[^\""])*(""|\n)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVESTRING, regex)
            Me.Tokens.Add(TokenType.DIRECTIVESTRING)
            regex = New Regex("^(@TinyPG|@Parser|@Scanner|@Grammar|@ParseTree|@TextHighlighter)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVEKEYWORD, regex)
            Me.Tokens.Add(TokenType.DIRECTIVEKEYWORD)
            regex = New Regex("^(@|(%[^>])|=|"")+?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVESYMBOL, regex)
            Me.Tokens.Add(TokenType.DIRECTIVESYMBOL)
            regex = New Regex("[^%@=""]+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVENONKEYWORD, regex)
            Me.Tokens.Add(TokenType.DIRECTIVENONKEYWORD)
            regex = New Regex("<%", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVEOPEN, regex)
            Me.Tokens.Add(TokenType.DIRECTIVEOPEN)
            regex = New Regex("%>", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVECLOSE, regex)
            Me.Tokens.Add(TokenType.DIRECTIVECLOSE)
            regex = New Regex("[^\[\]]", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.ATTRIBUTESYMBOL, regex)
            Me.Tokens.Add(TokenType.ATTRIBUTESYMBOL)
            regex = New Regex("^(Skip|Color)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.ATTRIBUTEKEYWORD, regex)
            Me.Tokens.Add(TokenType.ATTRIBUTEKEYWORD)
            regex = New Regex("[^\(\)\]\n\s]+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.ATTRIBUTENONKEYWORD, regex)
            Me.Tokens.Add(TokenType.ATTRIBUTENONKEYWORD)
            regex = New Regex("\[\s*", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.ATTRIBUTEOPEN, regex)
            Me.Tokens.Add(TokenType.ATTRIBUTEOPEN)
            regex = New Regex("\s*\]\s*", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.ATTRIBUTECLOSE, regex)
            Me.Tokens.Add(TokenType.ATTRIBUTECLOSE)
            regex = New Regex("^(abstract|as|base|break|case|catch|checked|class|const|continue|decimal|default|delegate|double|do|else|enum|event|explicit|extern|false|finally|fixed|float|foreach|for|get|goto|if|implicit|interface|internal|int|in|is|lock|namespace|new|null|object|operator|out|override|params|partial|private|protected|public|readonly|ref|return|sealed|set|sizeof|stackalloc|static|struct|switch|throw|this|true|try|typeof|unchecked|unsafe|ushort|using|virtual|void|volatile|while)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CS_KEYWORD, regex)
            Me.Tokens.Add(TokenType.CS_KEYWORD)
            regex = New Regex("^(AddHandler|AddressOf|Alias|AndAlso|And|Ansi|Assembly|As|Auto|Boolean|ByRef|Byte|ByVal|Call|Case|Catch|CBool|CByte|CChar|CDate|CDec|CDbl|Char|CInt|Class|CLng|CObj|Const|CShort|CSng|CStr|CType|Date|Decimal|Declare|Default|Delegate|Dim|DirectCast|Double|Do|Each|ElseIf|Else|End|Enum|Erase|Error|Event|Exit|False|Finally|For|Friend|Function|GetType|Get|GoSub|GoTo|Handles|If|Implements|Imports|Inherits|Integer|Interface|In|Is|Let|Lib|Like|Long|Loop|Me|Mod|Module|MustInherit|MustOverride|MyBase|MyClass|Namespace|New|Next|Nothing|NotInheritable|NotOverridable|Not|Object|On|Optional|Option|OrElse|Or|Overloads|Overridable|Overrides|ParamArray|Preserve|Private|Property|Protected|Public|RaiseEvent|ReadOnly|ReDim|REM|RemoveHandler|Resume|Return|Select|Set|Shadows|Shared|Short|Single|Static|Step|Stop|String|Structure|Sub|SyncLock|Then|Throw|To|True|Try|TypeOf|Unicode|Until|Variant|When|While|With|WithEvents|WriteOnly|Xor|Source)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.VB_KEYWORD, regex)
            Me.Tokens.Add(TokenType.VB_KEYWORD)
            regex = New Regex("^(abstract|as|base|break|case|catch|checked|class|const|continue|decimal|default|delegate|double|do|else|enum|event|explicit|extern|false|finally|fixed|float|foreach|for|get|goto|if|implicit|interface|internal|int|in|is|lock|namespace|new|null|object|operator|out|override|params|partial|private|protected|public|readonly|ref|return|sealed|set|sizeof|stackalloc|static|struct|switch|throw|this|true|try|typeof|unchecked|unsafe|ushort|using|virtual|void|volatile|while)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOTNET_KEYWORD, regex)
            Me.Tokens.Add(TokenType.DOTNET_KEYWORD)
            regex = New Regex("^(Array|AttributeTargets|AttributeUsageAttribute|Attribute|BitConverter|Boolean|Buffer|Byte|Char|CharEnumerator|CLSCompliantAttribute|ConsoleColor|ConsoleKey|ConsoleKeyInfo|ConsoleModifiers|ConsoleSpecialKey|Console|ContextBoundObject|ContextStaticAttribute|Converter|Convert|DateTimeKind|DateTimeOffset|DateTime|DayOfWeek|DBNull|Decimal|Delegate|Double|Enum|Environment.SpecialFolder|EnvironmentVariableTarget|Environment|EventArgs|EventHandler|Exception|FlagsAttribute|GCCollectionMode|GC|Guid|ICloneable|IComparable|IConvertible|ICustomFormatter|IDisposable|IEquatable|IFormatProvider|IFormattable|IndexOutOfRangeException|InsufficientMemoryException|Int16|Int32|Int64|IntPtr|InvalidCastException|InvalidOperationException|InvalidProgramException|MarshalByRefObject|Math|MidpointRounding|NotFiniteNumberException|NotImplementedException|NotSupportedException|Nullable|NullReferenceException|ObjectDisposedException|Object|ObsoleteAttribute|OperatingSystem|OutOfMemoryException|OverflowException|ParamArrayAttribute|PlatformID|PlatformNotSupportedException|Predicate|Random|SByte|SerializableAttribute|Single|StackOverflowException|StringComparer|StringComparison|StringSplitOptions|String|SystemException|TimeSpan|TimeZone|TypeCode|TypedReference|TypeInitializationException|Type|UInt16|UInt32|UInt64|UIntPtr|UnauthorizedAccessException|UnhandledExceptionEventArgs|UnhandledExceptionEventHandler|ValueType|Void|WeakReference|Comparer|Dictionary|EqualityComparer|ICollection|IComparer|IDictionary|IEnumerable|IEnumerator|IEqualityComparer|IList|KeyNotFoundException|KeyValuePair|List|ASCIIEncoding|Decoder|DecoderExceptionFallback|DecoderExceptionFallbackBuffer|DecoderFallback|DecoderFallbackBuffer|DecoderFallbackException|DecoderReplacementFallback|DecoderReplacementFallbackBuffer|EncoderExceptionFallback|EncoderExceptionFallbackBuffer|EncoderFallback|EncoderFallbackBuffer|EncoderFallbackException|EncoderReplacementFallback|EncoderReplacementFallbackBuffer|Encoder|EncodingInfo|Encoding|NormalizationForm|StringBuilder|UnicodeEncoding|UTF32Encoding|UTF7Encoding|UTF8Encoding)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOTNET_TYPES, regex)
            Me.Tokens.Add(TokenType.DOTNET_TYPES)
            regex = New Regex("//[^\n]*\n?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CS_COMMENTLINE, regex)
            Me.Tokens.Add(TokenType.CS_COMMENTLINE)
            regex = New Regex("/\*([^*]+|\*[^/])+(\*/)?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CS_COMMENTBLOCK, regex)
            Me.Tokens.Add(TokenType.CS_COMMENTBLOCK)
            regex = New Regex("[^}]", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CS_SYMBOL, regex)
            Me.Tokens.Add(TokenType.CS_SYMBOL)
            regex = New Regex("([^""\n\s/;.}\(\)\[\]]|/[^/*]|}[^;])+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CS_NONKEYWORD, regex)
            Me.Tokens.Add(TokenType.CS_NONKEYWORD)
            regex = New Regex("@?[""]([""][""]|[^\""\n])*[""]?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CS_STRING, regex)
            Me.Tokens.Add(TokenType.CS_STRING)
            regex = New Regex("'[^\n]*\n?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.VB_COMMENTLINE, regex)
            Me.Tokens.Add(TokenType.VB_COMMENTLINE)
            regex = New Regex("REM[^\n]*\n?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.VB_COMMENTBLOCK, regex)
            Me.Tokens.Add(TokenType.VB_COMMENTBLOCK)
            regex = New Regex("[^}]", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.VB_SYMBOL, regex)
            Me.Tokens.Add(TokenType.VB_SYMBOL)
            regex = New Regex("([^""\n\s/;.}\(\)\[\]]|/[^/*]|}[^;])+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.VB_NONKEYWORD, regex)
            Me.Tokens.Add(TokenType.VB_NONKEYWORD)
            regex = New Regex("@?[""]([""][""]|[^\""\n])*[""]?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.VB_STRING, regex)
            Me.Tokens.Add(TokenType.VB_STRING)
            regex = New Regex("//[^\n]*\n?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOTNET_COMMENTLINE, regex)
            Me.Tokens.Add(TokenType.DOTNET_COMMENTLINE)
            regex = New Regex("/\*([^*]+|\*[^/])+(\*/)?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOTNET_COMMENTBLOCK, regex)
            Me.Tokens.Add(TokenType.DOTNET_COMMENTBLOCK)
            regex = New Regex("[^}]", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOTNET_SYMBOL, regex)
            Me.Tokens.Add(TokenType.DOTNET_SYMBOL)
            regex = New Regex("([^""\n\s/;.}\[\]\(\)]|/[^/*]|}[^;])+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOTNET_NONKEYWORD, regex)
            Me.Tokens.Add(TokenType.DOTNET_NONKEYWORD)
            regex = New Regex("@?[""]([""][""]|[^\""\n])*[""]?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOTNET_STRING, regex)
            Me.Tokens.Add(TokenType.DOTNET_STRING)
            regex = New Regex("\{", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CODEBLOCKOPEN, regex)
            Me.Tokens.Add(TokenType.CODEBLOCKOPEN)
            regex = New Regex("\};", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CODEBLOCKCLOSE, regex)
            Me.Tokens.Add(TokenType.CODEBLOCKCLOSE)
            regex = New Regex("(Start)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.GRAMMARKEYWORD, regex)
            Me.Tokens.Add(TokenType.GRAMMARKEYWORD)
            regex = New Regex("->", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.GRAMMARARROW, regex)
            Me.Tokens.Add(TokenType.GRAMMARARROW)
            regex = New Regex("[^{}\[\]/<>]|[</]$", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.GRAMMARSYMBOL, regex)
            Me.Tokens.Add(TokenType.GRAMMARSYMBOL)
            regex = New Regex("([^;""\[\n\s/<{\(\)]|/[^/*]|<[^%])+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.GRAMMARNONKEYWORD, regex)
            Me.Tokens.Add(TokenType.GRAMMARNONKEYWORD)
            regex = New Regex("@?[""]([""][""]|[^\""\n])*[""]?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.GRAMMARSTRING, regex)
            Me.Tokens.Add(TokenType.GRAMMARSTRING)
        End Sub

        Public Function GetToken(ByVal type As TokenType) As Token
            Return New Token(Me.StartPos, Me.EndPos) With { _
                .Type = type _
            }
        End Function

        Public Sub Init(ByVal input As String)
            Me.Input = input
            Me.StartPos = 0
            Me.EndPos = 0
            Me.CurrentLine = 0
            Me.CurrentColumn = 0
            Me.CurrentPosition = 0
            Me.LookAheadToken = Nothing
        End Sub

        Public Function LookAhead(ByVal ParamArray expectedtokens As TokenType()) As Token
            Dim scantokens As List(Of TokenType)
            Dim startpos As Integer = Me.StartPos
            Dim tok As Token = Nothing
            If (((Not Me.LookAheadToken Is Nothing) AndAlso (Me.LookAheadToken.Type <> TokenType._UNDETERMINED_)) AndAlso (Me.LookAheadToken.Type <> TokenType._NONE_)) Then
                Return Me.LookAheadToken
            End If
            If (expectedtokens.Length = 0) Then
                scantokens = Me.Tokens
            Else
                scantokens = New List(Of TokenType)(expectedtokens)
                scantokens.AddRange(Me.SkipList)
            End If
            Do
                Dim len As Integer = -1
                Dim index As TokenType = DirectCast(&H7FFFFFFF, TokenType)
                Dim input As String = Me.Input.Substring(startpos)
                tok = New Token(startpos, Me.EndPos)
                Dim i As Integer
                For i = 0 To scantokens.Count - 1
                    Dim m As Match = Me.Patterns.Item(scantokens.Item(i)).Match(input)
                    If ((m.Success AndAlso (m.Index = 0)) AndAlso ((m.Length > len) OrElse ((DirectCast(scantokens.Item(i), TokenType) < index) AndAlso (m.Length = len)))) Then
                        len = m.Length
                        index = scantokens.Item(i)
                    End If
                Next i
                If ((index >= TokenType._NONE_) AndAlso (len >= 0)) Then
                    tok.EndPos = (startpos + len)
                    tok.Text = Me.Input.Substring(tok.StartPos, len)
                    tok.Type = index
                ElseIf (tok.StartPos < (tok.EndPos - 1)) Then
                    tok.Text = Me.Input.Substring(tok.StartPos, 1)
                End If
                If Me.SkipList.Contains(tok.Type) Then
                    startpos = tok.EndPos
                    Me.Skipped.Add(tok)
                Else
                    tok.Skipped = Me.Skipped
                    Me.Skipped = New List(Of Token)
                End If
            Loop While Me.SkipList.Contains(tok.Type)
            Me.LookAheadToken = tok
            Return tok
        End Function

        Public Function Scan(ByVal ParamArray expectedtokens As TokenType()) As Token
            Dim tok As Token = Me.LookAhead(expectedtokens)
            Me.LookAheadToken = Nothing
            Me.StartPos = tok.EndPos
            Me.EndPos = tok.EndPos
            Return tok
        End Function


        ' Fields
        Public CurrentColumn As Integer
        Public CurrentLine As Integer
        Public CurrentPosition As Integer
        Public EndPos As Integer = 0
        Public Input As String
        Private LookAheadToken As Token = Nothing
        Public Patterns As Dictionary(Of TokenType, Regex) = New Dictionary(Of TokenType, Regex)
        Private SkipList As List(Of TokenType) = New List(Of TokenType)
        Public Skipped As List(Of Token) = New List(Of Token)
        Public StartPos As Integer = 0
        Private Tokens As List(Of TokenType) = New List(Of TokenType)
    End Class
End Namespace

