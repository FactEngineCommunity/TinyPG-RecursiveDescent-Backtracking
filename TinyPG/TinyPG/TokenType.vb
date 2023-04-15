Imports System

Namespace TinyPG
    Public Enum TokenType
        ' Fields
        _NONE_ = 0
        _UNDETERMINED_ = 1

        Start = 2
        CommentBlock = 3
        DirectiveBlock = 4
        GrammarBlock = 5
        AttributeBlock = 6
        CodeBlock = 7

        ARROW = &H1C
        ASSIGN = 20
        Attribute = 6
        BRACKETCLOSE = 15
        BRACKETOPEN = 14
        COMMA = &H11
        COMMENTLINE = &H22
        ConcatRule = 12
        Directive = 3
        DIRECTIVECLOSE = 30
        DIRECTIVEOPEN = &H1D
        [DOUBLE] = &H1A
        EOF = &H1F
        ExtProduction = 5
        HEX = &H1B
        IDENTIFIER = &H18
        [INTEGER] = &H19
        NameValue = 4
        Param = 8
        Params = 7
        PIPE = &H15
        Production = 9
        Rule = 10
        SEMICOLON = &H16
        SQUARECLOSE = &H13
        SQUAREOPEN = &H12
        [STRING] = &H20
        Subrule = 11
        Symbol = 13
        UNARYOPER = &H17
        WHITESPACE = &H21
        CS_COMMENTLINE = 27
        CS_COMMENTBLOCK = 28
        CS_SYMBOL = 29
        CS_NONKEYWORD = 30
        CS_STRING = 31
        VB_COMMENTLINE = 32
        VB_COMMENTBLOCK = 33
        VB_SYMBOL = 34
        VB_NONKEYWORD = 35
        VB_STRING = 36
        DOTNET_COMMENTLINE = 37
        DOTNET_COMMENTBLOCK = 38
        DOTNET_SYMBOL = 39
        DOTNET_NONKEYWORD = 40
        DOTNET_STRING = 41
        DIRECTIVESTRING = 12
        DIRECTIVEKEYWORD = 13
        DIRECTIVESYMBOL = 14
        DIRECTIVENONKEYWORD = 15
        CS_KEYWORD = 23
        VB_KEYWORD = 24
        DOTNET_KEYWORD = 25

    End Enum
End Namespace

