Imports System

Namespace TinyPG.Highlighter
    Public Enum TokenType
        ' Fields
        _NONE_ = 0
        _UNDETERMINED_ = 1
        AttributeBlock = 6
        ATTRIBUTECLOSE = &H16
        ATTRIBUTEKEYWORD = &H13
        ATTRIBUTENONKEYWORD = 20
        ATTRIBUTEOPEN = &H15
        ATTRIBUTESYMBOL = &H12
        CodeBlock = 7
        CODEBLOCKCLOSE = &H2B
        CODEBLOCKOPEN = &H2A
        CommentBlock = 3
        CS_COMMENTBLOCK = &H1C
        CS_COMMENTLINE = &H1B
        CS_KEYWORD = &H17
        CS_NONKEYWORD = 30
        CS_STRING = &H1F
        CS_SYMBOL = &H1D
        DirectiveBlock = 4
        DIRECTIVECLOSE = &H11
        DIRECTIVEKEYWORD = 13
        DIRECTIVENONKEYWORD = 15
        DIRECTIVEOPEN = &H10
        DIRECTIVESTRING = 12
        DIRECTIVESYMBOL = 14
        DOTNET_COMMENTBLOCK = &H26
        DOTNET_COMMENTLINE = &H25
        DOTNET_KEYWORD = &H19
        DOTNET_NONKEYWORD = 40
        DOTNET_STRING = &H29
        DOTNET_SYMBOL = &H27
        DOTNET_TYPES = &H1A
        EOF = 9
        GRAMMARARROW = &H2D
        GrammarBlock = 5
        GRAMMARCOMMENTBLOCK = 11
        GRAMMARCOMMENTLINE = 10
        GRAMMARKEYWORD = &H2C
        GRAMMARNONKEYWORD = &H2F
        GRAMMARSTRING = &H30
        GRAMMARSYMBOL = &H2E
        Start = 2
        VB_COMMENTBLOCK = &H21
        VB_COMMENTLINE = &H20
        VB_KEYWORD = &H18
        VB_NONKEYWORD = &H23
        VB_STRING = &H24
        VB_SYMBOL = &H22
        WHITESPACE = 8
    End Enum
End Namespace

