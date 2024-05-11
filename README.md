# TinyPG-with Recursive Descent and Backtracking

A fork of the TinyPG parser, modified to perform recursive descent and backtracking. https://www.codeproject.com/Articles/28294/A-Tiny-Parser-Generator-v1-2

Default Tiny PG is a LL(1) Parser Generator.

We needed a parser with recursive descent and backtracking at FactEngine. We had already done quite a bit of work with TinyPG and are very taken by the technology. It's simple and effective. So we modified 

This fork modifies TinyPG to provide ordered recursive descent and backtracking.

**Suitability**

Small parsing jobs over a single document at a time.

**How it works**
The Parser Generator produces classes that generate a Parse Tree. 

A copy of the parse tree is kept.

1. A MaxLength is kept that keeps track of how far the parser progressed before finding an error or a solution. MaxLength = number of characters successfully parsed;
2. If no solution is found, the parser backtracks, by default of cascading method calls (in reverse);
3. The  parser will follow any further Options/Optionals and their parse tree is developed;
4. If MaxLength is surpassed...the parse tree is cloned to the copy of the parse tree.
5. If a successfull path/solution is found, then the copy of the parse tree is returned.

In this way we turn TinyPG into a Parser Generator with Recursive Descent and Backtracking.

**Efficiency**

This strategy (cloning the parse tree) was suitable for our purposes. There may be other more efficient ways of amending the parse tree returned. Contributions are welcome.

**Contributions**

Please get in touch if you'd like to contribute. Look us up at FactEngine.

**ToDos**

Only VB.Net is catered for. C# needs to be done if anyone needs it. We merely modified the templates for VB.Net in TinyPG to achieve our result, so the same changes can be made to the C# generator templates.
