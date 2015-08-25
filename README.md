Foxpro .mem File Parser for VBA
==============================

If you need to parse .mem files in VBA applications, use this
God, Foxpro is ancient, but sometimes, you need to support legacy software.
UBS Sdn Bhd I'm looking at you.

This class only supports three data types: Character sequences, numeric (double-precision floating point) and dates.
Apparently it is compatible with VFP9 (according to the link below).
If you need to parse other data types and/or .MEM files from other versions of Foxpro, you have to extend the parser.
Follow the link below to the code in another language.

Example usage (compatible with MS Excel, MS Access VBA):

    Dim varsMemFile as FoxproMemFile
    
    Set varsMemFile = New FoxproMemFile
    varsMemFile.Init("data.mem")
    
    Dim someNum As Double
    someNum = varsMemFile.data("NUMERIC_VARIABLE")
    
    Dim someDate As Date
    someDate = varsMemFile.data("DATE_VARIABLE")

Couldn't have written this without the information [here](http://www.tek-tips.com/viewthread.cfm?qid=1687712).

