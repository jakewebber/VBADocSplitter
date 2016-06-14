# VBADocSplitter
*VBA-Word script that splits a Word Doc into HTML files by provided delimiters. Performs autoformatting for a specific report template.*


## Sub Procedures & Functions
####FormatDocSplitter()
Main procedure for the script. Runs all other functions required for formatting the document properly both for splitting and HTML conversion. Edit these procedure variables below to fit your document and remove any function calls that perform unneeded autoformatting.
``` VBA
startDelimiter = "start"  'Start delimiter for document splitting
endDelimiter = "end"      'End delimiter for document splitting
maxFileName = 180         'Sections with larger names will be truncated to this size and have " ..." appended.
```
The delimiters are found and replaced with whitespace characters (combination of tabs/newlines) in order to effectively remove the delimiters from the document while still being searchable for splitting.


-----
#####Example template for this script:
![Alt text](http://i.imgur.com/3EdFXog.png "Optional title")

-----
For This Report, the `startDelimiter` would be **Return to Table of Contents** 
and the `endDelimiter` Would be the string regex **(\-{6,})**, which finds any number of 6 or more hyphens. 
Each section would be saved as the section header, such as 

**1.1  Seven Year-Old Unable To Maintain Single Cohesive Storyline While Playing With Action Figures.htm**


The main procedure includes error checking for sections that delimiter searches may have missed and skipped in the split. If the number found in the header does not increment by .1 or .01 relatively, an error message will appear with the option to cancel the script and prints to the Debug Log for tracking.  
For instance, splitting section *1.2*  then *1.4* would produce an error with this information for skipping *1.3*. Same for *1.21* to *1.23*. However, there would be no error splitting *1.34* and then *2.1*, since the parenting integer was incremented instead of the decimal. 

#####*Skeleton of the Beef*
```VBA
Set Section = ActiveDocument.Range.Duplicate
With Section.Find                                   'Find sections that split the document.
    While .Execute
        Set Header = Section.Duplicate
        With Header.Find                            'Find Header within Sections
            If .Execute Then
                Set HeaderNum = Header.Duplicate
                With HeaderNum.Find
                    If .Execute Then                'Find number label in header for checking split errors
                       
                       'Code for counting incrementing headers and error printing
                       
                    End If
                End With
                Call CopyAndSave(Section, Header, maxFileName) '<- Function for saving split section
            Else
              'No section header was found at all from the given Header.Find
            End If
        End With
    Wend
End With
```


####CopyAndSave(Section As Range, Header As Range, maxFileName As Integer)
Called in `FormatDocSplitter()` for saving each section from the original document and and additional formatting for each section. 

`Section` range covers from current `startDelimiter` to `endDelimiter`. 

`Header` range covers from the header number (i.e. 1.1) to the next newline (paragraph) character. Header is also removed from the document's content. 

`maxFileName` sets the limit for the size of the file save name (truncated if larger) 

This procedure also cleans up the file name from header: a large number of characters that are not valid in a file name (such as \ / : * ? " < > |), as well as characters that are not valid in a WordPress post name (needs to be UTF8), are removed or otherwise reformatted. Leading/trailing whitespace is removed as well. 

Files are saved as simple HTM, or `wdFormatFilteredHTML`

####RemoveAllHyperlinks()
Strips all hyperlinks from a document, preserving original linked text. This simplifies delimiter searching if the delimiters contain hyperlinks. 

####URLtoHyperlink()
Restores hyperlinks that are plain URLs. Also helps with formatting reports where many URLs were not initially hyperlinked. 

####StripAccent(aString As String)
Strips all accent characters from the given `aString` and replaces them with plaintext alphabet chars. 

####DeleteShapes()
Removes Word shapes from the document. Word's bad HTML conversion will preserve shapes in documents as linked images in HTML. Use this if your document uses line shapes as visual separators. 
