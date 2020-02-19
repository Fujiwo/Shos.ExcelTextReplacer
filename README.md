# Shos.ExcelTextReplacer
Excel Text Replacer

## Usage

Shos.ExcelTextReplacer [targetExcelFilePath] [replacementListExcelFilePath]

ex.

- targetExcelFile (before):

                apple,iphone,3
                apple,IPHONE,2
                Yahoo,PIXEL,1

- replacementListExcelFile:

                old text,new text
                apple,Apple
                IPHONE,iPhone
                iphone,iPhone
                Yahoo,Google

- targetExcelFile (after):

                Apple,iPhone,3
                Apple,iPhone,2
                Google,PIXEL,1
