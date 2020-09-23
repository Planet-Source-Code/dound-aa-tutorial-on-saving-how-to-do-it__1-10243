<div align="center">

## Aa Tutorial on Saving: How to do it\!


</div>

### Description

This code will teach a user how to save their data and also how to retrieve it. Very easy to understand if you read the ' comments. :) Enjoy!
 
### More Info
 
Just read the code/comments and be prepared to learn!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dound](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dound.md)
**Level**          |Intermediate
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dound-aa-tutorial-on-saving-how-to-do-it__1-10243/archive/master.zip)





### Source Code

```
Dim strTest As String
Private Sub Command1_Click()
FileNum = FreeFile() 'Finds a freefile where it can write to
Open App.Path & "\Test.test" For Input As FileNum 'opens the file to (input = get data)
 Input #FileNum, strTest 'Get data by putting Input then the FileNumber you opened in (we used a variable FileNum) then a comma then the variable you want to store.
Close FileNum 'Close the FileNumber you opened...'Close' by itself will close ALL of your open files.
Text1.Text = strTest 'sets the textbox's text = to what was is the file
End Sub
Private Sub Command2_Click()
FileNum = FreeFile()
Open App.Path & "\Test.test" For Output As FileNum 'Output clears the file and gives you access to write to it with the Write command
 Write #FileNum, strTest 'Write then #FileNumber (we used a variable) then a comma then the variable you want to save
Close FileNum 'You can save multiple variables at once if you seperate them by commas
End Sub
Private Sub Command3_Click()
Kill App.Path & "\Test.test"
'Deletes File "Test.test" at the application's path
End Sub
Private Sub Text1_Change()
strTest = Text1.Text 'puts text in the string whenever the textbox's text changes
End Sub
'End of code...
'The easiest way to write multiple values to a file is like this:
'Write #FileNum, strTest, strTest2, strTest3
'...just keep adding commas and then the next vaule to save
'Be sure to load EVERYTHING you saved and in the same order you saved it! or it wont work!!
'A shortcut for saving arrays:
'For Q = 1 To 5: Write #NFile, Array(Q):Next
'...and so on...Enjoy!
```

