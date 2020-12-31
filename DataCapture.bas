Attribute VB_Name = "DataCapture"


Sub AddNewData()

'This will capture user input to data
'25-09-2020
'---------------------------------------

Sheets("Data").Activate
Range("A2").Select

Do Until ActiveCell.Value = ""
ActiveCell.Offset(1, 0).Select
Loop



ActiveCell.Value = Sheets("Form").Range("Client").Value
ActiveCell.Offset(0, 1).Value = Sheets("Form").Range("Date").Value
ActiveCell.Offset(0, 2).Value = Sheets("Form").Range("Amount").Value

Lapas1.Activate
Lapas1.Range("C13").Value = "Last Execution Info : Data Submitted Successfully. " & Now()


End Sub

