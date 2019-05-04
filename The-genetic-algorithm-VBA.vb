The text of the VBA macro program

'Start of the program
'1. Declaring constant values and arrays.
Option Explicit
Option Base 1
Dim InputDataSheet,  OutDataSheet As Worksheet
Dim Population_Size,  NumOfSimulation, CurrentNumOfSimulation, CurrentNumOfPopulation As Long
Dim Objective As String
Dim AorBn, AorB, i, j As Integer
Dim RateOfMutation, RateOfSelection,  RateOfCrossover, Constrain1, Constrain2,  ObjectiveF(),  Constrain1K,  Constrain2K,  ObjectiveK,  RZ,  RY,  RX, MZ,  Population_Input(),Population_Output(), ParamGeom() As Double
Sub ReadInputData()
'2. Reading input data from a worksheet «InputData».
Set InputDataSheet = ThisWorkbook.Sheets("InputData")
'Population size
Population_Size = InputDataSheet.Range("C5")
If Population_Size > 1000 Or Population_Size < 2 Then
    MsgBox "Population size is not included in the specified boundaries. Set the value from 2 to 1000!", vbOKOnly, Error
End If
'Definition of an input and output population array
ReDim Population_Input(Population_Size, 8) As Double
ReDim Population_Output(Population_Size, 8) As Double
ReDim ObjectiveF(Population_Size, 13) As Double
'The value of the mutation coefficient
RateOfMutation = InputDataSheet.Range("C6")
If RateOfMutation > 1 Or RateOfMutation < 0.1 Then
    MsgBox "The value of the mutation coefficient is not included in the specified boundaries. Set the value from 0.0000001 to 1!", vbOKOnly, Error
End If
'The value of the coefficient of selection
RateOfSelection = InputDataSheet.Range("C7")
If RateOfSelection > 1 Or RateOfSelection < 0.01 Then
    MsgBox "The value of the selection coefficient is not included in the specified boundaries. Set the value from 0.01 to 1!", vbOKOnly, Error
End If
'Recombination step value
RateOfCrossover = InputDataSheet.Range("C8")
If Not RateOfCrossover = 2 Or RateOfCrossover = 4 Or RateOfCrossover = 4 Then
    MsgBox "The value of the recombination step is not included in the specified boundaries. Set the value = 2, 4, 8!", vbOKOnly, Error
End If
'Number of generations
NumOfSimulation = Round(InputDataSheet.Range("C9"))
If NumOfSimulation < 1 Or NumOfSimulation > 100 Then
    MsgBox "The number of generations is not included in the specified boundaries. Set the value from 1 to 100!", vbOKOnly, Error
End If
'Limit value # 1
Constrain1 = InputDataSheet.Range("C25")
If Constrain1 < 1 Or Constrain1 > 1000 Then
    MsgBox "Значение ограничения №1 (реакция в опоре по оси X) не входит в заданные границы. Задайте значение от 1 до 1000!", vbOKOnly, Error
End If
'Limit value # 2
Constrain2 = InputDataSheet.Range("C26")
If Constrain2 < 1 Or Constrain2 > 1000 Then
    MsgBox "The value of restriction No. 2 (reaction in the support along the Y axis) does not fall within the specified limits. Set the value from 1 to 1000!", vbOKOnly, Error
End If
'Determination of min or max values of the objective function.
Objective = InputDataSheet.Range("C27")
If Not Objective = "min" Or Objective = "max" Then
    MsgBox "Введите min или max!", vbOKOnly, Error
End If
'The coefficient of significance for the number 1 restrictions
Constrain1K = InputDataSheet.Range("D25")
If Constrain1K < 0 Or Constrain1K > 1 Then
    MsgBox "The coefficient of significance for the number 1 restrictions is not included in the specified boundaries. Set the value from 0 to 1!", vbOKOnly, Error
End If
'Coefficient of significance for the number 2 restrictions
Constrain2K = InputDataSheet.Range("D26")
If Constrain2K < 0 Or Constrain2K > 1 Then
    MsgBox " The coefficient of significance for the number 2 restrictions is not included in the specified boundaries. Set the value from 0 to 1!", vbOKOnly, Error
End If
'Coefficient of significance for the objective function
ObjectiveK = InputDataSheet.Range("D27")
If ObjectiveK < 0 Or ObjectiveK > 1 Then
    MsgBox "The coefficient of significance of the objective function is not included in the specified boundaries. Set the value from 0 to 1!", vbOKOnly, Error
End If
'Determination of the array of parameters of the geometry of the rock-cutting tool.
ReDim ParamGeom(8, Population_Size) As Double
For i = 1 To 8
    For j = 1 To Population_Size
    ParamGeom(i, j) = InputDataSheet.Cells(14 + i, 2 + j)
    Next j
Next i
Set OutDataSheet = ThisWorkbook.Sheets("OutputData")
RateOfSelection = Int(RateOfSelection * Population_Size)
If RateOfSelection < 1 Then RateOfSelection = 1
AorB = (Population_Size - RateOfSelection)
End Sub
'3. Generation of random initial values of tool geometry parameters.
Sub BuildNewRandomPopulation()
Dim i , j As Integer
CurrentNumOfSimulation = 0
CurrentNumOfPopulation = 0
ReadInputData
Randomize
For i = 1 To Population_Size
    For j = 1 To 8
    Population_Input(i, j) = Random_Integer(ParamGeom(i, 1), ParamGeom(i, 2), ParamGeom(i, 3), ParamGeom(i, 4))
  '  Population_Output(i, j) = Population_Input(i, j)
    Next j
Next i
PritOutputDataSheet
RunNewRandomPopulation
RunCurrentNumOfSimulation
End Sub
'4. Calculation of the current population.
Sub RunCurrentNumOfSimulation()
Dim i, i1, j, k As Integer
Dim Numclone,  mutnum As Integer
Set OutDataSheet = ThisWorkbook.Sheets("OutputData")
For i = 1 To NumOfSimulation
'Sort by functional value and record the data of the last population
OutDataSheet.Activate
OutDataSheet.Range(Cells(4 + CurrentNumOfSimulation * Population_Size, 3), _
Cells(4 + CurrentNumOfSimulation * Population_Size + Population_Size - 1, 15)).Sort _
Key1:=OutDataSheet.Range(Cells(4 + CurrentNumOfSimulation * Population_Size, 15), _
Cells(4 + CurrentNumOfSimulation * Population_Size + Population_Size - 1, 15)), Order1:=xlAscending
For j = 1 To Population_Size
    For k = 1 To 13
    ObjectiveF(j, k) = OutDataSheet.Cells(4 + CurrentNumOfSimulation * Population_Size + j - 1, 2 + k)
    Next k
Next j
'Selection, cloning, recombination
For i1 = RateOfSelection + 1 To Population_Size
   For j = 1 To 8
   Randomize
    AorBn = Int(AorB * Rnd) + 1
    If AorBn < 1 Then AorBn = 1
     For k = j To RateOfCrossover + j - 1
      ObjectiveF(i1, k) = (ObjectiveF(AorBn, k))
      Next k
      j = k - 1
   Next j
Next i1
'Modification
For i1 = 1 To Population_Size
   For j = 1 To Int(8 * RateOfMutation)
        Randomize
    mutnum = Int(8 * Rnd) + 1
    If mutnum < 1 Then mutnum = 1
    Randomize
      ObjectiveF(i1, mutnum) = Random_Integer(ParamGeom(i1, 1), ParamGeom(i1, 2), ParamGeom(i1, 3), ParamGeom(i1, 4))
   Next j
Next i1
ReDim Population_Input(Population_Size, 8) As Double
For i1 = 1 To Population_Size
    For j = 1 To 8
    Population_Input(i1, j) = ObjectiveF(i1, j)
    Next j
Next i1
CurrentNumOfSimulation = CurrentNumOfSimulation + 1
RunNewRandomPopulation
Next i
End Sub
'5. Generator of random parameters of tool geometry
Function Random_Integer(UpInteger As Double, DwInteger As Double, NInteger As Double, StInteger As Double) As Double
Dim OutDouble As Double
OutDouble = (UpInteger + DwInteger) / 2 + (StInteger) * (Rnd * NInteger - NInteger / 2)
If OutDouble > UpInteger Then OutDouble = UpInteger
If OutDouble < DwInteger Then OutDouble = DwInteger
Random_Integer = OutDouble
End Function
'6. Calculation of the first population.
Sub RunNewRandomPopulation()
Dim i, j, k As Integer
For i = 1 To Population_Size
j = 1
Dim oFSO As New Scripting.FileSystemObject
        Dim txtread As Scripting.TextStream
        Dim txtwrite As Scripting.TextStream
        Dim ReadLine As String
'Enter the model geometry parameters into the model file.
        Set txtread = oFSO.OpenTextFile("C:\GA\DD-DATAIN.py", ForReading)
        Set txtwrite = oFSO.CreateTextFile("C:\GA\DD-DATAIN1.py", ForWriting, False)
      k = 0
      Do While Not txtread.AtEndOfStream
      ReadLine = Trim(txtread.ReadLine)
      If Not Left(ReadLine, 34) = "s.CircleByCenterPerimeter(center=(" Then GoTo line2
      k = k + 1
      txtwrite.WriteLine ReadLine
        Do While Not txtread.AtEndOfStream
        ReadLine = Trim(txtread.ReadLine)
          If Left(ReadLine, 34) = "s.CircleByCenterPerimeter(center=(" Then
          If k = 1 Then
          txtwrite.WriteLine Left(ReadLine, 34) & Replace((Format(Population_Input(i, j), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") _
          & "), point1=(" & Replace((Format((-0.022 + Population_Input(i, j)), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") & "))"
         j = j + 2
         k = k + 1
         GoTo line1
         End If
           If k = 2 Then
          txtwrite.WriteLine Left(ReadLine, 34) & Replace((Format(Population_Input(i, j), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") _
          & "), point1=(" & Replace((Format((-0.022 + Population_Input(i, j)), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") & "))"
         j = j + 2
         k = k + 1
         GoTo line1
         End If
           If k = 3 Then
          txtwrite.WriteLine Left(ReadLine, 34) & Replace((Format(Population_Input(i, j), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") _
          & "), point1=(" & Replace((Format((-0.028 + Population_Input(i, j)), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") & "))"
         j = j + 2
         k = k + 1
         GoTo line1
         End If
           If k = 4 Then
          txtwrite.WriteLine Left(ReadLine, 34) & Replace((Format(Population_Input(i, j), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") _
          & "), point1=(" & Replace((Format((-0.033 + Population_Input(i, j)), "0.0000")), ",", ".") & "," & Replace((Format(Population_Input(i, j + 1), "0.0000")), ",", ".") & "))"
         j = j + 2
         k = k + 1
         GoTo line1
         End If
         End If
         txtwrite.WriteLine ReadLine
line1:
              Loop
line2:
      txtwrite.WriteLine ReadLine
      Loop
        txtread.Close
        txtwrite.Close
        Set txtread = Nothing
        Set txtwrite = Nothing
SupportFiles
Next i
End Sub
'7. Copying model files to a folder for further calculation.
Sub SupportFiles()
    Dim strFile As Object
    FileCopy "c:\Supportfiles\DATAIN.py", "c:\RUN\DATAIN.py"
    FileCopy "c:\Supportfiles\START.bat", "c:\RUN\START.bat"
    FileCopy "c:\Supportfiles\Job-Base.inp", "c:\RUN\Job-Base.inp"
    FileCopy "c:\Supportfiles\model.cae", "c:\RUN\model.cae"
    FileCopy "c:\Supportfiles\Part-2.exe", "c:\RUN\Part-2.exe"
    FileCopy "c:\Supportfiles\CAE.py", "c:\RUN\CAE.py"
    FileCopy "c:\Supportfiles\Macros_Results.py", "c:\RUN\Macros_Results.py"  
        Set strFile = CreateObject("WScript.shell")
        strFile.Run "cmd /c cd c:\RUN\ & c:\RUN\START.bat", 5, True
      ReadDATAOUTFile
End Sub
'8. Reading calculation results from the model output file.
Sub ReadDATAOUTFile()
Dim i, k As Integer
Dim oFSO As New Scripting.FileSystemObject
        Dim txtread As Scripting.TextStream
        Dim ReadLine As String
        Set txtread = oFSO.OpenTextFile("C:\GA\RUN\DATAOUT.rpt", ForReading)
      For i = 1 To 27
      ReadLine = Trim(txtread.ReadLine)
      Next i
      txtread.Skip (55)
      ReadLine = txtread.Read(13)
      RZ = Replace(ReadLine, ".", ",")
      txtread.Skip (32)
      ReadLine = txtread.Read(13)
      MZ = Replace(ReadLine, ".", ",")
      ReadLine = Trim(txtread.ReadLine)
      ReadLine = Trim(txtread.ReadLine)
      txtread.Skip (116)
      ReadLine = txtread.Read(13)
      RY = Replace(ReadLine, ".", ",")
      ReadLine = txtread.Read(14)
      RX = Replace(ReadLine, ".", ",")
        txtread.Close
        Set txtread = Nothing
CurrentNumOfPopulation = CurrentNumOfPopulation + 1
    Kill "c:\RUN\*.*"
PritOutputDataSheet
End Sub
'9. Print the results of the calculation in the final table on the sheet «OutputData».
Sub PritOutputDataSheet()
Dim i, j, k As Integer
Set OutDataSheet = ThisWorkbook.Sheets("OutputData")
 If CurrentNumOfPopulation = 0 Then
OutDataSheet.Range(OutDataSheet.Cells(4, 1), OutDataSheet.Cells(1000, 19)).Clear
For i = 1 To Population_Size
    OutDataSheet.Cells(3 + i, 1) = CurrentNumOfSimulation
    OutDataSheet.Cells(3 + i, 2) = i
    For j = 1 To 8
    OutDataSheet.Cells(3 + i, 2 + j) = Population_Input(i, j)
    Next j
Next i
GoTo line1
End If
 If CurrentNumOfSimulation > 0 Then
For i = 1 To Population_Size
    OutDataSheet.Cells(3 + CurrentNumOfSimulation * Population_Size + i, 1) = CurrentNumOfSimulation
    OutDataSheet.Cells(3 + CurrentNumOfSimulation * Population_Size + i, 2) = i
    For j = 1 To 8
    OutDataSheet.Cells(3 + CurrentNumOfSimulation * Population_Size + i, 2 + j) = Population_Input(i, j)
    Next j
Next i
End If
Set InputDataSheet = ThisWorkbook.Sheets("InputData")
OutDataSheet.Cells(3 + CurrentNumOfPopulation, 11) = MZ
OutDataSheet.Cells(3 + CurrentNumOfPopulation, 12) = RX
OutDataSheet.Cells(3 + CurrentNumOfPopulation, 13) = RY
OutDataSheet.Cells(3 + CurrentNumOfPopulation, 14) = Abs(RZ)
OutDataSheet.Cells(3 + CurrentNumOfPopulation, 15) = Abs(RZ) * InputDataSheet.Range("D27") + (RX - InputDataSheet.Range("C25")) * InputDataSheet.Range("D25") _
+ (RY - InputDataSheet.Range("C26")) * InputDataSheet.Range("D26")
line1:
End Sub
'End of the program