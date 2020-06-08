'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Environment
Imports System.IO
Imports System.Linq

'This module contains this program's core procedures.
Public Module ViewerModule
   'This enumeration lists the types of highlights used.
   Private Enum HighlightTypesE As Integer
      Normal            'Indicates data.
      Marked            'Indicates data that has been marked.
      AllSame           'Indicates that the data at a specific position is the same for the entire fileset.
   End Enum

   'This enumeration lists the types of sorthing methods used.
   Private Enum SortingMethodsE As Integer
      None             'Does not sort the files.
      ByData           'Sorts the files by their data.
      ByExtension      'Sorts the files by their extension.
      ByName           'Sorts the files by their name.
      BySize           'Sorts the files by their size.
   End Enum

   'This structure defines a file set's location.
   Private Structure FileSetLocationStr
      Public Path As String      'Defines the path of the set's files.
      Public Pattern As String   'Defines the pattern of the set's files.
   End Structure

   'This structure defines a file.
   Private Structure FileStr
      Public Data() As Byte                       'Defines the file's unmodified data.
      Public DisplayedData() As Byte              'Defines the file's data as displayed.
      Public FileName As String                   'Defines the file's name.
      Public HighlightTypes() As HighlightTypesE  'Defines the file data's highlight types
      Public Size As Integer                      'Defines the file's length.
   End Structure

   'This structure defines the statistics for the fileset.
   Private Structure StatisticsStr
      Public Longest As Integer                   'Defines the size of the largest file.
      Public Highest() As Byte                    'Defines the highest data values.
      Public Lowest() As Byte                     'Defines the lowest data values.
      Public Shortest As Integer                  'Defines the size of the smallest file.  
   End Structure

   Private ReadOnly COLORS() As ConsoleColor = {ConsoleColor.Gray, ConsoleColor.Green, ConsoleColor.DarkGray}   'The highlighting colors used.
   Private ReadOnly NOT_DISPLAYED() As Byte = {&H7%, &H8%, &H9%, &HA%, &HD%}                                    'The byte values that cannot be displayed as text.

   Private Const COLUMN_COUNT As Integer = &H11%    'The number of data columns displayed at the same time
   Private Const FILE_NAME_LENGTH As Integer = 13   'The maximum number of characters displayed for a file name.
   Private Const HALF_ROW_COUNT As Integer = &HA%   'The number of data rows displayed at the same time when the split view is used.
   Private Const ROW_COUNT As Integer = &H16%       'The number of data rows displayed at the same time.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim InvertBits As Boolean = False
         Dim Offset As Integer = &H0%
         Dim Shifter As Integer = &H0%
         Dim ShowStatistics As Boolean = False
         Dim ShowText As Boolean = False
         Dim TopRow As Integer = &H0%

         Initialize()

         Console.Clear()
         Do
            DisplayData(Offset, TopRow, Shifter, InvertBits, ShowStatistics, ShowText)
            Do Until Console.KeyAvailable
            Loop
            Select Case Console.ReadKey(intercept:=True).Key
               Case ConsoleKey.Add
                  If Shifter < &HFF% Then Shifter += &H1%
                  FileSet(, , , , , Shifter:=Shifter)
               Case ConsoleKey.DownArrow
                  If TopRow < FileSet().Count Then TopRow += &H1%
               Case ConsoleKey.End
                  Offset = If(My.Computer.Keyboard.ShiftKeyDown, Statistics().Shortest, Statistics().Longest) - COLUMN_COUNT
               Case ConsoleKey.Escape
                  If Choose("Quit y/n? ", "NYny", Left:=2, Top:=24).ToString().ToUpper() = "Y" Then Exit Do
               Case ConsoleKey.F1
                  DisplayHelp()
               Case ConsoleKey.F2
                  FileSet(, , SortingMethodsE.ByData, Offset)
               Case ConsoleKey.F3
                  FileSet(, , SortingMethodsE.ByName)
               Case ConsoleKey.F4
                  FileSet(, , SortingMethodsE.ByExtension)
               Case ConsoleKey.F5
                  FileSet(, , SortingMethodsE.BySize)
               Case ConsoleKey.F6
                  ShowStatistics = Not ShowStatistics
                  ShowText = False
               Case ConsoleKey.F7
                  ShowText = Not ShowText
                  ShowStatistics = False
               Case ConsoleKey.F8
                  InvertBits = Not InvertBits
                  FileSet(, , , , InvertBits:=True)
               Case ConsoleKey.F9
                  FileSet(, , , , , , ToggleBackwards:=True)
                  SetHighlights()
               Case ConsoleKey.F10
                  Markers(Refresh:=True)
                  SetHighlights()
               Case ConsoleKey.F12
                  ExportData()
               Case ConsoleKey.Home
                  Offset = &H0%
               Case ConsoleKey.LeftArrow
                  If Offset > UInt32.MinValue Then Offset -= &H1%
               Case ConsoleKey.PageDown
                  If Offset + COLUMN_COUNT <= UInt32.MaxValue Then Offset += COLUMN_COUNT
               Case ConsoleKey.PageUp
                  If Offset - COLUMN_COUNT >= UInt32.MinValue Then Offset -= COLUMN_COUNT Else Offset = &H0%
               Case ConsoleKey.RightArrow
                  If Offset < UInt32.MaxValue Then Offset += &H1%
               Case ConsoleKey.Subtract
                  If Shifter > &H0% Then Shifter -= &H1%
                  FileSet(, , , , , Shifter:=Shifter)
               Case ConsoleKey.UpArrow
                  If TopRow > &H0% Then TopRow -= &H1%
            End Select
         Loop

         Console.Clear()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the fileset's data.
   Private Sub DisplayData(Offset As Integer, TopRow As Integer, Shifter As Integer, InvertBits As Boolean, ShowStatistics As Boolean, ShowText As Boolean)
      Try
         Console.Clear()
         Console.SetCursorPosition(left:=0, top:=0)
         Console.BackgroundColor = ConsoleColor.Gray
         Console.ForegroundColor = ConsoleColor.Black
         Console.Write($"{"File:",-14}{"Size: ",-11}")
         For Position As Integer = Offset To Offset + COLUMN_COUNT
            Console.Write((Position And &HFF%).ToString("X").PadLeft(3))
         Next Position
         Console.Write(" ")

         Console.SetCursorPosition(left:=11, top:=24)
         Console.Write($" Row: {TopRow,10}   Offset: {Offset,10}   Invert: {InvertBits,5}   Shifter: {Shifter:X2} ")
         Console.BackgroundColor = ConsoleColor.Black
         Console.ForegroundColor = ConsoleColor.Gray
         Console.SetCursorPosition(left:=0, top:=1)
         For Row As Integer = TopRow To TopRow + If(ShowText, HALF_ROW_COUNT, ROW_COUNT - If(ShowStatistics, 2, 0))
            If Row >= FileSet().Count Then Exit For
            With FileSet()(Row)
               Console.ForegroundColor = ConsoleColor.Gray
               Console.Write(If(.FileName.Length > FILE_NAME_LENGTH, .FileName.Substring(0, FILE_NAME_LENGTH), .FileName).PadRight(14))
               Console.Write(.Size.ToString("X").PadLeft(9).PadRight(11))
               For Position As Integer = Offset To Offset + COLUMN_COUNT
                  If Position >= .DisplayedData.Length Then Exit For
                  Console.ForegroundColor = COLORS(.HighlightTypes(Position))
                  Console.Write(.DisplayedData(Position).ToString("X").PadLeft(3))
               Next Position
               Console.WriteLine()
            End With
         Next Row

         If ShowStatistics Then
            DisplayStatistics(Offset)
         ElseIf ShowText Then
            DisplayText(Offset, TopRow)
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the data markers.
   Private Sub DisplayMarkers(TopRow As Integer)
      Try
         Console.Clear()
         Console.BackgroundColor = ConsoleColor.Gray
         Console.ForegroundColor = ConsoleColor.Black
         Console.SetCursorPosition(left:=1, top:=23)
         Console.WriteLine($" DEL = remove, ESC = back, ENTER = add, UP/DOWN = scroll   Marker: { If(Markers().Count > 0, TopRow + 1, 0)}/{Markers().Count} ")

         Console.BackgroundColor = ConsoleColor.Black
         Console.ForegroundColor = ConsoleColor.Gray
         Console.SetCursorPosition(left:=2, top:=1)
         Console.WriteLine("[Markers]")
         For Marker As Integer = TopRow To Markers().Count - 1
            Console.SetCursorPosition(left:=2, top:=Console.CursorTop)
            Array.ForEach(Markers()(Marker), Sub(ByteO As Byte) Console.Write($"{ByteO:X2} "))
            Console.WriteLine()
            If Console.CursorTop >= Console.WindowHeight - 3 Then Exit For
         Next Marker
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the fileset's statistics.
   Private Sub DisplayStatistics(Offset As Integer)
      Try
         With Statistics()
            Console.ForegroundColor = ConsoleColor.Gray
            Console.CursorLeft = 17
            Console.Write("Lowest: ")
            For Position As Integer = Offset To Offset + COLUMN_COUNT
               If Position >= .Longest Then Exit For
               Console.Write(.Lowest(Position).ToString("X").PadLeft(3))
            Next Position

            Console.WriteLine()
            Console.CursorLeft = 16
            Console.Write("Highest: ")
            For Position As Integer = Offset To Offset + COLUMN_COUNT
               If Position >= .Longest Then Exit For
               Console.Write(.Highest(Position).ToString("X").PadLeft(3))
            Next Position
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the fileset's data as text.
   Private Sub DisplayText(Offset As Integer, TopRow As Integer)
      Try
         Dim ByteO As New Byte

         Console.CursorTop = HALF_ROW_COUNT + 2
         Console.BackgroundColor = ConsoleColor.Gray
         Console.ForegroundColor = ConsoleColor.Black
         For Position As Integer = Offset To Offset + Console.WindowWidth - 3 Step 3
            Console.Write(((Position + &H2%) And &HFF%).ToString("X").PadLeft(3))
         Next Position
         Console.Write("  ")
         Console.BackgroundColor = ConsoleColor.Black
         Console.ForegroundColor = ConsoleColor.Gray

         For Row As Integer = TopRow To TopRow + HALF_ROW_COUNT
            If Row >= FileSet().Count Then Exit For
            With FileSet()(Row)
               For Position As Integer = Offset To Offset + Console.WindowWidth - 2
                  If Position >= .DisplayedData.Length Then Exit For
                  Console.ForegroundColor = COLORS(.HighlightTypes(Position))
                  ByteO = .DisplayedData(Position)
                  Console.Write(If(Array.IndexOf(NOT_DISPLAYED, ByteO) < 0, ToChar(ByteO), " "))
               Next Position
               Console.WriteLine()
            End With
         Next Row
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure exports the current fileset's data.
   Private Sub ExportData()
      Try
         Dim ExportFile As String = Nothing

         Do
            Console.BackgroundColor = ConsoleColor.Black
            Console.ForegroundColor = ConsoleColor.Gray
            Console.Clear()
            Console.SetCursorPosition(left:=1, top:=1)
            ExportFile = GetInput("Export to: ")
            If ExportFile = Nothing Then Exit Sub
            If Not Directory.Exists(Path.GetDirectoryName(ExportFile)) Then
               HandleError(New DirectoryNotFoundException)
            End If
         Loop Until Directory.Exists(Path.GetDirectoryName(ExportFile))

         Using FileO As New StreamWriter(ExportFile)
            For Index As Integer = 0 To FileSet.Count - 1
               With FileSet()(Index)
                  FileO.Write($"{ .FileName};{ .Size:X};")
                  Array.ForEach(.Data, Sub(ByteO As Byte) FileO.Write("{ByteO:X};"))
                  FileO.WriteLine()
               End With
            Next Index
         End Using
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the fileset's data.
   Private Function FileSet(Optional FileSetPath As String = Nothing, Optional FileNamePattern As String = Nothing, Optional SortingMethod As SortingMethodsE = SortingMethodsE.None, Optional Offset As Integer = Nothing, Optional InvertBits As Boolean = False, Optional Shifter As Integer? = Nothing, Optional ToggleBackwards As Boolean = False) As List(Of FileStr)
      Static CurrentFiles As New List(Of FileStr)

      Try
         Dim FileItem As New FileStr

         If FileSetPath IsNot Nothing Then
            CurrentFiles.Clear()
            For Each FileO As FileInfo In My.Computer.FileSystem.GetDirectoryInfo(FileSetPath).GetFiles(FileNamePattern)
               FileItem = New FileStr With {.Data = File.ReadAllBytes(FileO.FullName), .FileName = FileO.Name.ToLower, .Size = CInt(FileO.Length)}
               With FileItem
                  ReDim .DisplayedData(0 To .Data.GetUpperBound(0))
                  Array.Copy(.Data, .DisplayedData, .Data.Length)
                  ReDim FileItem.HighlightTypes(0 To FileItem.Size)
               End With
               CurrentFiles.Add(FileItem)
            Next FileO

            Statistics(Refresh:=True)
         End If

         Select Case SortingMethod
            Case SortingMethodsE.ByData
               CurrentFiles = (From Item In CurrentFiles Order By If(Offset <= Item.DisplayedData.GetUpperBound(0), Item.DisplayedData(Offset), Nothing)).ToList
            Case SortingMethodsE.ByExtension
               CurrentFiles.Sort(Function(Item1 As FileStr, Item2 As FileStr) Path.GetExtension(Item1.FileName).CompareTo(Path.GetExtension(Item2.FileName)))
            Case SortingMethodsE.ByName
               CurrentFiles.Sort(Function(Item1 As FileStr, Item2 As FileStr) Item1.FileName.CompareTo(Item2.FileName))
            Case SortingMethodsE.BySize
               CurrentFiles.Sort(Function(Item1 As FileStr, Item2 As FileStr) Item1.Size.CompareTo(Item2.Size))
         End Select

         If InvertBits Then
            For Index As Integer = 0 To CurrentFiles.Count - 1
               With CurrentFiles(Index)
                  For Position As Integer = .DisplayedData.GetLowerBound(0) To .DisplayedData.GetUpperBound(0)
                     .DisplayedData(Position) = CByte(.DisplayedData(Position) Xor &HFF%)
                  Next Position
               End With
            Next Index
         End If

         If Shifter IsNot Nothing Then
            For Index As Integer = 0 To CurrentFiles.Count - 1
               With CurrentFiles(Index)
                  For Position As Integer = .DisplayedData.GetLowerBound(0) To .DisplayedData.GetUpperBound(0)
                     If .Data(Position) + Shifter.Value > &HFF% Then
                        .DisplayedData(Position) = CByte((.Data(Position) + Shifter.Value) - &HFF%)
                     Else
                        .DisplayedData(Position) = CByte(.Data(Position) + Shifter.Value)
                     End If
                  Next Position
               End With
            Next Index
         End If

         If ToggleBackwards Then
            For Index As Integer = 0 To CurrentFiles.Count - 1
               Array.Reverse(CurrentFiles(Index).DisplayedData)
            Next Index
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return CurrentFiles
   End Function

   'This procedure returns the fileset path and file name pattern specified by the user.
   Private Function GetPathPattern() As String
      Dim PathPattern As String = Nothing

      Try
         With My.Application.CommandLineArgs
            If .Count = 0 Then
               Console.Clear()
               Console.WriteLine(ProgramInformation())
               Console.WriteLine()
               PathPattern = GetInput("Specify ""[Path]*.Extension"": ")
            Else
               PathPattern = .First
            End If
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return PathPattern
   End Function

   'This procedure initializes this program.
   Private Sub Initialize()
      Dim FileSetLocation As New FileSetLocationStr
      Dim PathPattern As String = Nothing

      Try
         Console.ForegroundColor = ConsoleColor.Gray
         Console.BackgroundColor = ConsoleColor.Black
         Console.CursorVisible = False
         Console.Title = My.Application.Info.Title
         Console.WindowWidth = 80
         Console.WindowHeight = 25
         Console.SetBufferSize(Console.WindowWidth, Console.WindowHeight)

         With FileSetLocation
            Do
               PathPattern = GetPathPattern()
               If PathPattern = Nothing Then
                  [Exit](0)
               Else
                  If PathPattern.Contains("*") Then
                     .Path = PathPattern.Substring(0, PathPattern.IndexOf("*"))
                     .Pattern = PathPattern.Substring(PathPattern.IndexOf("*"))
                  Else
                     .Path = PathPattern
                     .Pattern = "*.*"
                  End If
               End If

               If Directory.Exists(.Path) Then
                  Exit Do
               Else
                  HandleError(New DirectoryNotFoundException)
               End If
            Loop

            FileSet(.Path, .Pattern)
         End With

         SetHighlights()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the data markers.
   Private Function Markers(Optional Refresh As Boolean = False) As List(Of Byte())
      Static CurrentMarkers As New List(Of Byte())

      Try
         Dim IsDuplicate As New Boolean
         Dim NewMarker As String = Nothing
         Dim NewMarkerBytes As New List(Of Byte)
         Dim TopRow As Integer = 0

         If Refresh Then
            Console.ForegroundColor = ConsoleColor.Gray
            Do
               DisplayMarkers(TopRow)
               Do Until Console.KeyAvailable
               Loop
               Select Case Console.ReadKey(intercept:=True).Key
                  Case ConsoleKey.Delete
                     If Markers().Count() > 0 Then Markers().RemoveAt(TopRow)
                  Case ConsoleKey.DownArrow
                     If TopRow < Markers().Count - 1 Then TopRow += 1
                  Case ConsoleKey.Enter
                     Console.SetCursorPosition(left:=2, top:=1)
                     NewMarker = GetInput("New marker: ", Filter:="0123456789abcdefABCDEF ")
                     If Not NewMarker = Nothing Then
                        NewMarkerBytes.Clear()
                        For Each ByteO As String In NewMarker.Split(" "c)
                           If Not ByteO.Trim = Nothing Then
                              If ByteO.Trim.Length > 2 Then
                                 NewMarkerBytes.Clear()
                                 Exit For
                              Else
                                 NewMarkerBytes.Add(ToByte(ByteO.Trim, fromBase:=16))
                              End If
                           End If
                        Next ByteO
                        If NewMarkerBytes.Count > 0 Then
                           IsDuplicate = False
                           For Each Marker As Byte() In Markers()
                              If Marker.Length = NewMarkerBytes.Count AndAlso FindBytes(0, NewMarkerBytes.ToArray, Marker) >= 0 Then
                                 Console.SetCursorPosition(left:=2, top:=1)
                                 Console.WriteLine("Marker already has been added.")
                                 Do Until Console.ReadKey(intercept:=True).Key = ConsoleKey.Enter : Loop
                                 IsDuplicate = True
                                 Exit For
                              End If
                           Next Marker
                           If Not IsDuplicate Then
                              CurrentMarkers.Add(NewMarkerBytes.ToArray)
                              If CurrentMarkers.Count > Console.WindowHeight - 5 AndAlso TopRow < CurrentMarkers.Count - 1 Then TopRow = CurrentMarkers.Count - 1
                           End If
                        Else
                           Console.SetCursorPosition(left:=2, top:=1)
                           Console.WriteLine("Please specify space delimited hexadecimal byte values.")
                           Do Until Console.ReadKey(intercept:=True).Key = ConsoleKey.Enter : Loop
                        End If
                     End If
                  Case ConsoleKey.Escape
                     Exit Do
                  Case ConsoleKey.UpArrow
                     If TopRow > 0 Then TopRow -= 1
               End Select
            Loop
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return CurrentMarkers
   End Function

   'This procedure sets the highlights for the data.
   Private Sub SetHighlights()
      Try
         Dim AllSame As New Boolean
         Dim Offset As New Integer?
         Dim PreviousValue As New Integer

         If FileSet().Count > 0 Then
            For Position As Integer = 0 To Statistics().Longest
               AllSame = True
               If Position < FileSet().First.Size Then PreviousValue = FileSet().First.DisplayedData(Position)
               For Index As Integer = 0 To FileSet().Count - 1
                  With FileSet()(Index)
                     If Position < .Size Then
                        If Not .DisplayedData(Position) = PreviousValue Then
                           AllSame = False
                           PreviousValue = .DisplayedData(Position)
                        End If
                     End If
                  End With
                  If Not AllSame Then Exit For
               Next Index

               For Index As Integer = 0 To FileSet().Count - 1
                  With FileSet()(Index)
                     If Position < .Size Then .HighlightTypes(Position) = If(AllSame, HighlightTypesE.AllSame, HighlightTypesE.Normal)
                  End With
               Next Index
            Next Position

            For Index As Integer = 0 To FileSet().Count - 1
               With FileSet()(Index)
                  For Each Marker As Byte() In Markers()
                     Offset = FindBytes(0, .DisplayedData, Marker)
                     Do Until Offset Is Nothing
                        If Offset IsNot Nothing Then
                           For Position As Integer = Offset.Value To Offset.Value + Marker.Length - 1
                              .HighlightTypes(Position) = HighlightTypesE.Marked
                           Next Position
                        End If
                        Offset = FindBytes(Offset.Value + Marker.Length, .DisplayedData, Marker)
                     Loop
                  Next Marker
               End With
            Next Index
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the statistics for the files.
   Private Function Statistics(Optional Refresh As Boolean = False) As StatisticsStr
      Static CurrentStatistics As New StatisticsStr

      Try
         Dim Highest As New Byte
         Dim Lowest As New Byte

         If Refresh Then
            With CurrentStatistics
               .Longest = FileSet().Max(Function(FileO As FileStr) FileO.Size)
               .Shortest = FileSet().Min(Function(FileO As FileStr) FileO.Size)
               ReDim .Highest(.Longest)
               ReDim .Lowest(.Longest)
            End With

            For Position As Integer = 0 To CurrentStatistics.Longest - 1
               Highest = Byte.MinValue
               Lowest = Byte.MaxValue

               For Each File As FileStr In FileSet()
                  With File
                     If Position < .Size Then
                        If .DisplayedData(Position) <= Lowest Then Lowest = .DisplayedData(Position)
                        If .DisplayedData(Position) >= Highest Then Highest = .DisplayedData(Position)
                     End If
                  End With
               Next File

               CurrentStatistics.Lowest(Position) = Lowest
               CurrentStatistics.Highest(Position) = Highest
            Next Position
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return CurrentStatistics
   End Function
End Module
