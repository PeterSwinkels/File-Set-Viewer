'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Environment
Imports System.Text

'This module contains this program's core procedures.
Public Module CoreModule
   'This procedure returns the user's selection from the specified options.
   Public Function Choose(Prompt As String, Choices As String, Left As Integer, Top As Integer) As Char
      Try
         Dim Choice As New Char

         Console.ForegroundColor = ConsoleColor.Gray
         Console.SetCursorPosition(Left, Top)
         Console.Write(Prompt)
         Do
            Choice = Console.ReadKey(intercept:=True).KeyChar
            If Choices?.Contains(Choice) Then Exit Do
         Loop
         Console.Write(Choice)
         Return Choice
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure displays the help for this program.
   Public Sub DisplayHelp()
      Try
         Console.Clear()
         Console.BackgroundColor = ConsoleColor.Gray
         Console.ForegroundColor = ConsoleColor.Black
         Console.WriteLine($" { My.Application.Info.Title} - Help ")
         Console.BackgroundColor = ConsoleColor.Black
         Console.ForegroundColor = ConsoleColor.Gray
         Console.WriteLine()
         Console.WriteLine(" Key:         Function:")
         Console.WriteLine(" End          Jump to the end of the longest file.")
         Console.WriteLine(" Escape       Close this help.")
         Console.WriteLine(" F1           Display this help.")
         Console.WriteLine(" F2           Sort files by the data in the left most column displayed.")
         Console.WriteLine(" F3           Sort files by their name.")
         Console.WriteLine(" F4           Sort files by their extension.")
         Console.WriteLine(" F5           Sort files by their size.")
         Console.WriteLine(" F6           Turn displaying basic statistics for the file set on/off.")
         Console.WriteLine(" F7           Turn displaying the files' contents as text on/off.")
         Console.WriteLine(" F8           Invert the current fileset's bits.")
         Console.WriteLine(" F9           Toggle between displaying the file data backwards or forwards.")
         Console.WriteLine(" F10          Display/manage markers.")
         Console.WriteLine(" F12          Export the current fileset's data.")
         Console.WriteLine(" Home         Jump to the start of the files.")
         Console.WriteLine(" Page Down    Jump one ""screen"" forward.")
         Console.WriteLine(" Page Up      Jump one ""screen"" back.")
         Console.WriteLine(" Down         Scroll down in the file data displayed.")
         Console.WriteLine(" Left         Scroll left in the file data displayed.")
         Console.WriteLine(" Right        Scroll right in the file data displayed.")
         Console.WriteLine(" Up           Scroll up in the file data displayed.")
         Console.WriteLine(" +/-          Shift the current fileset's bytes.")
         Do Until Console.ReadKey(intercept:=True).Key = ConsoleKey.Escape : Loop
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure returns the position of the specified bytes.
   Public Function FindBytes(Offset As Integer, ToSearch() As Byte, ToFind() As Byte) As Integer?
      Try
         Dim MatchCount As Integer = 0
         Dim Position As Integer = Offset

         Do While Position <= ToSearch.GetUpperBound(0)
            If ToSearch(Position) = ToFind(MatchCount) Then MatchCount += 1 Else MatchCount = 0
            If MatchCount = ToFind.Length Then Return Position - MatchCount + 1
            Position += 1
         Loop
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure displays a prompt and returns the user's input.
   Public Function GetInput(Prompt As String, Optional Filter As String = Nothing) As String
      Dim Input As New StringBuilder

      Try
         Dim KeyStroke As New ConsoleKeyInfo
         Dim x As Integer = Console.CursorLeft
         Dim y As Integer = Console.CursorTop

         Do
            Console.SetCursorPosition(left:=x, top:=y)
            Console.Write($"{Prompt}{Input}_ ")
            Do Until Console.KeyAvailable
            Loop
            Console.SetCursorPosition(left:=x, top:=y)
            Console.Write($"{Prompt}{Input} ")
            KeyStroke = Console.ReadKey(intercept:=True)
            Select Case KeyStroke.Key
               Case ConsoleKey.Backspace
                  If Input.Length > 0 Then Input.Remove(Input.Length - 1, 1)
               Case ConsoleKey.Enter
                  Console.WriteLine()
                  Return Input.ToString()
               Case ConsoleKey.Escape
                  Console.WriteLine()
                  Return Nothing
               Case Else
                  If Filter = Nothing OrElse Filter.Contains(KeyStroke.KeyChar) Then Input.Append(KeyStroke.KeyChar)
            End Select
         Loop
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Input.ToString()
   End Function

   'This procedure handles any errors that occur.
   Public Sub HandleError(ExceptionO As Exception)
      Try
         Console.WriteLine()
         Console.WriteLine("Error:")
         Console.WriteLine(ExceptionO.Message)
         Console.WriteLine("Press Enter to continue.")
         Console.WriteLine()
         Do Until Console.ReadKey(intercept:=True).Key = ConsoleKey.Enter : Loop
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure returns information about this program.
   Public Function ProgramInformation() As String
      Try
         With My.Application.Info
            Return $"{ .Title}, v{ .Version} - by: { .CompanyName}"
         End With
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
