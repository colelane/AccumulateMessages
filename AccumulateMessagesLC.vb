Option Strict On
Option Explicit On
Option Compare Text

'Lane Coleman
'RCET 0265
'Fall 2020
'Accumulate Messages


Module AccumulateMessagesLC

    Sub Main()
        Dim message As String
        Dim userInput As String
        Dim clearData As Boolean

        Do
            userInput = Console.ReadLine()
            If userInput = "call" Then
                MsgBox(message)
            ElseIf userinput = "clear" Then
                clearData = True
            End If
            message = AccumulateMessage(userInput, clearData)

            'Console.WriteLine(message)

            clearData = False
        Loop

    End Sub
    Function AccumulateMessage(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static userMessage As String

        If clear Then
            userMessage = ""
        ElseIf newMessage = "call" Then

        Else
            userMessage &= newMessage & vbNewLine


        End If
        Return userMessage


    End Function

End Module
