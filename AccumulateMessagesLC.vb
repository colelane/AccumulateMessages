Option Strict On
Option Explicit On
Option Compare Text

'Lane Coleman
'RCET 0265
'Fall 2020
'Accumulate Messages
'https://github.com/colelane/AccumulateMessages.git


Module AccumulateMessagesLC

    Sub Main()
        Dim message As String
        Dim userInput As String

        Console.WriteLine("Enter messages and they will be stored." & vbNewLine _
                          & "Enter 'call' at any time to read stored messages." & vbNewLine _
                          & "Enter 'clear' at any time to delete messages")
        Do
            userInput = Console.ReadLine()
            If userInput = "call" Then
                MsgBox(message)
            ElseIf userInput = "clear" Then
                message = AccumulateMessage("", True)
            Else
                message = AccumulateMessage(userInput, False)
            End If
        Loop

    End Sub

    Function AccumulateMessage(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static userMessage As String
        'static means that when you come back to the function later, previous values are saved.
        'the if statements store a message that can be later called up.  concatenates new messages onto the original usermessage
        If clear Then
            userMessage = ""
        Else
            userMessage &= newMessage & vbNewLine

        End If
        Return userMessage
    End Function

End Module
