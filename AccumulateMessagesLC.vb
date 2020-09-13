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
        Dim clearData As Boolean

        Console.WriteLine($"Enter messages and they will be stored. 
Enter 'call' at any time to read stored messages.
Enter 'clear' at any time to delete messages")
        Do
            userInput = Console.ReadLine()
            If userInput = "call" Then
                MsgBox(message)
            ElseIf userInput = "clear" Then
                clearData = True
            End If
            message = AccumulateMessage(userInput, clearData)



            clearData = False
        Loop

    End Sub
    Function AccumulateMessage(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static userMessage As String
        'static means that when you come back to the function later, previous values are saved.
        'the if statements store a message that can be later called up.  concatenates new messages onto the original usermessage
        If clear Then
            userMessage = ""
        ElseIf newMessage = "call" Then
            'this is intentionally left blank.  Didn't want to have the message box display the word call.
        Else
            userMessage &= newMessage & vbNewLine


        End If
        Return userMessage


    End Function

End Module
