Attribute VB_Name = "Assert"
Private Const PASS As String = "Pass: "
Private Const FAIL As String = "Fail: "

Public Sub DoesNotThrow(message As String, ParamArray methodNames() As Variant)
    On Error GoTo Catch
    Dim methodNumber As Integer
    For methodNumber = LBound(methodNames) To UBound(methodNames)
        Application.Run methodNames(methodNumber)
    Next methodNumber
    passMessage message
    Exit Sub
Catch:
    assertError message
End Sub

Public Sub Equal(x As Variant, y As Variant, message As String)
    On Error GoTo Catch
    If x = y Then
        passMessage message
    Else
        failMessage message
        failReason "Expected " & CStr(x) & " to equal " & CStr(y) & "."
    End If
    Exit Sub
Catch:
    assertError message
End Sub

Public Sub Exists(obj As Object, message As String)
    On Error GoTo Catch
    If Not obj Is Nothing Then
        passMessage message
    Else
        failMessage message
        failReason "The argument passed was equal to Nothing."
    End If
    Exit Sub
Catch:
    assertError message
End Sub

Public Sub IsNothing(obj As Object, message As String)
    On Error GoTo Catch
    If obj Is Nothing Then
        passMessage message
    Else
        failMessage message
        failReason "Expected " & CStr(obj1) & " to equal Nothing."
    End If
    Exit Sub
Catch:
    assertError message
End Sub

Public Sub IsFalse(condition As Variant, message As String)
    Equal condition, False, message
End Sub

Public Sub IsTrue(condition As Variant, message As String)
    Equal condition, True, message
End Sub

Public Sub ObjectEquals(obj1 As Object, obj2 As Object, message As String)
On Error GoTo Catch
    If obj1 Is obj2 Then
        passMessage message
    Else
        failMessage message
        failReason "Expected " & CStr(obj1) & " to equal " & CStr(obj2) & "."
    End If
    Exit Sub
Catch:
    assertError message
End Sub

Private Sub assertError(message As String)
    failMessage message
    failReason "This test threw an Error of type " & Err.Description & " from " & Err.Source & "."
End Sub

Private Static Sub passMessage(message As String)
    Debug.Print PASS & message
End Sub

Private Static Sub failMessage(message As String)
    Debug.Print FAIL & message
End Sub

Private Static Sub failReason(message As String)
    Debug.Print "  " & message & Chr(10)
End Sub


