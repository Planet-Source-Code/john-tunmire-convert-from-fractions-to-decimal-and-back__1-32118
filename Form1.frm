VERSION 5.00
Begin VB.Form frmConvertor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Fractions and Decimals"
   ClientHeight    =   4995
   ClientLeft      =   2595
   ClientTop       =   1755
   ClientWidth     =   6975
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6975
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "Convert Decimal to Fraction"
      Height          =   510
      Left            =   2115
      TabIndex        =   3
      Top             =   2445
      Width           =   2850
   End
   Begin VB.TextBox txtTop 
      Height          =   330
      Left            =   2115
      TabIndex        =   2
      Top             =   1395
      Width           =   2850
   End
   Begin VB.CommandButton cmdFraction 
      Caption         =   "Convert Fraction to Decimal"
      Height          =   510
      Left            =   2115
      TabIndex        =   1
      Top             =   1830
      Width           =   2850
   End
   Begin VB.TextBox txtBottom 
      Height          =   330
      Left            =   2115
      TabIndex        =   0
      Top             =   3330
      Width           =   2850
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Input:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   1125
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2115
      TabIndex        =   4
      Top             =   3105
      Width           =   1275
   End
End
Attribute VB_Name = "frmConvertor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function ConvertFractionToDecimal(str As String) As String
Dim strFraction As String
Dim strWholeNumber As String
Dim strNumerator As String
Dim StrDenominator As String
Dim intFirst As Integer
Dim intLength As Long

If InStr(1, str, "/") = 0 Then Exit Function

If InStr(1, str, " ") Then
    strWholeNumber = Mid(str, 1, InStr(1, str, " ") - 1)
Else
    strWholeNumber = "0"
End If

If strWholeNumber <> "0" Then
    str = Trim(Mid(str, InStr(1, str, " ")))
    intLength = InStr(1, str, "/") - 1
    strNumerator = Mid(str, 1, intLength)
    StrDenominator = Mid(str, InStr(1, str, "/") + 1)
Else
    str = Trim(str)
    intLength = InStr(1, str, "/") - 1
    strNumerator = Mid(str, 1, intLength)
    StrDenominator = Mid(str, InStr(1, str, "/") + 1)
End If
ConvertFractionToDecimal = Val(strWholeNumber) + Val(strNumerator) / Val(StrDenominator)


End Function

Private Sub cmdDecimal_Click()
txtBottom.Text = ConvertDecimalToFraction(txtTop.Text)
End Sub

Private Sub cmdFraction_Click()
txtBottom.Text = ConvertFractionToDecimal(Trim(txtTop.Text))
End Sub

Public Function ConvertDecimalToFraction(str As String) As String
Dim intCountDecimalPoints As Integer
Dim strWholeNumber As String
Dim strDecimal As String
Dim StrDenominator As String
Dim intDecimalMarker As Integer
Dim intWholeNumberLength As Integer
Dim temp As String

If InStr(1, str, ".") = 0 Then Exit Function
str = Trim(str)
intDecimalMarker = InStr(1, str, ".")
intCountDecimalPoints = Len(Mid(str, InStr(1, str, ".") + 1))

intWholeNumberLength = intDecimalMarker - 1
strWholeNumber = Mid(str, 1, intWholeNumberLength)
strDecimal = Mid(str, intDecimalMarker + 1)

temp = CheckForRepeatingDecimal(strDecimal)

If temp <> "0" Then
    strDecimal = temp
    intCountDecimalPoints = Len(strDecimal)
End If

StrDenominator = "1"
For i = 1 To intCountDecimalPoints
    StrDenominator = StrDenominator & "0"
Next

If temp <> "0" Then
    StrDenominator = StrDenominator - 1
End If


If strWholeNumber = "0" Then
    ConvertDecimalToFraction = Trim(ReduceToLCD(Val(strDecimal), Val(StrDenominator)))
Else
    ConvertDecimalToFraction = Trim(strWholeNumber & " " & ReduceToLCD(Val(strDecimal), Val(StrDenominator)))
End If
End Function


Public Function ReduceToLCD(dblNumerator, dblDenominator As Double) As String
Dim i As Long



For i = 2 To dblDenominator
    If i > dblDenominator Then GoTo fini
    If dblNumerator Mod i = 0 And dblDenominator Mod i = 0 Then
        dblNumerator = dblNumerator / i
        dblDenominator = dblDenominator / i
        i = 1
    End If
Next

fini:
ReduceToLCD = dblNumerator & "/" & dblDenominator

End Function


Public Function CheckForRepeatingDecimal(strDecimal) As String
Dim i As Integer
Dim j As Integer
Dim temp As String
Dim strcheck As String
Dim strCheckPrevious As String
Dim intCount As Integer

For i = Len(strDecimal) To 1 Step -1
    strcheck = Mid(strDecimal, 1, i)
    If Len(strcheck) > 0 Then
        For j = 1 To Len(strDecimal) Step Len(strcheck)
            If InStr(j, strDecimal, strcheck) Then
                intCount = intCount + 1
            End If
        Next
        If intCount > 1 Then
            CheckForRepeatingDecimal = strcheck
        Exit Function
        End If
        intCount = 0
    End If
Next

    CheckForRepeatingDecimal = "0"
    
End Function
