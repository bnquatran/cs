Sub FormatFormulasWithSpaces()

    ' --------------------------------------------------------------------
    ' This macro reformats formulas across all worksheets in the workbook.
    ' It normalizes spacing around operators, handles protected sheets,
    ' and preserves unary minus formatting (e.g., "-1" or "-A1" stays together).
    ' Sheets with math symbols in their names are skipped to avoid
    ' breaking external references.
    ' --------------------------------------------------------------------

    Dim ws As Worksheet
    Dim cell As Range
    Dim formulaText As String
    Dim mathOps As Variant
    Dim op As Variant
    Dim i As Integer

    ' Password-handling variables
    Dim wasProtected As Boolean
    Dim pw As String
    Dim commonPw As String
    Dim useCommonPw As Boolean
    Dim commonPwSet As Boolean
    Dim unprotected As Boolean

    Dim body As String
    Dim skippedSheets As String

    mathOps = Array("+", "*", "/", "<", ">", "&")
    skippedSheets = ""

    For Each ws In ThisWorkbook.Worksheets

        ' ============================================================
        ' Skip sheets whose names contain math symbols
        ' ============================================================
        Dim invalidChars As Variant
        Dim badChar As Variant
        invalidChars = Array("+", "-", "*", "/", "<", ">", "=")

        For Each badChar In invalidChars
            If InStr(ws.Name, badChar) > 0 Then
                skippedSheets = skippedSheets & IIf(skippedSheets = "", "", ", ") & ws.Name
                GoTo NextSheet
            End If
        Next badChar

        wasProtected = ws.ProtectContents
        unprotected = False
        pw = ""

        If wasProtected Then

            ' ------------------------------------------------------------
            ' Try unprotecting without password first
            ' ------------------------------------------------------------
            On Error Resume Next
            ws.Unprotect
            If Err.Number = 0 Then
                unprotected = True
                pw = ""
            Else
                Err.Clear
                On Error GoTo 0

                ' Try common password if set
                If useCommonPw And commonPwSet Then
                    On Error Resume Next
                    ws.Unprotect Password:=commonPw
                    If Err.Number = 0 Then
                        unprotected = True
                        pw = commonPw
                    Else
                        skippedSheets = skippedSheets & IIf(skippedSheets = "", "", ", ") & ws.Name
                        Err.Clear
                    End If
                    On Error GoTo 0
                Else
                    ' Prompt user for password if needed
                    Dim response As VbMsgBoxResult
                    Do While Not unprotected
                        pw = InputBox( _
                            "Sheet '" & ws.Name & "' is protected." & vbCrLf & _
                            "Enter password to process this sheet," & vbCrLf & _
                            "or leave blank and press OK to skip:", _
                            "Sheet Password Required")

                        If pw = "" Then
                            skippedSheets = skippedSheets & IIf(skippedSheets = "", "", ", ") & ws.Name
                            GoTo NextSheet
                        End If

                        On Error Resume Next
                        ws.Unprotect Password:=pw
                        If Err.Number = 0 Then
                            unprotected = True
                        Else
                            MsgBox "Incorrect password for sheet '" & ws.Name & "'. Please try again.", vbExclamation
                            Err.Clear
                        End If
                        On Error GoTo 0
                    Loop

                    ' Ask if same password should be reused
                    If Not commonPwSet And Not useCommonPw Then
                        response = MsgBox( _
                            "Do all password-protected sheets use the SAME password you just entered?", _
                            vbYesNo + vbQuestion, _
                            "Use Common Password?")
                        If response = vbYes Then
                            useCommonPw = True
                            commonPwSet = True
                            commonPw = pw
                        End If
                    End If
                End If
            End If
        Else
            unprotected = True
        End If

        ' ====================================================================
        ' PROCESS FORMULAS ON THIS SHEET
        ' ====================================================================
        If unprotected Then

            For Each cell In ws.UsedRange
                If cell.HasFormula Then

                    formulaText = Trim(cell.Formula)

                    ' Add a space after "=" at start
                    If Left(formulaText, 1) = "=" Then
                        formulaText = "= " & LTrim(Mid(formulaText, 2))
                    End If

                    ' Normalize equals and relational operators
                    If Len(formulaText) > 2 Then
                        body = Mid(formulaText, 3)

                        body = Replace(body, " = ", "=")
                        body = Replace(body, " =", "=")
                        body = Replace(body, "= ", "=")
                        body = Replace(body, "=", " = ")

                        formulaText = "= " & body
                    End If

                    ' Relational operators
                    formulaText = Replace(formulaText, "<  >", "<>")
                    formulaText = Replace(formulaText, "< >", "<>")
                    formulaText = Replace(formulaText, "<>", " <> ")

                    formulaText = Replace(formulaText, "<  =", "<=")
                    formulaText = Replace(formulaText, "< =", "<=")
                    formulaText = Replace(formulaText, "<=", " <= ")

                    formulaText = Replace(formulaText, ">  =", ">=")
                    formulaText = Replace(formulaText, "> =", ">=")
                    formulaText = Replace(formulaText, ">=", " >= ")

                    ' Standard operators
                    For i = LBound(mathOps) To UBound(mathOps)
                        op = mathOps(i)
                        formulaText = Replace(formulaText, " " & op & " ", op)
                        formulaText = Replace(formulaText, " " & op, op)
                        formulaText = Replace(formulaText, op & " ", op)
                        formulaText = Replace(formulaText, op, " " & op & " ")
                    Next i

                    ' -------------------------------
                    ' Handle minus signs cleanly
                    ' -------------------------------
                    Dim j As Long
                    Dim output As String
                    Dim ch As String
                    Dim prev As String
                    Dim nextCh As String

                    output = ""

                    For j = 1 To Len(formulaText)
                        ch = Mid(formulaText, j, 1)

                        If ch = "-" Then
                            prev = IIf(j > 1, Mid(formulaText, j - 1, 1), "")
                            nextCh = IIf(j < Len(formulaText), Mid(formulaText, j + 1, 1), "")

                            ' Unary vs binary minus
                            If prev = "" Or InStr(" (,+-*/^=<>&", prev) > 0 Then
                                output = output & " -"
                            Else
                                output = output & " - "
                            End If
                        Else
                            output = output & ch
                        End If
                    Next j

                    formulaText = output

                    ' Commas
                    formulaText = Replace(formulaText, ", ", ",")
                    formulaText = Replace(formulaText, ",", ", ")

                    ' Carets
                    formulaText = Replace(formulaText, " ^ ", "^")
                    formulaText = Replace(formulaText, " ^", "^")
                    formulaText = Replace(formulaText, "^ ", "^")

                    ' Collapse multiple spaces
                    Do While InStr(formulaText, "  ") > 0
                        formulaText = Replace(formulaText, "  ", " ")
                    Loop

                    ' Apply reformatted formula
                    cell.Formula = formulaText

                End If
            Next cell
        End If

        ' Reprotect sheet if needed
        If wasProtected And unprotected Then
            If pw <> "" Then
                ws.Protect Password:=pw
            ElseIf useCommonPw And commonPwSet Then
                ws.Protect Password:=commonPw
            Else
                ws.Protect
            End If
        End If

NextSheet:
    Next ws

    ' Report skipped sheets
    If skippedSheets <> "" Then
        MsgBox "The following sheets were skipped (password not provided, did not match, or name contains math symbols):" & _
               vbCrLf & skippedSheets, vbInformation, "Skipped Sheets"
    End If

End Sub
