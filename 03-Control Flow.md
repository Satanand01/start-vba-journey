# 📝 Control Flow
# Contents

1. [If...Then...Else](#1-ifthenelse)
2. [Select Case](#2-select-case)
3. [Loops](#3-loops)
4. [Exit and Continue statements](#4-exit-and-continue-statements)

# 1. If...Then...Else
In VBA, the `If...Then...Else` statement is used for **decision-making**—to execute different blocks of code depending on whether a condition is **True** or **False**.

---

##  Basic Syntax:

```vba
If condition Then
    ' Code to execute if condition is True
Else
    ' Code to execute if condition is False
End If
```

---

##  1. **Single-Line If Statement** (no Else, no End If):

Use when you only need one line of action and no alternative:

```vba
If score > 50 Then MsgBox "Passed"
```

> ⚠️ Avoid this for complex logic—only good for quick conditions.

---

##  2. **Multi-Line If...Then...Else**:

```vba
If score >= 50 Then
    MsgBox "You passed"
Else
    MsgBox "You failed"
End If
```

---

##  3. **If...ElseIf...Else** (Multiple Conditions):

```vba
If score >= 90 Then
    MsgBox "Grade A"
ElseIf score >= 75 Then
    MsgBox "Grade B"
ElseIf score >= 50 Then
    MsgBox "Grade C"
Else
    MsgBox "Fail"
End If
```

---

##  4. **Nested If** (If inside another If):

```vba
If age >= 18 Then
    If citizen = True Then
        MsgBox "Eligible to vote"
    Else
        MsgBox "Not a citizen"
    End If
Else
    MsgBox "Too young to vote"
End If
```

---

##  Comparison Operators Used in If:

| Operator | Description      | Example   |
| -------- | ---------------- | --------- |
| =        | Equal to         | `x = 10`  |
| <>       | Not equal to     | `x <> 10` |
| >        | Greater than     | `x > 10`  |
| <        | Less than        | `x < 10`  |
| >=       | Greater or equal | `x >= 10` |
| <=       | Less or equal    | `x <= 10` |

---

##  Logical Operators (Combine conditions):

| Operator | Description        | Example                    |
| -------- | ------------------ | -------------------------- |
| `And`    | Both must be true  | `If x > 5 And y < 10 Then` |
| `Or`     | Either can be true | `If x > 5 Or y < 10 Then`  |
| `Not`    | Reverses logic     | `If Not isAdmin Then`      |

---

##  Example: Full If...Then...ElseIf...Else

```vba
Sub CheckEligibility()
    Dim age As Integer
    age = InputBox("Enter your age:")
    
    If age >= 18 Then
        MsgBox "You are eligible to vote."
    ElseIf age >= 16 Then
        MsgBox "You can apply for a learner's license."
    Else
        MsgBox "You're too young for both."
    End If
End Sub
```

---

### Use Cases:

* Validating input
* Making decisions based on user choices
* Controlling program flow

--
# 2. Select Case

`Select Case` in VBA is a cleaner, more readable alternative to `If...Then...ElseIf...Else`, especially when you're checking **one variable against multiple possible values**.

---

##  Syntax:

```vba
Select Case expression
    Case value1
        ' Code to run if expression = value1
    Case value2
        ' Code to run if expression = value2
    Case Else
        ' Code to run if no Case matches
End Select
```

---

##  Example 1: Simple Grade Check

```vba
Sub GradeCheck()
    Dim grade As String
    grade = InputBox("Enter your grade (A/B/C):")

    Select Case grade
        Case "A"
            MsgBox "Excellent!"
        Case "B"
            MsgBox "Good Job!"
        Case "C"
            MsgBox "Needs Improvement"
        Case Else
            MsgBox "Invalid Grade"
    End Select
End Sub
```

---

##  Example 2: Using Numeric Ranges

```vba
Sub ScoreEvaluation()
    Dim score As Integer
    score = InputBox("Enter your score:")

    Select Case score
        Case 90 To 100
            MsgBox "Grade A"
        Case 75 To 89
            MsgBox "Grade B"
        Case 50 To 74
            MsgBox "Grade C"
        Case Is < 50
            MsgBox "Fail"
        Case Else
            MsgBox "Invalid Score"
    End Select
End Sub
```

---

##  Comparison: When to Use What?

| Use `If...Then`                   | Use `Select Case`                       |
| --------------------------------- | --------------------------------------- |
| Multiple **different conditions** | Multiple values of **same variable**    |
| Involving **And/Or/Not** logic    | Clean handling of fixed options/ranges  |
| Small, simple checks              | Clean UI for multiple value comparisons |

---

### Tip:

You **can't use multiple variables** directly in `Select Case`. It's designed to evaluate **one expression only**.

---
# 3. Loops
Loops in VBA are used to **repeat a block of code** multiple times. They’re essential when you want to automate repetitive tasks like processing data in rows or running a set of actions until a condition is met.

---

##  Types of Loops in VBA

### 1. **For...Next Loop**

Used when you know **exactly how many times** to loop.

```vba
Dim i As Integer
For i = 1 To 5
    MsgBox "Value of i is: " & i
Next i
```

 Can also step by increments:

```vba
For i = 1 To 10 Step 2
    ' i = 1, 3, 5, 7, 9
Next i
```

 Can loop backward:

```vba
For i = 10 To 1 Step -1
    MsgBox i
Next i
```

---

### 2. **For Each...Next Loop**

Best for looping through **collections** (like cells in a range).

```vba
Dim cell As Range
For Each cell In Range("A1:A5")
    cell.Value = "Done"
Next cell
```

---

### 3. **Do While Loop**

Repeats **as long as** the condition is True.

```vba
Dim i As Integer
i = 1
Do While i <= 5
    MsgBox i
    i = i + 1
Loop
```

---

### 4. **Do Until Loop**

Repeats **until** the condition becomes True.

```vba
Dim i As Integer
i = 1
Do Until i > 5
    MsgBox i
    i = i + 1
Loop
```

> 🔁 `Do While` = keep looping while it's true
> 🔁 `Do Until` = keep looping until it's true

---

### 5. **Do...Loop While / Until**

The code runs **at least once**, then checks the condition after.

```vba
Dim i As Integer
i = 1
Do
    MsgBox i
    i = i + 1
Loop While i <= 5
```

---

##  Exiting a Loop Early

You can break out of a loop using:

```vba
Exit For   ' for For/Next and For Each loops
Exit Do    ' for Do loops
```

Example:

```vba
For i = 1 To 10
    If i = 5 Then Exit For
    MsgBox i
Next i
```

---

##  Skipping Iterations

Use `Continue For` or `GoTo` workaround (VBA doesn't support `Continue` directly):

```vba
For i = 1 To 5
    If i = 3 Then GoTo Skip
    MsgBox i
Skip:
Next i
```

---

## Summary Table

| Loop Type               | Use Case                                |
| ----------------------- | --------------------------------------- |
| `For...Next`            | Known number of iterations              |
| `For Each...Next`       | Looping through ranges, sheets, objects |
| `Do While...Loop`       | Loop while a condition is true          |
| `Do Until...Loop`       | Loop until a condition becomes true     |
| `Do...Loop While/Until` | Always loop at least once               |

---

## Excel examples of using loops in VBA
Perfect. Let's look at **real Excel examples** of using loops in VBA. These are practical and directly useful for your daily MIS or automation tasks.

---

##  Example 1: Loop Through Rows and Highlight Empty Cells

```vba
Sub HighlightEmptyCells()
    Dim cell As Range
    For Each cell In Range("A1:A100")
        If IsEmpty(cell.Value) Then
            cell.Interior.Color = vbYellow
        End If
    Next cell
End Sub
```

 **Use case**: Find and highlight empty entries in a column (like "Name", "Email", etc.).

---

##  Example 2: Sum Values from A1 to A10

```vba
Sub SumValues()
    Dim i As Integer
    Dim total As Double
    total = 0

    For i = 1 To 10
        total = total + Cells(i, 1).Value
    Next i

    MsgBox "Total is " & total
End Sub
```

 Adds up numbers in column A (rows 1 to 10) and shows total.

---
##  Example 3: Loop Through All Sheets and Show Names

```vba
Sub ListAllSheetNames()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        MsgBox "Sheet Name: " & ws.Name
    Next ws
End Sub
```

 Useful if you're managing multiple sheets dynamically.

---

##  Example 4: Loop Through a Range and Color Values > 100

```vba
Sub HighlightAbove100()
    Dim cell As Range
    For Each cell In Range("B2:B20")
        If IsNumeric(cell.Value) And cell.Value > 100 Then
            cell.Interior.Color = RGB(255, 0, 0) ' Red
        End If
    Next cell
End Sub
```

 Good for flagging large amounts or outliers in reports.

---

## Example 5: Copy Non-Empty Cells to Another Sheet

```vba
Sub CopyDataToSheet2()
    Dim i As Long, lastRow As Long, destRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    destRow = 1

    For i = 1 To lastRow
        If Cells(i, 1).Value <> "" Then
            Sheets("Sheet2").Cells(destRow, 1).Value = Cells(i, 1).Value
            destRow = destRow + 1
        End If
    Next i
End Sub
```

Automates data cleaning: filters and moves valid data to another sheet.

---

##  Example 6: Use `Do While` to Find First Empty Row

```vba
Sub FindFirstEmptyRow()
    Dim i As Long
    i = 1
    Do While Cells(i, 1).Value <> ""
        i = i + 1
    Loop
    MsgBox "First empty row is: " & i
End Sub
```

 Can be used to dynamically append data to the bottom of a list.

---

# 4. Exit and Continue statements

In VBA, `Exit` and `Continue` (kind of) are used to **control the flow inside loops or procedures**, especially when you want to stop or skip part of the loop under certain conditions.

---

##  `Exit` Statement

###  Purpose: Immediately **exit** a loop or procedure, no matter the loop's original ending condition.

###  `Exit For` – breaks out of a `For` or `For Each` loop

```vba
Sub ExitForExample()
    Dim i As Integer
    For i = 1 To 10
        If i = 5 Then Exit For
        Debug.Print i
    Next i
End Sub
```

 Output: 1, 2, 3, 4 — loop stops at `i = 5`

---

###  `Exit Do` – breaks out of a `Do While` or `Do Until` loop

```vba
Sub ExitDoExample()
    Dim i As Integer
    i = 1
    Do While i <= 10
        If i = 5 Then Exit Do
        Debug.Print i
        i = i + 1
    Loop
End Sub
```

 Output: 1, 2, 3, 4

---

###  `Exit Sub` – stops a procedure

```vba
Sub CheckValue()
    If Range("A1").Value = "" Then
        MsgBox "No value in A1"
        Exit Sub
    End If
    MsgBox "Value is: " & Range("A1").Value
End Sub
```

If A1 is empty, rest of the code is skipped.

---

###  `Exit Function` – exits a function and returns control to the caller

```vba
Function GetAgeMessage(age As Integer) As String
    If age < 0 Then
        GetAgeMessage = "Invalid age"
        Exit Function
    End If
    GetAgeMessage = "Valid age: " & age
End Function
```

---

##  VBA Doesn't Have `Continue For` or `Continue Do` (like in Python or C#)

### But you can **simulate Continue** using a `GoTo` label:

```vba
Sub ContinueForSimulated()
    Dim i As Integer
    For i = 1 To 5
        If i = 3 Then GoTo SkipLoop
        Debug.Print "i is: " & i
SkipLoop:
    Next i
End Sub
```

 Output: Skips when `i = 3`

---

##  Summary

| Statement       | Used In           | Purpose                            |
| --------------- | ----------------- | ---------------------------------- |
| `Exit For`      | `For`, `For Each` | Exit the loop immediately          |
| `Exit Do`       | `Do While/Until`  | Exit the loop immediately          |
| `Exit Sub`      | Sub procedures    | Stop executing the sub             |
| `Exit Function` | Functions         | Stop executing and return a value  |
| `GoTo`          | Any loop          | Simulate `Continue` (jump forward) |

---

[Go to the top](#contents)