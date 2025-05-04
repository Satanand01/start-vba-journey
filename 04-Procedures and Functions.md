# 📝 Procedures and Functions
# Contents
1. [Sub Procedures vs Function Procedures](#1-sub-procedures-vs-function-procedures)
2. [Calling procedures](#2-calling-procedures)
3. [Passing arguments (ByVal vs ByRef)](#3-passing-arguments-byval-vs-byref)
4. [Scope of variables (Public, Private, Module-level)](#4-scope-of-variables-public-private-module-level)

---

# 1. Sub Procedures vs Function Procedures
In VBA, **Sub Procedures** and **Function Procedures** are both used to group code into reusable blocks, but they serve different purposes and behave differently. Here's a straight comparison:

---

###  **Sub Procedures (Subroutines)**

* **Does not return a value**
* Used to perform actions/tasks like formatting, updating data, etc.
* Called using `Call` or just the name
* Syntax:

  ```vba
  Sub MySub()
      MsgBox "This is a Sub"
  End Sub
  ```

###  **Function Procedures**

* **Returns a value**
* Used when you need a result (e.g., calculation)
* Can be used in formulas within cells (if it's in a module)
* Syntax:

  ```vba
  Function AddNumbers(a As Double, b As Double) As Double
      AddNumbers = a + b
  End Function
  ```

---

###  Key Differences:

| Feature                 | **Sub Procedure** | **Function Procedure**     |
| ----------------------- | ----------------- | -------------------------- |
| Returns a value?        |  No               |  Yes                       |
| Usable in cell formula? |  No               |  Yes                       |
| Call from worksheet?    |  Not directly     |  If it's a public function |
| Syntax start            | `Sub`             | `Function`                 |
| Syntax end              | `End Sub`         | `End Function`             |

---

### When to Use What?

* Use a **Sub** when you’re just *doing* something (like copying data, updating cells).
* Use a **Function** when you’re *calculating* something and need a **result**.
---
# 2. Calling procedures

Sure, here's a clean explanation **without emojis** on how to call procedures in VBA:

---

### 1. **Calling a Sub Procedure**

**Method 1: Direct Call**

```vba
Sub GreetUser()
    MsgBox "Hello, Satanand!"
End Sub

Sub CallGreet()
    GreetUser ' Calling the Sub procedure
End Sub
```

**Method 2: Using `Call` keyword**

```vba
Sub CallGreet()
    Call GreetUser ' Same as above, just uses the Call keyword
End Sub
```

Note: The `Call` keyword is optional. It is mostly used in older VBA code.

---

### 2. **Calling a Function Procedure**

**From another procedure:**

```vba
Function Square(n As Double) As Double
    Square = n * n
End Function

Sub CallSquare()
    Dim result As Double
    result = Square(5)
    MsgBox "Square is: " & result
End Sub
```

**From a worksheet cell (Function must be Public and in a module):**

```vba
Public Function AddTwo(a As Double, b As Double) As Double
    AddTwo = a + b
End Function
```

Then in Excel, use:

```
=AddTwo(10, 20)
```

---

### 3. **Calling a Sub with Parameters**

```vba
Sub ShowName(name As String)
    MsgBox "Hello, " & name
End Sub

Sub Start()
    ShowName "Satanand"
End Sub
```

---

### Key Rules:

* Sub procedures **do not return** values.
* Function procedures **return** a value and can be used in formulas.
* Only **Function procedures** can be called from worksheet cells.
* Always define functions as `Public` if you want to use them in worksheets.
---
# 3. Passing arguments (ByVal vs ByRef)

In VBA, when you **pass arguments to a procedure**, you can do it in two ways: **ByVal (By Value)** or **ByRef (By Reference)**. These determine **whether the original variable can be changed** by the procedure.

Let’s break it down simply and clearly:

---

### 1. **ByVal (By Value)**

* Passes **a copy** of the variable.
* Changes made inside the procedure **do not affect** the original variable.
* Safe if you don’t want the procedure to mess with your data.

**Example:**

```vba
Sub ChangeByVal(ByVal x As Integer)
    x = x + 10
End Sub

Sub TestByVal()
    Dim a As Integer
    a = 5
    ChangeByVal a
    MsgBox a  ' Output will still be 5
End Sub
```

Here, `x` is a **copy** of `a`, so the change doesn't affect the original `a`.

---

### 2. **ByRef (By Reference)**

* Passes the **actual variable** (memory address).
* Changes made inside the procedure **affect** the original variable.
* Use this when you want the sub/function to **modify** the input.

**Example:**

```vba
Sub ChangeByRef(ByRef x As Integer)
    x = x + 10
End Sub

Sub TestByRef()
    Dim a As Integer
    a = 5
    ChangeByRef a
    MsgBox a  ' Output will be 15
End Sub
```

Now `x` is not a copy — it **is** `a`, so the value changes directly.

---

### Key Differences at a Glance:

| Feature                  | ByVal                    | ByRef                                    |
| ------------------------ | :------------------------| ----------------------------------------|
| Copies value?            | Yes                      | No                                       |
| Changes affect original? | No                       | Yes                                      |
| Default in VBA?          | ByRef (if not specified) | But you should always declare explicitly |

---

### Best Practices:

* Use **ByVal** when the procedure should not alter the original value.
* Use **ByRef** when you want the procedure to update or return values through arguments.
* Always **explicitly declare** `ByVal` or `ByRef` for clarity, even though `ByRef` is default in VBA.

---
### ➡️ [Excel file for better understanding](https://github.com/Satanand01/start-vba-journey/blob/main/Resources-file/ByVal_vs_ByRef_Example.xlsm)
---

# 4. Scope of variables (Public, Private, Module-level)

In VBA, the **scope of a variable** determines **where the variable can be accessed** and **how long it lasts** in memory. There are three main types of scope in VBA:

1. **Public Variables**
2. **Private Variables**
3. **Module-level Variables**

Let's go through each of them:

---

### 1. **Public Variables**

* **Scope:** Available **across all modules** in the entire workbook (for the whole project).
* **Lifetime:** Exists for the duration of the application (i.e., until Excel is closed).
* **Visibility:** Can be accessed from any **Sub**, **Function**, or **Module** within the same workbook.
* **Declaration:** Declared using the `Public` keyword, typically at the **top** of a module.

**Example:**

```vba
Public myPublicVar As Integer  ' Accessible from any module

Sub TestPublic()
    myPublicVar = 10
End Sub
```

**Usage:** Public variables are great when you need to share a value or data across multiple modules or between procedures.

---

### 2. **Private Variables**

* **Scope:** Available **only within the module** in which they are declared. Not visible to other modules.
* **Lifetime:** Exists for the duration of the procedure or until the module is unloaded.
* **Visibility:** Cannot be accessed from other modules, forms, or classes.
* **Declaration:** Declared using the `Private` keyword.

**Example:**

```vba
Private myPrivateVar As Integer  ' Only available in this module

Sub TestPrivate()
    myPrivateVar = 5
    MsgBox myPrivateVar  ' Will work here
End Sub
```

**Usage:** Private variables are useful when you want to **restrict access** to data within a specific module or procedure, improving data encapsulation and preventing accidental changes.

---

### 3. **Module-Level Variables**

* **Scope:** Available **throughout the module** (but not across the entire workbook).
* **Lifetime:** Exists as long as the module is loaded.
* **Visibility:** Accessible by any **Sub** or **Function** within the same module, but not outside of it.
* **Declaration:** Declared with `Dim` at the top of a module but outside any procedures.

**Example:**

```vba
Dim myModuleVar As Integer  ' Available throughout this module only

Sub TestModuleVar()
    myModuleVar = 20
    MsgBox myModuleVar  ' Can access and modify this variable
End Sub
```

**Usage:** Use module-level variables when you need a value to persist across multiple procedures within the same module, but don’t want it to be accessible from other modules.

---

### 🛠 **Summary of Scope**

| Variable Type    | Scope                                | Lifetime                     | Visibility                   |
| ---------------- | ------------------------------------ | ---------------------------- | ---------------------------- |
| **Public**       | Entire workbook (all modules)        | Until Excel is closed        | Accessible everywhere        |
| **Private**      | Only within the module it’s declared | Until procedure finishes     | Limited to its module        |
| **Module-Level** | Entire module                        | Until the module is unloaded | Accessible within its module |

---

### Example to Show All Scopes:

```vba
' Public variable (accessible throughout the workbook)
Public publicVar As Integer

' Module-level variable (accessible only within this module)
Dim moduleVar As Integer

' Private variable (accessible only within this sub)
Private privateVar As Integer

Sub TestScopes()
    publicVar = 10      ' Can be accessed anywhere
    moduleVar = 20      ' Accessible only in this module
    privateVar = 30     ' Accessible only in this sub

    MsgBox "Public: " & publicVar & ", Module: " & moduleVar & ", Private: " & privateVar
End Sub
```

---

### **Best Practices**

* **Public:** Use for global data that needs to be accessed across multiple modules.
* **Private:** Use for data or variables you want to keep **encapsulated** within a procedure or module.
* **Module-level:** Use when you need data shared within a single module, but not outside it.

---

[Go to the top](#-procedures-and-functions)