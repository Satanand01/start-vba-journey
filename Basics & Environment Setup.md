# 📝 Basics & Environment Setup

# Contents :
1. [What is VBA(Visual basic application)?](#1-what-is-vba-visual-basic-application)
2. [VBA Editor (VBE) tour](#2-vba-editor-vbe-tour)
3. [Modules, procedures, and functions](#3-modules-procedures-and-functions)
4. [Macros vs VBA](#4-macros-vs-vba)
5. [Recording & modifying the macros](#5-recording--modifying-the-macros)
---
# 1. What is VBA (Visual basic application)

**VBA (Visual Basic for Applications)** is a programming language developed by Microsoft that is used to automate tasks and create custom functionality in Microsoft Office applications like Excel, Word, Outlook, and Access.

### In simple terms:

It lets you **write code inside Excel (or other Office apps)** to automate repetitive tasks, build custom functions, create forms, interact with databases, or even control other Office applications.

### Key features of VBA:

* **Automation** – e.g., automatically format data, generate reports, or send emails from Excel.
* **Custom Functions** – Create your own Excel functions using VBA.
* **User Forms** – Build interactive forms with buttons, text boxes, etc.
* **Macros** – VBA is what runs behind the scenes when you record a macro.
* **Event-driven** – You can trigger actions based on user activity (like clicking a button or changing a cell).

### Example:

```vba
Sub HelloWorld()
    MsgBox "Hello, Satanand! Welcome to VBA."
End Sub
```

This small script shows a pop-up message when run.
---

# 2. VBA Editor (VBE) tour

Got it — no emojis. Here's a clean and professional **VBA Editor (VBE) Tour** without any fluff:


### How to Open the VBA Editor:

* Press `Alt + F11` in Excel
* Or go to the **Developer** tab → click **Visual Basic**

---

### Main Sections of the VBA Editor (VBE):

#### 1. **Project Explorer (Top-Left)**

* Displays all open workbooks and their components like:

  * Worksheets
  * ThisWorkbook
  * Modules
  * UserForms (if any)
* Acts like a navigation panel
* Shortcut: `Ctrl + R` to show/hide

#### 2. **Code Window (Center Area)**

* This is where you write and edit your VBA code
* When you double-click any item in Project Explorer, the corresponding code window opens
* You write procedures (`Sub`, `Function`) here

#### 3. **Properties Window (Bottom-Left)**

* Lets you view and change properties of selected objects
* Useful when working with UserForms or controls
* Shortcut: `F4` to show/hide

#### 4. **Immediate Window (Bottom Center)**

* Used for testing code snippets or checking values during debugging
* Example: type `? Range("A1").Value` and press Enter to see the cell's value
* Shortcut: `Ctrl + G` to show/hide

#### 5. **Menu Bar and Toolbar (Top)**

* Contains common options:

  * File, Edit, View, Insert, Debug, Run, Tools
  * Run button (F5)
  * Break (Ctrl + Break)
  * Reset, Step Into, Step Over, Step Out (for debugging)

#### 6. **Modules (Under Project Explorer)**

* You can insert a Module using `Insert > Module`
* Modules are where you store most of your macros and general-purpose code
* You can create multiple modules if needed to keep your code organized

---

### Summary Table:

| Section           | Purpose                         | Shortcut          |
| ----------------- | ------------------------------- | ----------------- |
| Project Explorer  | Navigate between sheets/modules | Ctrl + R          |
| Code Window       | Write and edit your VBA code    | Double-click item |
| Properties Window | Change properties of objects    | F4                |
| Immediate Window  | Run test commands/debug code    | Ctrl + G          |
| Toolbar/Menu      | Run, debug, insert, save, etc.  | F5 (Run)          |

---
# 3. Modules, procedures, and functions

Alright, let’s break down **Modules**, **Procedures**, and **Functions** in VBA — these are the building blocks of your automation logic. You must understand these clearly if you want to master VBA.



## 1. **Modules** – The Container

### What is it?

A **Module** is like a blank notebook where you write your VBA code. It stores **Procedures** (macros) and **Functions**.

### Types of Modules:

* **Standard Modules**: Created manually using `Insert > Module`. Used for general-purpose code.
* **Object Modules**: Attached to sheets, `ThisWorkbook`, or UserForms. Used to handle events (like when a cell changes or workbook opens).

### Example:

```vba
' This is a standard module
Sub MyMacro()
    MsgBox "This is a macro inside a module."
End Sub
```

---

## 2. **Procedures** – The Action Block

### What is it?

A **Procedure** is a block of code that performs an action. There are two types:

### a) **Sub Procedures** (`Sub`)

* Performs actions like selecting ranges, copying data, showing messages, etc.
* Can be run manually or triggered by events/buttons.
* Doesn't return a value.

#### Example:

```vba
Sub GreetUser()
    MsgBox "Hello, Satanand!"
End Sub
```

### b) **Event Procedures**

* Special `Sub`s that run when something happens (like opening the workbook or clicking a button).
* Written inside `ThisWorkbook` or sheet modules.

#### Example:

```vba
Private Sub Workbook_Open()
    MsgBox "Welcome to this file!"
End Sub
```

---

## 3. **Functions** – Return a Value

### What is it?

A **Function** is similar to a procedure, but it **returns a value**. You can call it in formulas or other code.

#### Example:

```vba
Function AddNumbers(a As Double, b As Double) As Double
    AddNumbers = a + b
End Function
```

You can use this in code like:

```vba
Sub TestFunction()
    MsgBox AddNumbers(10, 20)  ' Outputs: 30
End Sub
```

Or even call it directly in a worksheet formula (if it’s a public function).

---

## Summary:

| Term      | Purpose                        | Returns Value | Can Run Directly | Where It’s Written        |
| --------- | ------------------------------ | ------------- | ---------------- | ------------------------- |
| Module    | Container for code             | No            | No               | Standard or Object Module |
| Procedure | Performs tasks (Sub)           | No            | Yes              | Module or Sheet/Workbook  |
| Function  | Calculates and returns a value | Yes           | No (usually)     | Module                    |

---
# 4. Macros vs VBA

Sure, here's a **simple and short** comparison:


### **Macro**:

* A **macro** is a recorded or written **set of instructions** to automate tasks in Excel.
* You can **record** a macro without writing code.
* Stored as a **Sub Procedure**.

### **VBA** (Visual Basic for Applications):

* **VBA** is the **programming language** used to write or edit macros.
* It gives you **full control**: conditions, loops, custom functions, forms, etc.
* Macros run because of VBA code.

---

### In short:

* **Macro** = What you run
* **VBA** = The language behind it

---

# 5. Recording & modifying the macros

Perfect — here's exactly how you can **record** and **modify** macros in Excel:


## **Step 1: Enable the Developer Tab (if not visible)**

1. Go to **File > Options > Customize Ribbon**
2. On the right, check **Developer**, then click OK

---

## **Step 2: Record a Macro**

1. Go to the **Developer tab**
2. Click **Record Macro**
3. Fill in:

   * **Macro name** (no spaces)
   * **Shortcut key** (optional)
   * **Store macro in** (This Workbook is fine for now)
4. Click **OK**
5. Do some actions (e.g., type in a cell, apply bold, change color, etc.)
6. Click **Stop Recording** (from Developer tab)

---

## **Step 3: Modify the Macro (VBA Code)**

1. Press **`Alt + F11`** to open the **VBA Editor**
2. In **Project Explorer**, find:

   * `VBAProject (YourWorkbookName)`
   * Expand **Modules**
   * Double-click **Module1**
3. You’ll see code like this:

```vba
Sub Macro1()
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Hello"
    Range("A1").Font.Bold = True
End Sub
```

You can now **edit** this code. For example, change `"Hello"` to `"Satanand"` or remove the `Select` line to make it cleaner.

---

## Example: Modified Macro

```vba
Sub Macro1()
    Range("A1").Value = "Satanand"
    Range("A1").Font.Bold = True
End Sub
```

---

## Summary:

| Task            | How                            |
| --------------- | ------------------------------ |
| Record Macro    | Developer tab → Record Macro   |
| View/Edit Macro | Alt + F11 → Module1            |
| Stop Recording  | Developer tab → Stop Recording |
| Run Macro       | Developer tab → Macros → Run   |

---
