Sub EditEntryByCode()
    Dim wsData As Worksheet
    Dim wsForm As Worksheet
    Dim searchCode As String
    Dim foundCell As Range
    Dim itemName As String
    Dim price As Variant
    Dim currentUser As String
    
    ' Track current user
    currentUser = Application.UserName
    
    Set wsData = ThisWorkbook.Sheets("Sheet2")
    Set wsForm = ThisWorkbook.Sheets("Sheet1")
    
    ' Check if form is being used by another user
    If wsForm.Range("B1").Value <> "" And wsForm.Range("B1").Value <> currentUser Then
        MsgBox "Form is currently being edited by another user. Please wait.", vbExclamation
        Exit Sub
    End If
    
    searchCode = wsForm.Range("B2").Value

    ' Check for empty ItemCode
    If searchCode = "" Then
        MsgBox "Please enter an Item Code in B2.", vbExclamation
        Exit Sub
    End If
    
    ' Validate that ItemCode is numeric
    If Not IsNumeric(searchCode) Then
        MsgBox "Item Code must be a valid number.", vbExclamation
        Exit Sub
    End If
    
    ' Search in column A of Sheet2 for the ItemCode
    Set foundCell = wsData.Range("A:A").Find(What:=searchCode, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' ItemCode not found
    If foundCell Is Nothing Then
        MsgBox "Item Code not found!", vbCritical
    Else
        ' Load data into the form for editing
        If wsForm.Range("B3").Value = "" And wsForm.Range("B4").Value = "" Then
            wsForm.Range("B1").Value = currentUser ' Lock form with current user
            
            wsForm.Range("B3").Value = foundCell.Offset(0, 1).Value ' Load ItemName
            wsForm.Range("B4").Value = foundCell.Offset(0, 2).Value ' Load Price
            
            MsgBox "Record loaded. Edit and click the button again to save changes.", vbInformation
        Else
            ' Validate the inputs
            itemName = wsForm.Range("B3").Value
            price = wsForm.Range("B4").Value
            
            If itemName = "" Then
                MsgBox "Item Name cannot be empty.", vbExclamation
                Exit Sub
            End If
            
            ' Validate price
            If Not IsNumeric(price) Then
                MsgBox "Price must be a valid number.", vbExclamation
                Exit Sub
            End If
            
            ' Save changes back to Sheet2
            foundCell.Offset(0, 1).Value = itemName
            foundCell.Offset(0, 2).Value = price
            
            MsgBox "Changes saved to database!", vbInformation
            
            ' Clear form fields and unlock
            wsForm.Range("B2:B4").ClearContents
            wsForm.Range("B1").ClearContents
        End If
    End If
End Sub

# Explaining the above code step by step

Let's break down the **VBA code** you've provided, step-by-step:

---

### 1. **Variable Declarations**

```vba
Dim wsData As Worksheet
Dim wsForm As Worksheet
Dim searchCode As String
Dim foundCell As Range
Dim itemName As String
Dim price As Variant
Dim currentUser As String
```

* **`wsData`**: A variable to represent the **Sheet2** where the data is stored.
* **`wsForm`**: A variable to represent **Sheet1**, which contains the form where the user will input or edit data.
* **`searchCode`**: A string variable to store the `ItemCode` entered by the user in `Sheet1` (cell `B2`).
* **`foundCell`**: A range variable that will store the cell where the `ItemCode` is found in **Sheet2**.
* **`itemName`**: A string variable to store the `ItemName` entered or fetched from the form.
* **`price`**: A variant variable to store the `Price` entered or fetched from the form.
* **`currentUser`**: A string variable that stores the username of the person currently running the macro.

---

### 2. **Track Current User**

```vba
currentUser = Application.UserName
```

* **`Application.UserName`**: This retrieves the current user's name from Excel. It will be used to track which user is interacting with the form.

---

### 3. **Set References to Worksheets**

```vba
Set wsData = ThisWorkbook.Sheets("Sheet2")
Set wsForm = ThisWorkbook.Sheets("Sheet1")
```

* **`wsData`** is assigned to `Sheet2` (where the data is).
* **`wsForm`** is assigned to `Sheet1` (where the form is).

---

### 4. **Prevent Editing by Multiple Users**

```vba
If wsForm.Range("B1").Value <> "" And wsForm.Range("B1").Value <> currentUser Then
    MsgBox "Form is currently being edited by another user. Please wait.", vbExclamation
    Exit Sub
End If
```

* The cell `B1` on **Sheet1** stores the username of the user currently editing the form.
* If `B1` is not empty and the value in `B1` is not the current user's name, the macro will display a message saying the form is already in use and **exit** the procedure without doing anything further.

---

### 5. **Fetch the ItemCode**

```vba
searchCode = wsForm.Range("B2").Value
```

* The `ItemCode` entered by the user is stored in cell `B2` of **Sheet1**, and this value is assigned to the `searchCode` variable.

---

### 6. **Check for Empty `ItemCode`**

```vba
If searchCode = "" Then
    MsgBox "Please enter an Item Code in B2.", vbExclamation
    Exit Sub
End If
```

* If the `ItemCode` is empty (`""`), a message box will prompt the user to enter a valid `ItemCode` and then the macro will **exit**.

---

### 7. **Validate that `ItemCode` is Numeric**

```vba
If Not IsNumeric(searchCode) Then
    MsgBox "Item Code must be a valid number.", vbExclamation
    Exit Sub
End If
```

* The macro checks if the `ItemCode` is numeric using the `IsNumeric` function. If it's not a number, a message box will appear to notify the user, and the macro will **exit**.

---

### 8. **Search for the `ItemCode` in `Sheet2`**

```vba
Set foundCell = wsData.Range("A:A").Find(What:=searchCode, LookIn:=xlValues, LookAt:=xlWhole)
```

* The **`Find`** method searches for the entered `ItemCode` (`searchCode`) in **column A** of `Sheet2`.
* `What:=searchCode`: Specifies the value to search for.
* `LookIn:=xlValues`: Searches in the cell values (not formulas).
* `LookAt:=xlWhole`: Ensures the entire content of the cell matches the search term (not just part of it).
* **`foundCell`** will store the reference to the cell where the `ItemCode` is found.

---

### 9. **If `ItemCode` is Not Found**

```vba
If foundCell Is Nothing Then
    MsgBox "Item Code not found!", vbCritical
```

* If the `Find` method does not find the `ItemCode` (i.e., `foundCell` is `Nothing`), a message box will appear, notifying the user that the `ItemCode` was not found.

---

### 10. **Load the Data into the Form (If Item Found)**

```vba
Else
    If wsForm.Range("B3").Value = "" And wsForm.Range("B4").Value = "" Then
        wsForm.Range("B1").Value = currentUser ' Lock form with current user
        wsForm.Range("B3").Value = foundCell.Offset(0, 1).Value ' Load ItemName
        wsForm.Range("B4").Value = foundCell.Offset(0, 2).Value ' Load Price
        MsgBox "Record loaded. Edit and click the button again to save changes.", vbInformation
```

* If the `ItemCode` is found:

  * **`wsForm.Range("B1").Value = currentUser`**: Locks the form by assigning the current user’s name to cell `B1`.
  * The **`ItemName`** (from the cell next to the `ItemCode`) is loaded into `B3`.
  * The **`Price`** (from the cell two columns over from the `ItemCode`) is loaded into `B4`.
  * A message is displayed telling the user that the record has been loaded and asking them to edit the data and click the button again to save changes.

---

### 11. **Validate and Save Changes to Sheet2**

```vba
Else
    itemName = wsForm.Range("B3").Value
    price = wsForm.Range("B4").Value
    
    If itemName = "" Then
        MsgBox "Item Name cannot be empty.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(price) Then
        MsgBox "Price must be a valid number.", vbExclamation
        Exit Sub
    End If
    
    ' Save changes back to Sheet2
    foundCell.Offset(0, 1).Value = itemName
    foundCell.Offset(0, 2).Value = price
    
    MsgBox "Changes saved to database!", vbInformation
```

* If the user has edited the form, the macro will:

  * Retrieve the updated `ItemName` from `B3` and `Price` from `B4`.
  * Ensure `ItemName` is not empty and `Price` is a valid number.
  * If valid, the updated values are saved back into **Sheet2** using the `Offset` property to place them in the correct columns next to the `ItemCode`.
  * A message box confirms that the changes have been saved.

---

### 12. **Clear the Form and Unlock It**

```vba
wsForm.Range("B2:B4").ClearContents
wsForm.Range("B1").ClearContents
```

* After saving the changes, the **form fields** (`B2:B4`) are cleared so the user can enter new data if needed.
* The **form lock** (`B1`) is also cleared to allow the form to be used by others.

---

### 🏁 Summary:

* The macro starts by ensuring the user is allowed to interact with the form.
* It validates the input (`ItemCode`) and retrieves the associated data from `Sheet2`.
* The user can edit the form, and the changes are saved back into `Sheet2`.
* The form is locked while it's being used and cleared after the action is completed.

This entire process ensures smooth, secure, and validated editing of the data without conflict between users.


