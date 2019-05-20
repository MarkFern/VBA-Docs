---
title: Writing assignment statements (VBA)
keywords: vbcn6.chm1076692
f1_keywords:
- vbcn6.chm1076692
ms.prod: office
ms.assetid: 7699bec2-c5a2-6f35-3ec0-8aa7cefa622d
ms.date: 12/26/2018
localization_priority: Normal
---


# Writing assignment statements

Assignment statements assign a value (of a literal, [constant](../../Glossary/vbe-glossary.md#constant), [variable](../../glossary/vbe-glossary.md#variable), readable [property](../../glossary/vbe-glossary.md#property) or [expression](../../Glossary/vbe-glossary.md#expression) evaluation) to a variable, writable property or constant. Assignment statements always include an equal sign (**=**).

The following example assigns the return value of the **InputBox** function to the variable.

```vb
Sub Question() 
 Dim yourName As String 
 yourName = InputBox("What is your name?") 
 MsgBox "Your name is " & yourName 
End Sub
```


The **Let** statement is optional and is usually omitted. For example, the preceding assignment statement can be written as follows:

```vb
Let yourName = InputBox("What is your name?") 

```

The **[Set](../../reference/user-interface-help/set-statement.md)** statement is used to assign an object to a variable or writable property that has been declared as an object or variant. The **Set** keyword is required. In the following example, the **Set** statement assigns a range on `Sheet1` first to the [object variable](../../glossary/vbe-glossary.md#object-variable) `myCell`, and then secondly to the variant variable `myVariantVariable`. Finally, the example shows an assignment of a value that is again a computation of `Worksheets("Sheet1").Range("A1")` however, because **Set** is not used, the value assigned is not the range object (the value assigned is probably the value returned by the default member of the range object).

```vb
Sub SetExample() 
 Dim myCell As Range 
 Set myCell = Worksheets("Sheet1").Range("A1")            ' Object assignment.
 With myCell.Font 
 .Bold = True 
 .Italic = True 
 End With
 Dim myVariantVariable as Variant
 Set myVariantVariable = Worksheets("Sheet1").Range("A1") ' Object assignment.

' The following line doesn't assign the range object!
 myVariantVariable = Worksheets("Sheet1").Range("A1")     ' Assignment of non-object value.
End Sub
```

The following example sets the **Bold** property of the **Font** object for the active cell.

```vb
ActiveCell.Font.Bold = True 

```

If the type of the left-hand side of an assignment statement is not the same type as the right-hand side, an implicit type conversion or cast may occur on the right-hand side in order that the assignment can be executed. For more information on this, see the [Assignment statements _(implicit conversions & casts)_](../../Reference/User-Interface-Help/data-type-summary.md#assignment-statements-implicit-conversions--casts) section on the Data type summary page.

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
