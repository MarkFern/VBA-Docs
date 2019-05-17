---
title: Let statement (VBA)
keywords: vblr6.chm1008960
f1_keywords:
- vblr6.chm1008960
ms.prod: office
ms.assetid: da1ec875-3c6a-b66d-a85f-bbf33f9a307a
ms.date: 12/03/2018
localization_priority: Normal
---


# Let statement

Assigns or coerces a value to a [variable](../../Glossary/vbe-glossary.md#variable) or [property](../../Glossary/vbe-glossary.md#property).

## Syntax

[ **Let** ] _varname_ **=** _value_

<br/>

The **Let** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Let**|Optional. Explicit use of the **Let** [keyword](../../Glossary/vbe-glossary.md#keyword) is a matter of style, but it is usually omitted.|
| _varname_|Required. Name of the variable or property; follows standard variable naming conventions.|
| _value_|Required. Literal, variable, [constant](../../glossary/vbe-glossary.md#constant), or [expression](../../glossary/vbe-glossary.md#expression), that evaluates to a value directly assigned or coerced to the variable or property.|

## Remarks

The **Let** statement will only be successful if either:
- _value_ has the same [data type](../../Glossary/vbe-glossary.md#data-type) as _varname_, _or_
- there exists a **Let**-coercion rule to coerce _value_ to the data type of _varname_ (**Let**-coercion is the same thing as [implicit type conversion in the context of the value assignment of an assignment statement](../../Reference/User-Interface-Help/data-type-summary.md#assignment-statements-implicit-conversions--casts)).

Information loss can occur in **Let**-coercions.

If execution of the statement fails, a [run-time error](../../glossary/vbe-glossary.md#run-time-error) will be raised. A [compile-time](../../glossary/vbe-glossary.md#compile-time) error may occur if the compiler perceives the **Let** statement as being invalid.

[Variant](../../Glossary/vbe-glossary.md#variant-data-type) variables can be assigned to either string or numeric expressions. However, the reverse is not always true. Many non-object **Variant** values can be assigned to a string variable, whereas whether a **Variant** value can be assigned to a numeric variable depends on things such as whether the value can be interpreted as a number, and whether the value is within the variable's range. The **IsNumeric** function can help to determine whether a **Variant** can be converted to a number.

**Let** statements can be used to assign one record variable to another only when both variables are of the same [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type). Use the **[LSet](lset-statement.md)** statement to assign record variables of different user-defined types. Use the **[Set](set-statement.md)** statement to assign object references to variables.

## Example

This example assigns the values of expressions to variables by using the explicit **Let** statement.

```vb
Dim MyStr, MyInt 
' The following variable assignments use the Let statement. 
Let MyStr = "Hello World" 
Let MyInt = 5 

```

<br/>

The following are the same assignments without the **Let** statement.

```vb
Dim MyStr, MyInt 
MyStr = "Hello World" 
MyInt = 5 

```


## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
