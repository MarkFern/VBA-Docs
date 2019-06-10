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

[Assigns](../../Concepts/Getting-Started/writing-assignment-statements.md) or [coerces](../../Reference/User-Interface-Help/data-type-summary.md#assignment-statements-implicit-conversions--casts) a value to a [variable](../../Glossary/vbe-glossary.md#variable) or writable [property](../../Glossary/vbe-glossary.md#property).

## Syntax

[ **Let** ] _varname_ **=** _value_

<br/>

The **Let** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Let**|Optional. Explicit use of the **Let** [keyword](../../Glossary/vbe-glossary.md#keyword) is a matter of style, but it is usually omitted.|
| _varname_|Required. Name of the variable, a property [expression](../../glossary/vbe-glossary.md#expression) evaluating to the writable property, or an [object variable](../../glossary/vbe-glossary.md#object-variable) holding an [object](../../glossary/vbe-glossary.md#object) that has the writable property as its default member; variable data type cannot be an object-based type or an [array](../../glossary/vbe-glossary.md#array) data type, when value is to be assigned or coerced to a **variable**; follows standard naming conventions.|
| _value_|Required. Literal, variable, readble property, [constant](../../glossary/vbe-glossary.md#constant), or expression, that evaluates to a value directly assigned or coerced to the variable or writable property.|

## Remarks

The **Let** statement will only be successful if either:
- _value_ has the same [data type](../../Glossary/vbe-glossary.md#data-type) as _varname_, _or_
- there exists a **Let**-coercion rule to coerce _value_ to the data type of _varname_ (click [here](../../Reference/User-Interface-Help/data-type-summary.md#assignment-statements-implicit-conversions--casts) to access a section documenting all of the **Let**-coercion rules).

Information loss can occur in **Let**-coercions.

If execution of the statement fails, a [run-time error](../../glossary/vbe-glossary.md#run-time-error) will be raised. A [compile-time](../../glossary/vbe-glossary.md#compile-time) error may occur if the compiler perceives the **Let** statement as being invalid.

[Variant](../../Glossary/vbe-glossary.md#variant-data-type) variables can be assigned to either [string](../../glossary/vbe-glossary.md#string-expression) or [numeric expressions](../../glossary/vbe-glossary.md#numeric-expression). However, the reverse is not always true. Many non-[object](../../glossary/vbe-glossary.md#object) **Variant** values can be assigned to a string variable, whereas whether a **Variant** value can be assigned to a numeric variable depends on things such as whether the value can be interpreted as a number, and whether the value is within the variable's range. The [**IsNumeric**](../../reference/user-interface-help/isnumeric-function.md) function can help to determine whether a **Variant** can be converted to a number.

**Let** statements can be used to assign one record variable to another only when both variables are of the same [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type). Use the **[LSet](lset-statement.md)** statement to assign record variables of different user-defined types. Use the **[Set](set-statement.md)** statement to assign object references to variables.

## Example

This example assigns the values of literals to variables by using the explicit **Let** statement.

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
