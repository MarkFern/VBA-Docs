---
title: Variant data type
keywords: vblr6.chm1009056
f1_keywords:
- vblr6.chm1009056
ms.prod: office
ms.assetid: 19750b07-c2bf-dff7-67a1-91b06338cbc6
ms.date: 11/19/2018
localization_priority: Priority
---


# Variant data type

The **Variant** data type is the [data type](../../Glossary/vbe-glossary.md#data-type) for all [variables](../../Glossary/vbe-glossary.md#variable) that are not explicitly declared as some other type (using [statements](../../Glossary/vbe-glossary.md#statement) such as **Dim**, **Private**, **Public**, or **Static**). It has no [type-declaration character](../../Glossary/vbe-glossary.md#type-declaration-character).

The data type is a special data type that can contain any kind of data except fixed-length [String](../../Glossary/vbe-glossary.md#string-data-type) data, and data of [**user-defined types**](../../Glossary/vbe-glossary.md#user-defined-type) declared in the VBE using VBA's [**Type**](../../reference/user-interface-help/type-statement.md) statement. The **Variant** type supports **user-defined types** accessed through [VBE library references](../../how-to/set-reference-to-a-type-library.md). A **Variant** can also contain the special values [**Empty**](../../Glossary/vbe-glossary.md#empty), [**Nothing**](../../reference/user-interface-help/nothing-keyword.md), and [**Null**](../../Glossary/vbe-glossary.md#null). It can also contain values of the special [**Error** sub-type](../../reference/user-interface-help/cverr-function.md), as well as the **Variant** special value corresponding to the argument that the [**IsMissing**](../Reference/User-Interface-Help/ismissing-function.md) function returns **True** for. Currently, **Error** values are also considered to be **Variant** special values, and can only be used when 'wrapped' as **Variant** data.

You can determine how the data in a **Variant** is treated by using the [**VarType** function](../../reference/user-interface-help/vartype-function.md) in conjunction with the [**IsObject** function](../../reference/user-interface-help/isobject-function.md). The [**TypeName** function](../../reference/user-interface-help/typename-function.md), when used together with **IsObject** & **VarType**, can increase clarity of determination for user-defined types & object types (by being able to get the name of the specific type used, when these types are used). See [example](#determining-how-the-data-in-a-variant-is-treated) below showing how you can use these functions to obtain such information.

Numeric data can be any integer or real number value ranging from -1.797693134862315E308 to -4.94066E-324 for negative values and from 4.94066E-324 to 1.797693134862315E308 for positive values. 

Generally, numeric **Variant** data is maintained in its original data type within the **Variant**. For example, if you assign an [Integer](../../Glossary/vbe-glossary.md#integer-data-type) to a **Variant**, subsequent operations treat the **Variant** as an **Integer**. However, if an arithmetic operation is performed on a **Variant** containing a [Byte](../../Glossary/vbe-glossary.md#byte-data-type), an **Integer**, a [Long](../../Glossary/vbe-glossary.md#long-data-type), or a [Single](../../Glossary/vbe-glossary.md#single-data-type), and the result exceeds the normal range for the original data type, the result is promoted within the **Variant** to the next larger data type. A **Byte** is promoted to an **Integer**, an **Integer** is promoted to a **Long**, and a **Long** and a **Single** are promoted to a [Double](../../Glossary/vbe-glossary.md#double-data-type). 

An error occurs when **Variant** variables containing [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type), and **Double** values exceed their respective ranges.

You can use the **Variant** data type in place of almost any data type to work with data in a more flexible way. If the contents of a **Variant** variable are digits, they may be either the string representation of the digits or their actual value, depending on the context. For example:

```vb
Dim MyVar As Variant 
MyVar = 98052 

```

In the preceding example, `MyVar` contains a numeric representation&mdash;the actual value `98052`. Arithmetic operators work as expected on **Variant** variables that contain numeric values or string data that can be interpreted as numbers. If you use the **+** operator to add `MyVar` to another **Variant** containing a number or to a variable of a [numeric type](../../Glossary/vbe-glossary.md#numeric-type), the result is an arithmetic sum.

The value [Empty](../../Glossary/vbe-glossary.md#empty) denotes a **Variant** variable that hasn't been initialized (assigned an initial value). A **Variant** containing **Empty** is 0 if it is used in a numeric context, and a zero-length string ("") if it is used in a string context.

Don't confuse **Empty** with [Null](../../Glossary/vbe-glossary.md#null). **Null** indicates that the **Variant** variable intentionally contains no valid data.

In a **Variant**, an **Error** value is a special value used to indicate that an error condition has occurred in a [procedure](../../Glossary/vbe-glossary.md#procedure). However, unlike for other kinds of errors, normal application-level error handling does not occur. This allows you, or the application itself, to take some alternative action based on the **Error** value. **Error** values are created by converting positive integers to **Error** values by using the **[CVErr](cverr-function.md)** function.

If the **IsMissing** function returns **True** for a particular argument, it means that the argument holds the **Variant** special value that represents a missing procedure argument.

 > [!NOTE] 
	> The **Variant** data type cannot itself be used as a sub-type of the **Variant** data type (even though the COM [VARIANT type](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f) {that VBA's **Variant** type maps to} allows such things for its type). However, **Variant** data can be assigned to **Variant** variables.

## Examples

### Determining how the data in a Variant is treated

```vb
Function VariantInformation(argument) As Variant()
 Dim ReturnValue(1 To 2) As Variant
 If IsObject(argument) Then
  ' IsObject only returns true for vbObject types.
  ' For vbDataObject types, varType will return right value.
  ReturnValue(1) = vbObject
 Else
  ReturnValue(1) = varType(argument)
 End If
        
 ' TypeName is only used when it helps.
 Select Case ReturnValue(1)
 Case vbUserDefinedType, _
      vbArray + vbUserDefinedType, _
      vbArray + vbObject, _
      vbArray + vbDataObject
  ReturnValue(2) = TypeName(argument)
 Case vbObject, vbDataObject
  If Not (argument Is Nothing) Then
   ReturnValue(2) = TypeName(argument)
  End If
  ' Better to not run TypeName if no object, as the string
  ' 'Nothing' could be the name of a class.
 End Select
    
 VariantInformation = ReturnValue
End Function
```

## See also

- [Data type summary](data-type-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
