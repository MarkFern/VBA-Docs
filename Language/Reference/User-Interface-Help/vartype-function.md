---
title: VarType function (Visual Basic for Applications)
keywords: vblr6.chm1009057
f1_keywords:
- vblr6.chm1009057
ms.prod: office
ms.assetid: 7422fba5-7ea9-1d91-fc0e-5694c352d2d0
ms.date: 04/17/2019
localization_priority: Normal
---


# VarType function

Returns an **Integer** where the returned value will indicate one of the following things, the choice of which depending upon the parameter passed:
1) The subtype of a [**Variant**](../../Glossary/vbe-glossary.md#variant-data-type) [variable](../../Glossary/vbe-glossary.md#variable)/[expression](../../glossary/vbe-glossary.md#expression).
2) The type or lack of type for an object's default member's return value where the object is one returned by a **Variant** expression or is the value of a **Variant** variable.
3) The **Variant** special value that a **Variant** expression or variable evaluates to.

## Syntax

**VarType**(_varname_)

The required _varname_ [argument](../../Glossary/vbe-glossary.md#argument) is either a **Variant** variable/expression, or an argument that is automatically coerced to a **Variant** value.
 
## Return values

Return value is either:

- just one of the following constants excluding the `vbArray` constant & the `vbVariant` constant, _or_
- the `vbArray` constant added to any of the other constants from the following list.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbEmpty**|0|[**Empty**](../../Glossary/vbe-glossary.md#empty) value (variable uninitialized) \*|
|**vbNull**|1|[**Null**](../../Glossary/vbe-glossary.md#null) value (no valid data) \*|
|**vbInteger**|2|Integer of data type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|
|**vbLong**|3|[Long](../../Glossary/vbe-glossary.md#long-data-type) integer|
|**vbSingle**|4|[Single](../../Glossary/vbe-glossary.md#single-data-type) value (single-precision floating-point number)|
|**vbDouble**|5|[Double](../../Glossary/vbe-glossary.md#double-data-type) value (double-precision floating-point number)|
|**vbCurrency**|6|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|
|**vbDate**|7|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|
|**vbString**|8|[String](../../Glossary/vbe-glossary.md#string-data-type)|
|**vbObject**|9|[Object](../../glossary/vbe-glossary.md#object)|
|**vbError**|10|An [**Error**](../../reference/user-interface-help/cverr-function.md) value|
|**vbBoolean**|11|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|
|**vbVariant**|12|**Variant** (used for return value only when added to **vbArray** constant)|
|**vbDataObject**|13|Non-ActiveX [Automation object](../../glossary/vbe-glossary.md#automation-object-1)|
|**vbDecimal**|14|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|
|**vbByte**|17|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) value|
|**vbLongLong**|20|[LongLong](longlong-data-type.md) integer (valid on 64-bit platforms only)|
|**vbUserDefinedType**|36|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type)|
|**vbArray**|8192|[Array](../../Glossary/vbe-glossary.md#array) (always added to another constant when returned by this function)|

<sup>* **Variant** special value.</sup>

> [!NOTE] 
> These [constants](../../Glossary/vbe-glossary.md#constant) are specified by Visual Basic for Applications. The names can be used anywhere in your code in place of the actual values.

## Remarks

If a standard object (one that supports ActiveX) is passed, and has a parameterless default member (either a default property or default function), **VarType**(_object_) returns a value indicating the type of the default member's return value in the case that there is a return value, and the value of the **vbEmpty** constant when there is no return value. If an object is passed that doesn't fulfill this criteria, the constant **vbObject** or the constant **vbDataObject** is returned, the constant representing the object type.

The **VarType** function never returns the value for **vbArray** by itself. It is always added to some other value to indicate an array of a particular type. For example, the value returned for an array of integers is calculated as **vbInteger** + **vbArray**, or 8194. 

The constant **vbVariant** is only returned in conjunction with **vbArray** to indicate that the argument to the **VarType** function is an array of type **Variant**.

When the function's argument evaluates to a **Variant** special value, the constant associated with the special value is returned.

## Example

This example uses the **VarType** function to determine the subtypes of different variables, the type of an object's default member's return value, & the **Variant** special values that certain variables hold.

```vb
Dim MyCheck
Dim IntVar, StrVar, DateVar, AppVar, WorkbookVar
' `stdole` is a library reference to the OLE Automation library.
' The IUnknown interface is the most basic COM interface
' that was used before ActiveX started to be used.
Dim NonActiveXObjectVar As stdole.IUnknown
Dim ArrayVar
Dim UninitVar
Dim NullVar: NullVar = Null            ' Assign Null value.
IntVar = 459: StrVar = "Hello World": DateVar = #2/12/1969#
Set AppVar = Excel.Application
Set WorkbookVar = ActiveWorkbook       ' Workbook object.
ArrayVar = Array("1st Element", "2nd Element")

' Run VarType function on different types.
MyCheck = VarType(IntVar)              ' Returns 2.
MyCheck = VarType(DateVar)             ' Returns 7.
MyCheck = VarType(StrVar)              ' Returns 8.
' Assuming 'Microsoft Excel 16.0 Object Library' reference is 
' being used, return values for AppVar & WorkbookVar are as 
' follows.
MyCheck = VarType(AppVar)              ' Returns 8 (vbString)
                                       ' even though AppVar is
                                       ' an object.
MyCheck = VarType(WorkbookVar)         ' Returns 9 (vbObject)
                                       ' because it's a standard
                                       ' object without a
                                       ' default member.

MyCheck = VarType(NonActiveXObjectVar) ' Returns 13 (vbDataObject).
MyCheck = VarType(ArrayVar)            ' Returns 8204 which is
                                       ' `8192 + 12`, the computation of
                                       ' `vbArray + vbVariant`.

' Run VarType function on Variant special values.
MyCheck = VarType(UninitVar) ' Returns 0 (vbEmpty).
MyCheck = VarType(NullVar)   ' Returns 1 (vbNull).
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
