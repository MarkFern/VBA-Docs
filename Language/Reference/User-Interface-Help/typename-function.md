---
title: TypeName function (Visual Basic for Applications)
keywords: vblr6.chm1010100
f1_keywords:
- vblr6.chm1010100
ms.prod: office
ms.assetid: 9353f1d5-5b64-9cad-5cc3-e1487bdd3afd
ms.date: 12/13/2018
localization_priority: Normal
---


# TypeName function

Returns a **String** that provides information about a [variable](../../Glossary/vbe-glossary.md#variable).

## Syntax

**TypeName**(_varname_) 

The required _varname_ [argument](../../Glossary/vbe-glossary.md#argument) is any [Variant](../../Glossary/vbe-glossary.md#variant-data-type).

## Remarks

The string returned by **TypeName** can be any one of the following:

|String returned|Variable|
|:-----|:-----|
|_objecttype_|An [object](../../glossary/vbe-glossary.md#object) whose [object type](../../Glossary/vbe-glossary.md#object-type) is _objecttype_|
|_udtype_|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) where the type has name _udtype_|
|"Byte"|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) value|
|"Integer"|Integer of data-type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|
|"Long"|[Long](../../Glossary/vbe-glossary.md#long-data-type) integer|
|"Single"|[Single](../../Glossary/vbe-glossary.md#single-data-type) value (single-precision floating-point number)|
|"Double"|[Double](../../Glossary/vbe-glossary.md#double-data-type) value (double-precision floating-point number)|
|"Currency"|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|
|"Decimal"|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|
|"Date"|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|
|"String"|[String](../../Glossary/vbe-glossary.md#string-data-type)|
|"Boolean"|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|
|"Error"|An [**Error**](../../reference/user-interface-help/cverr-function.md) value|
|"Empty"|[**Empty**](../../Glossary/vbe-glossary.md#empty) value (variable uninitialized) \*\*|
|"Null"|[**Null**](../../Glossary/vbe-glossary.md#null) value (no valid data) \*\*|
|"Unknown"|An [object](../../glossary/vbe-glossary.md#object) whose type is unknown|
|"Nothing"|[**Nothing**](nothing-keyword.md) value (object variable that doesn't refer to an object) \*|

<sup>* Special value. ** Variant special value.</sup>

If _varname_ is an [array](../../Glossary/vbe-glossary.md#array), the returned string is a string from the above table (indicating the array type) with empty parentheses appended to it. For example, if _varname_ is an array of integers, **TypeName** returns "Integer()".

## Example

This example uses the **TypeName** function to return information about a variable.

```vb    
' Declare & assign variables.
Dim MyType
Dim StrVar As String, IntVar As Integer, CurVar As Currency
Dim UninitVar
Dim NullVar: NullVar = Null  ' Assign Null value.
Dim ArrayVar(1 To 5) As Integer
Dim AppVar As Object: Set AppVar = Excel.Application
Dim NoObjVar As Object
' Declare user-defined-type variable.
Dim UDTVar As mscorlib.Guid  ' From .NET Framework library.
        
MyType = TypeName(StrVar)    ' Returns "String".
MyType = TypeName(IntVar)    ' Returns "Integer".
MyType = TypeName(CurVar)    ' Returns "Currency".
MyType = TypeName(UninitVar) ' Returns "Empty".
MyType = TypeName(NullVar)   ' Returns "Null".
MyType = TypeName(ArrayVar)  ' Returns "Integer()".
MyType = TypeName(AppVar)    ' Returns "Application".
MyType = TypeName(NoObjVar)  ' Returns "Nothing".
MyType = TypeName(UDTVar)    ' Returns "Guid".

```


## See also

- [VarType function](../user-interface-help/vartype-function.md)
- [Data types](data-type-summary.md)
- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
