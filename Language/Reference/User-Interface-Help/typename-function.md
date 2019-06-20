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

Returns a **String** that provides type & data-status information concerning the passed [argument](../../Glossary/vbe-glossary.md#argument).

## Syntax

**TypeName**(_arg_) 

The required _arg_ argument must be of the [**Variant**](../../Glossary/vbe-glossary.md#variant-data-type) type, or be able to be [coerced](../../Reference/User-Interface-Help/data-type-summary.md#implicit-conversions--casts) to it.


## Remarks

The rules for what string is returned by **TypeName**, are shown in the following table:

|Argument|String returned|
|:-----|:-----|
|An [object](../../glossary/vbe-glossary.md#object) whose [object type](../../Glossary/vbe-glossary.md#object-type) has been determined by this function as being _objecttype_|_objecttype_|
|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) where the type has name _udtype_|_udtype_|
|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) integer|"Byte"|
|Integer of data-type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|"Integer"|
|[Long](../../Glossary/vbe-glossary.md#long-data-type) integer|"Long"|
|[LongLong](../../reference/user-interface-help/longlong-data-type.md) integer|"LongLong"|
|[Single](../../Glossary/vbe-glossary.md#single-data-type) value<br><sup>_(single-precision floating-point number)_</sup>|"Single"|
|[Double](../../Glossary/vbe-glossary.md#double-data-type) value<br><sup>_(double-precision floating-point number)_</sup>|"Double"|
|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|"Currency"|
|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|"Decimal"|
|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|"Date"|
|[String](../../Glossary/vbe-glossary.md#string-data-type)|"String"|
|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|"Boolean"|
|Argument can have either of the following forms:<br><ol type="i"><li>An [**Error**](../../reference/user-interface-help/cverr-function.md) value.</li><li>The [parameter](../../glossary/vbe-glossary.md#parameter) for a [_missing_](../../reference/user-interface-help/ismissing-function.md) [_optional_](../../concepts/getting-started/understanding-named-arguments-and-optional-arguments.md) **Variant** argument of some procedure, that hasn't yet had a conventional value assignment (the "missing" flag bit will have been set), or a [variable](../../glossary/vbe-glossary.md#variable) or [property](../../glossary/vbe-glossary.md#property) holding the value of such a parameter. At the time of writing, such values are also **Error** values of the [error number 448](../../reference/user-interface-help/named-argument-not-found-error-448.md).</li></ul><sup>[\*\*](#doubleasteriskfootnote "Variant special value.")</sup>|"Error"|
|[**Empty**](../../Glossary/vbe-glossary.md#empty) value<br><sup>_(represents uninitialized variable)_</sup> <sup>[\*\*](#doubleasteriskfootnote "Variant special value.")</sup>|"Empty"|
|[**Null**](../../Glossary/vbe-glossary.md#null) value<br><sup>_(represents no valid data)_</sup> <sup>[\*\*](#doubleasteriskfootnote "Variant special value.")</sup>|"Null"|
|An [object](../../glossary/vbe-glossary.md#object) whose type name cannot be determined with this function <sup>[&dagger;](#daggerfootnote "Such objects include all objects that do not implement the GetTypeInfo function from COM's IDispatch interface.")</sup>|"Unknown"|
|[**Nothing**](nothing-keyword.md) value<br><sup>_(default value of an unassigned [object variable](../../glossary/vbe-glossary.md#object-variable))_</sup> <sup>[\*](#asteriskfootnote "Special value.")</sup>|"Nothing"|
|An [array](../../Glossary/vbe-glossary.md#array) of a non-object-based type|The string obtained by this function when it is passed data of the base type of the array, with empty parentheses appended to it. <sup>[&Dagger;](#doubledaggerfootnote "For example, if arg is an array of integers, TypeName returns \"Integer()\"")</sup>|
|An array of an object-based type|If this function returns "Unknown" when it is passed an object of the base type of the array, "Unknown()", otherwise "Object()".|

<a name="asteriskfootnote"><sup>*</sup></a> Special value. <a name="doubleasteriskfootnote"><sup>**</sup></a> **Variant** special value.<br>
<a name="daggerfootnote"><sup>&dagger;</sup></a> Such objects include all objects that do not implement the **GetTypeInfo** function from COM's **IDispatch** interface.<br>
<a name="doubledaggerfootnote"><sup>&Dagger;</sup></a> For example, if _arg_ is an array of integers, **TypeName** returns "Integer()". 


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
