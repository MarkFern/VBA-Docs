---
title: Data type summary
keywords: vblr6.chm1008885
f1_keywords:
- vblr6.chm1008885
ms.prod: office
ms.assetid: 24723bdf-8454-f661-7914-d731e74d2e7b
ms.date: 11/19/2018 
localization_priority: Priority
---


# Data type summary

A data type is the characteristic of a variable that determines what kind of data it can hold.

## Non-intrinsic data types

Non-intrinsic data types include those in the following table as well as all other specific types of object.
 
|Non-intrinsic data type|Storage size|Range|
|:--------|:-----------|:----|
|**[User-defined](../../How-to/user-defined-data-type.md)** <BR>_<sup>(using **Type** or other means)</sup>_ |Number required by elements|The range of each element is the same as the range of its data type.|
|[**Collection** object](../../reference/user-interface-help/collection-object.md)|Unknown|Unknown|
|[**Dictionary** object](../../reference/user-interface-help/dictionary-object.md)|Unknown|Unknown|

## Intrinsic data types

The following table shows the supported intrinsic [data types](../../Glossary/vbe-glossary.md#data-type), including storage sizes and ranges.

|Intrinsic data type|Storage size|Range|
|:--------|:-----------|:----|
|**[Boolean](boolean-data-type.md)**|2 bytes|**True** or **False**|
|**[Byte](byte-data-type.md)**|1 byte|0 to 255|
|**[Currency](currency-data-type.md)** _(scaled integer)_|8 bytes|-922,337,203,685,477.5808 to 922,337,203,685,477.5807|
|**[Date](date-data-type.md)**|8 bytes|January 1, 100, to December 31, 9999|
|**[Decimal](decimal-data-type.md)**|14 bytes|+/-79,228,162,514,264,337,593,543,950,335 with no decimal point<br/><br/>+/-7.9228162514264337593543950335 with 28 places to the right of the decimal<br/><br/>Smallest non-zero number is+/-0.0000000000000000000000000001|
|**[Double](double-data-type.md)** <BR>_(double-precision floating-point)_|8 bytes|-1.79769313486231E308 to -4.94065645841247E-324 for negative values<br/><br/>4.94065645841247E-324 to 1.79769313486232E308 for positive values|
|**[Integer](integer-data-type.md)**|2 bytes|-32,768 to 32,767|
|**[Long](long-data-type.md)** _(Long integer)_|4 bytes|-2,147,483,648 to 2,147,483,647|
|**[LongLong](longlong-data-type.md)** _(LongLong integer)_|8 bytes|-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807<br/><br/>Valid on 64-bit platforms only.|
|**[LongPtr](longptr-data-type.md)** <BR>_(Long integer on 32-bit systems, LongLong integer on 64-bit systems)_|4 bytes on 32-bit systems<br/><br/>8 bytes on 64-bit systems|-2,147,483,648 to 2,147,483,647 on 32-bit systems<br/><br/>-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems|
|**[Object](object-data-type.md)**|4 bytes|Any **Object** reference|
|**[Single](single-data-type.md)** <BR>_(single-precision floating-point)_|4 bytes|-3.402823E38 to -1.401298E-45 for negative values<br/><br/>1.401298E-45 to 3.402823E38 for positive values|
|**[String](string-data-type.md)** _(variable-length)_|10 bytes + string length|0 to approximately 2 billion|
|**String** _(fixed-length)_|Length of string|1 to approximately 65,400|
|**[Variant](variant-data-type.md)** _(with numbers)_|16 bytes|Any numeric value up to the range of a **Double**|
|**Variant** _(with characters)_|22 bytes + string length (24 bytes on 64-bit systems)|Same range as for variable-length **String**|
|**Variant** _(with objects)_|Unknown|Same range as **Object**|
|**Variant** _(with [user-defined type](../../How-to/user-defined-data-type.md))_|Unknown|Only data of a **user-defined type** accessed through a [VBE library reference](../../reference/user-interface-help/references-dialog-box.md)|
|**Variant** <BR>_(with special values [**Empty**](../../Glossary/vbe-glossary.md#empty) or [**Null**](../../Glossary/vbe-glossary.md#null))_|Unknown|Unknown|
|**Variant** _(with [**Error** sub-type](../../reference/user-interface-help/cverr-function.md))_|Unknown|Unknown|


<br/>

A **Variant** containing an array requires 12 bytes more than the array alone.

> [!NOTE] 
> [Arrays](../../Glossary/vbe-glossary.md#array) of any data type require 20 bytes of memory plus 4 bytes for each array dimension plus the number of bytes occupied by the data itself. The memory occupied by the data can be calculated by multiplying the number of data elements by the size of each element.
> 
> For example, the data in a single-dimension array consisting of 4 **Integer** data elements of 2 bytes each occupies 8 bytes. The 8 bytes required for the data plus the 24 bytes of overhead brings the total memory requirement for the array to 32 bytes. On 64-bit platforms, SAFEARRAY's take up 24-bits (plus 4 bytes per Dim statement). The pvData member is an 8-byte pointer and it must be aligned on 8 byte boundaries.

> [!NOTE] 
> [LongPtr](longptr-data-type.md) is not a true data type because it transforms to a [Long](long-data-type.md) in 32-bit environments, or a [LongLong](longlong-data-type.md) in 64-bit environments. **LongPtr** should be used to represent pointer and handle values in [Declare statements](declare-statement.md) and enables writing portable code that can run in both 32-bit and 64-bit environments.

> [!NOTE] 
> Use the **StrConv** function to convert one type of string data to another.


## Conversion between data types

### Implicit conversions

#### Assignment statements

The following table summarizes several implicit type conversions that always take place in variable [assignment statements](../../Concepts/Getting-Started/writing-assignment-statements.md) of values, & that happen so that the assignments still assign values in some manner of speaking.

|Variable type|Value form|
|:--------|:-----------|
|**Variant**|Type same as valid **Variant** sub-type|
|The **Object** type|Valid object type|
|An [interface](../../Glossary/vbe-glossary.md#interface) type|Object type defined using the [**Implements**](../../reference/user-interface-help/implements-statement.md) statement to specify implementation of the interface|
|An object type|Object type defined using the **Implements** statement to specify implementation of the interface derived from the variable type|
|Integral type apart from the **Boolean** type|Intrinsic numerical type; within the range of the variable type|
|**Boolean** type|Intrinsic numerical type|
|Integral, **Single**, **Double** or **Boolean** type|**String** textual representation of a number that parses as a value of an intrinsic numerical type, and that if put into an appropriate numerical type, would be automatically implicitly coerced to the variable type in a related variable assignment statement|
|**Date** type|**String** textual representation of a valid date, that parses as a date.|
|**Date** type|**String** textual representation of a number that parses as a value of an intrinsic numerical type, and that if put into an appropriate numerical type, would be automatically implicitly coerced to the date type in a related **Date**-variable assignment statement|
|**Currency** type|**String** textual representation of a valid currency amount, that parses as a currency amount.|
|**Currency** type|**String** textual representation of a rational number that parses as such a number; within range of **Currency** type|
|**String**|Any intrinsic data type|
|Intrinsic data type|For each variable type, the form is that required by the union of all other rules in this table that apply to the particular variable type, except that the form applies to sub-type data of a **Variant** value where the **Variant** value is the actual value assigned|


#### Procedure invocations

The following table summarizes known implicit type conversions that always take place in procedure invocations, for particular start & end types. The start type is the type of a value passed as an argument in a standard invocation of a procedure. The end type is the final type of that same argument when it is accessed by the contents of the procedure. Note that it seems unlikely that other implicit type conversions exist in regard to procedure invocations.

|Argument type defined in procedure|Argument type at moment of standard invocation of procedure|
|:--------|:-----------|
|**Variant**|Valid **Variant** sub-type|
|The **Object** type|Valid object type|
|An interface type|Object type defined using the **Implements** statement to specify implementation of the interface|
|An object type|Object type defined using the **Implements** statement to specify implementation of the interface derived from the argument type defined in the procedure definition|


### Explicit conversions

See [Type conversion functions](../../concepts/getting-started/type-conversion-functions.md) for examples of how to use the following functions to coerce an expression to a specific data type: **CBool**, **CByte**, **CCur**, **CDate**, **CDbl**, **CDec**, **CInt**, **CLng**, **CLngLng**, **CLngPtr**, **CSng**, **[CStr](#returns-for-cstr)**, and **CVar**.

The [**Fix**, and **Int** functions](int-fix-functions.md) provide other forms of integeric conversion.

**[CVErr](cverr-function.md)** can be used to create values of the **Variant** sub-type **Error** from an error number.

> [!NOTE] 
> **CLngLng** is valid on 64-bit platforms only.

#### Returns for CStr

|If _expression_ is|CStr returns|
|:-----------------|:-----------|
|**Boolean**|A string containing **True** or **False**.|
|**Date**|A string containing a date in the short date format of your system.|
|[Empty](../../Glossary/vbe-glossary.md#empty)|A zero-length string ("").|
|**Error**|A string containing the word **Error** followed by the [error number](../../Glossary/vbe-glossary.md#error-number).|
|[Null](../../Glossary/vbe-glossary.md#null)|A [run-time error](../../Glossary/vbe-glossary.md#run-time-error).|
|Other numeric|A string containing the number.|

## Verify data types

To verify data types, see the following functions: 

- [IsArray](isarray-function.md)
- [IsDate](isdate-function.md)
- [IsEmpty](isempty-function.md)
- [IsError](iserror-function.md)
- [IsMissing](ismissing-function.md)
- [IsNull](isnull-function.md)
- [IsNumeric](isnumeric-function.md)
- [IsObject](isobject-function.md)
- [VarType](vartype-function.md)
- [TypeName](typename-function)

## See also

- [VarType constants](../../concepts/getting-started/vartype-constants.md)
- [Keywords by task](keywords-by-task.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
