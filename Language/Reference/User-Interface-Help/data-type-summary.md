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

A data type is the characteristic of a [variable](../../glossary/vbe-glossary.md#variable) that determines what kind of data it can hold.

## Non-intrinsic data types

Non-intrinsic data types include those in the following table as well as all other specific [object types](../../Glossary/vbe-glossary.md#object-type) (which includes all VBA compatible [interface](../../Glossary/vbe-glossary.md#interface) types through their representation using object types).
 
|Non&#8209;intrinsic&nbsp;data&nbsp;type|Storage size <sup>_(in bytes)_</sup>|Range|
|:--------|:-----------|:----|
|**[User-defined](../../How-to/user-defined-data-type.md)** <BR>_<sup>(using **Type** or other means)</sup>_ |Number required by elements|The range of each element is the same as the range of its data type.|
|[**Collection** object](../../reference/user-interface-help/collection-object.md)|<hr>|<hr>|
|[**Dictionary** object](../../reference/user-interface-help/dictionary-object.md)|<hr>|<hr>|

## Intrinsic data types

The following table shows the supported intrinsic [data types](../../Glossary/vbe-glossary.md#data-type), including storage sizes and ranges.

|Intrinsic data type|Storage size <sup>_(in bytes)_</sup>|Range|
|:--------|:-----------|:----|
|[Array](../../concepts/getting-started/using-arrays.md)|&nbsp;&nbsp;&nbsp;16 **+**<BR>&nbsp;&nbsp;&nbsp;_**LongPtr**&nbsp;storage&nbsp;byte&nbsp;size_ **+**<BR>&nbsp;&nbsp;&nbsp;(8 **&times;** _no.&nbsp;of&nbsp;array&nbsp;dimensions_) **+**<BR>**&Sigma;**&nbsp;(_each&nbsp;element's&nbsp;storage&nbsp;byte&nbsp;size_)<br><sup>[*](#asteriskfootnote "Read this footnote for an example showing calculation of the storage size for a particular array: the data in a single-dimension array consisting of 4 Integer data elements of 2 bytes each occupies 8 bytes; the 8 bytes required for the data plus the 24 bytes of overhead in 32-bit environments brings the total memory requirement for the array to 32 bytes in 32-bit environments.")</sup>|Each element must have the same data type. The data type in VBA, is chosen when executing the [variable's declaration](../../concepts/getting-started/declaring-variables.md). Once chosen, another cannot be chosen whilst the code is running. Each element has the same range as the chosen data type. The element configuration can have up to 60 dimensions, and has a maximum size limited by your operating system & amount of available RAM. The index range for each dimension is some contiguous set of integers or just one particular integer, specified in the element configuration.<BR><BR>Variable must be declared as either a fixed-size array or a dynamic (re-sizeable) array. During execution, if the variable is a fixed-size array, the variable cannot accommodate a dynamic array, & vice versa. Dynamic arrays have no initial element configuration, & their element configurations can be re-specified during run-time as many times as is needed. Fixed-size arrays on the other hand, have for their entire run-time life-times, immutable element configurations. The element configuration is programatically specified within VBA, through VBA variable declarations for fixed-size arrays, & through [**ReDim**](../../reference/user-interface-help/redim-statement.md) statements for dynamic arrays.|
|**[Boolean](boolean-data-type.md)**|2|**True** or **False**|
|**[Byte](byte-data-type.md)**|1|0 to 255|
|**[Currency](currency-data-type.md)** <sup>_(scaled integer)_</sup>|8|-922,337,203,685,477.5808 to 922,337,203,685,477.5807|
|**[Date](date-data-type.md)**|8|January 1, 100, to December 31, 9999|
|**[Decimal](decimal-data-type.md)**|14|+/-79,228,162,514,264,337,593,543,950,335 with no decimal point<br/><br/>+/-7.9228162514264337593543950335 with 28 places to the right of the decimal<br/><br/>Smallest non-zero number is+/-0.0000000000000000000000000001|
|**[Double](double-data-type.md)** <BR><sup>_(double-precision floating-point)_</sup>|8|-1.79769313486231E308 to -4.94065645841247E-324 for negative values<br/><br/>4.94065645841247E-324 to 1.79769313486232E308 for positive values|
|**[Integer](integer-data-type.md)**|2|-32,768 to 32,767|
|**[Long](long-data-type.md)** <sup>_(Long integer)_<sup>|4|-2,147,483,648 to 2,147,483,647|
|**[LongLong](longlong-data-type.md)** <sup>_(LongLong integer)_<sup>|8|-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807<br/><br/>Valid on 64-bit platforms only.|
|**[LongPtr](longptr-data-type.md)** <BR><sup>_(Long integer on 32-bit systems,<BR>LongLong integer on 64-bit systems)_<sup>|4 on 32-bit systems<br/><br/>8 on 64-bit systems|-2,147,483,648 to 2,147,483,647 on 32-bit systems<br/><br/>-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems|
|**[Object](object-data-type.md)**|4|Any **Object** reference|
|**[Single](single-data-type.md)** <BR><sup>_(single-precision floating-point)_</sup>|4|-3.402823E38 to -1.401298E-45 for negative values<br/><br/>1.401298E-45 to 3.402823E38 for positive values|
|**[String](string-data-type.md)**<BR>_(variable-length)_|4 + _**LongPtr**&nbsp;storage&nbsp;size_ + ((_string length_ + 1) &times; 2)|Each character can be any Unicode character; length of string can be changed to any non-negative integer up to approximately 2 billion.|
|**String**<BR>_(fixed-length)_|_Length of string_ &times; 2|Each character can be any Unicode character; length of string can be set to any positive integer up to approximately 65,400; string length can only be set during execution of the variable's declaration|
|**Variant**<BR>_(with arrays)_|12 more than the array-data-type version of array.|Any array-data-type array.|
|**Variant**<BR>_(with characters)_|22 + (_string length_ &times; 2)|Same range as for variable-length **String**|
|**[Variant](variant-data-type.md)**<BR>_(with numbers)_|16|Any numeric value up to the range of a **Double**|
|**Variant**<br>_(with **Object** objects)_|<hr>|Same range as **Object**|
|**Variant**<br>_(with objects not of the **Object** type)_|<hr>|Specified by object type|
|**Variant**<BR>_(with [user-defined type](../../How-to/user-defined-data-type.md))_|<hr>|**User-defined type** must be accessed through a [VBE library reference](../../reference/user-interface-help/references-dialog-box.md); range specified for the non-intrinsic **user-defined type** data type (in previous table) also applies.|
|**Variant**<br>_(with special [**Error** sub&#x2011;type](../../reference/user-interface-help/cverr-function.md) values)_|<hr>|Corresponds to valid [error numbers](../../glossary/vbe-glossary.md#error-number)|
|**Variant**<BR>_(with special values [**Empty**](../../Glossary/vbe-glossary.md#empty), [**Null**](../../Glossary/vbe-glossary.md#null), [**Nothing**](../../Reference/User-Interface-Help/nothing-keyword.md), & the special value representing a [missing procedure argument](../../Reference/User-Interface-Help/ismissing-function.md))_|<hr>|Just the four special values.|


<br/>

> [!NOTE] 
> [LongPtr](longptr-data-type.md) is not a true data type because it transforms to a [Long](long-data-type.md) in 32-bit environments, or a [LongLong](longlong-data-type.md) in 64-bit environments. **LongPtr** should be used to represent pointer and handle values in [Declare statements](declare-statement.md) and enables writing portable code that can run in both 32-bit and 64-bit environments.

> [!NOTE] 
> Use the [**StrConv**](../../reference/user-interface-help/strconv-function.md) function to convert a string to a different string 'style' (not primarily a data-type conversion). For example, with the function you can possibly change a string to & from a Unicode string, as well as possibly to a proper-case string.


## Conversion & casting between data types

### Implicit conversions & casts

#### Assignment statements _<sup>(implicit conversions & casts)</sup>_

The following two tables summarize several implicit type conversions & casts that always take place in variable, [property](../../glossary/vbe-glossary.md#property), & [constant](../../glossary/vbe-glossary.md#constant) [assignment statements](../../Concepts/Getting-Started/writing-assignment-statements.md) of values, whenever [by reference](../../glossary/vbe-glossary.md#by-reference) functionality hasn't been established for the identifier meant to be assigned a value in the statement, & that happen so that the assignments still assign potentially useful values.

##### Conversions

|Variable/property/constant type|Value form|
|:--------|:-----------|
|**Variant**|Type same as valid **Variant** sub-type|
| | |
|Intrinsic&nbsp;numerical&nbsp;type apart from the **Boolean**&nbsp;type|Intrinsic numerical type apart from the **Date** type; within the range of the variable/property/constant type|
|**Byte**&nbsp;or **Integer**&nbsp;type|**Date** type; -32768 &le; value &le; 32767|
|**Long**,&nbsp;**Single**,&nbsp;**Double**,&nbsp;or **Currency**&nbsp;type|**Date** type|
|**Boolean** type|Intrinsic numerical type|
| | |
|Intrinsic&nbsp;numerical&nbsp;type|**String** textual representation of a number that parses as a number, and that would be automatically implicitly coerced to the variable/property/constant type in a related variable/property/constant assignment statement|
|**Date** type|**String** textual representation of a valid date, that parses as a date.|
|**Currency** type|**String** textual representation of a valid currency amount, that parses as a currency amount.|
|**String**|Any intrinsic non-object data type; not having **Error** sub-type; not special values **Null**, **Nothing**, an object or an array.|
| | |
|Intrinsic data type|For each variable/property/constant type, the form is that required by the union of all other rules in this table that apply to the particular variable/property/constant type, except that the form applies to sub-type data of a **Variant** value where the **Variant** value is the actual value assigned|

##### Casts

Even though strictly speaking these casts always take place between object types or between an object type & the **Object** type, the style of casting is a kind of interface casting (not object casting.) [&dagger;](#daggerfootnote "VBA doesn't provide object inheritance as a standard mechanism, meaning that conventional object-oriented programming (OOP) object casting isn't fundamentally supported.")

|Variable/property type|Value form|
|:--------|:-----------|
|The&nbsp;**Object**&nbsp;type|Object reference exposing COM's **IDispatch** interface, or can be downcast to such a reference|
|An&nbsp;[interface](../../Glossary/vbe-glossary.md#interface)&nbsp;type|Object type defined using the [**Implements**](../../reference/user-interface-help/implements-statement.md) statement to specify implementation of the interface|
|A&nbsp;specific&nbsp;object&nbsp;type<BR><sup>_(not the **Object** type)_</sup>|Object type defined using the **Implements** statement to specify implementation of the interface derived from the variable/property type|

##### Operations involving a cast & a conversion

If **Variant** data containing an object reference is assigned to a variable or property having either an object data type or the **Object** data type, a conversion & then a cast can occur together.

<BR>

#### Procedure calls _<sup>(implicit conversions & casts)</sup>_

The two tables below summarize known implicit type conversions & casts that always take place for all arguments apart from the last argument, in [procedure calls](../../glossary/vbe-glossary.md#procedure-call). The tables also summarize known implicit type conversions & casts that always take place for the last argument in procedure calls, for calls not on the left-hand side of assignment statements. For type conversions & casts that always take place for the last argument in procedure calls for calls on the left-hand side of assignment statements (such as those that take place in **Procedure Let** & **Procedure Set** procedure calls), see the previous ['Assignment statements _(implicit conversions & casts)_'](#assignment-statements-implicit-conversions--casts) section.

For procedure arguments that aren't variables (such as for constants, literals, properties & [expressions](../../glossary/vbe-glossary.md#expression)), as well as for procedure arguments ['passed by value'](../../glossary/vbe-glossary.md#by-value), more implicit type conversions than those in the following two tables, can take place. The reason for this appears to be because no ['by reference'](../../glossary/vbe-glossary.md#by-reference) functionality needs to be maintained in such cases. The implicit type conversions for such arguments seem likely to be exactly those implicit type conversions that take place in assignment statements, where the procedure parameter is represented by the assignment-statement variable/property/constant & the procedure argument is represented by the assignment-statement value. See the previous ['Assignment statements _(implicit conversions & casts)_'](#assignment-statements-implicit-conversions--casts) section for details on the implicit type conversions that take place in assignment statements.

The implicit conversions & casts listed in the following two tables, convert or cast from start forms to end types. The start form is the form of a value passed as an [argument](../../glossary/vbe-glossary.md#argument) in a standard procedure call. The end type is the type of the internal [parameter](../../glossary/vbe-glossary.md#parameter) for the related argument (parameters are variables accessed by the contents of procedures). Note that arguments that are passed 'by reference' will maintain their 'by reference' functionality even though they be converted or cast in the ways described in the following tables (doesn't apply to last argument of **Procedure Let** & **Procedure Set** procedure calls).

##### Conversions

|Parameter&nbsp;type|Argument form|
|:---------|:-----------|
|**Variant**|Valid **Variant** sub-type|

##### Casts

Even though strictly speaking these casts always take place between object types or between an object type & the **Object** type, the style of casting is a kind of interface casting (not object casting.) [&dagger;](#daggerfootnote "VBA doesn't provide object inheritance as a standard mechanism, meaning that conventional object-oriented programming (OOP) object casting isn't fundamentally supported.")

|Parameter&nbsp;type|Argument form|
|:---------|:-----------|
|The&nbsp;**Object**&nbsp;type|Object reference exposing COM's **IDispatch** interface, or can be downcast to such a reference|
|An&nbsp;interface&nbsp;type|Object type defined using the **Implements** statement to specify implementation of the interface|
|A&nbsp;specific&nbsp;object&nbsp;type<BR><sup>_(not the **Object** type)_</sup>|Object type defined using the **Implements** statement where the statement specifies implementation of the interface derived from the parameter type|

##### Operations involving a cast & a conversion

If a **Variant** argument containing an object reference is assigned to a parameter having either an object data type or the **Object** data type, a conversion & then a cast can occur together.

<BR>
  
### Explicit conversions

See [Type conversion functions](../../concepts/getting-started/type-conversion-functions.md) for examples of how to use the following functions to convert an expression to a specific data type: **CBool**, **CByte**, **CCur**, **CDate**, **CDbl**, **CDec**, **CInt**, **CLng**, **CLngLng**, **CLngPtr**, **CSng**, **[CStr](#returns-for-cstr)**, and **CVar**.

The [**Fix**, and **Int** functions](int-fix-functions.md) provide other forms of integeric conversion.

**[CVErr](cverr-function.md)** can be used to create **Variant** special values of the **Variant** sub-type **Error** from an error number.

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

<br>

### Explicit casts

Explicit casts are not fundamentally supported in the grammar of the VBA language.

<br>

## Verify data types

To verify data types, see the following functions & operators: 

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
- [TypeOf](../../reference/user-interface-help/ifthenelse-statement.md)

|Footnotes|
|:-----------------|
|<sup><a name="asteriskfootnote">\*</a> Read this footnote for an example showing calculation of the storage size for a particular array: the data in a single-dimension array consisting of 4 **Integer** data elements of 2 bytes each occupies 8 bytes; the 8 bytes required for the data plus the 24 bytes of overhead in 32-bit environments, brings the total memory requirement for the array to 32 bytes in 32-bit environments.</sup>|
|<sup><a name="daggerfootnote">&dagger;</a> VBA doesn't provide object inheritance as a standard mechanism, meaning that conventional object-oriented programming (OOP) object casting isn't fundamentally supported.</sup> |

## See also

- [VarType constants](../../concepts/getting-started/vartype-constants.md)
- [Keywords by task](keywords-by-task.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
