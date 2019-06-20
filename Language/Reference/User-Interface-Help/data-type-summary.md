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

Non-intrinsic data types include those in the following table. Note that a VBA compatible [interface](../../Glossary/vbe-glossary.md#interface) type is used in VBA, by using its corresponding [object type](../../Glossary/vbe-glossary.md#object-type) (the interface's object type has the same name as the interface).

|Non&#8209;intrinsic&nbsp;data&nbsp;type|Range|Storage size <sup>_(in bytes)_</sup>|
|:----|:----|----:|
**[User-defined](../../How-to/user-defined-data-type.md)**<BR>_<sup>(defined using [**Type**](../../reference/user-interface-help/type-statement.md) or other means)</sup>_|The range of each element is the same as the range of its data type.|**Σ**&nbsp;(_each&#8239;element's&#8239;storage&#8239;byte&#8239;size_)|
|Specific [object type](../../glossary/vbe-glossary.md#object-type)<BR><sup>_(any object-based type that isn't the intrinsic [**Object**](object-data-type.md) data type)_</sup>|Any [object](../../glossary/vbe-glossary.md#object) of the specific object type, object that _implements_ the public _interface_ of the object-type [class](../../glossary/vbe-glossary.md#class), or the object-based special value [**Nothing**](../../reference/user-interface-help/nothing-keyword.md).|≥&nbsp;_**LongPtr**&#8239;storage&#8239;size_|
|[**Collection**](../../reference/user-interface-help/collection-object.md)&nbsp;or<BR>[**Dictionary**](../../reference/user-interface-help/dictionary-object.md)&nbsp;object<BR><sup>_(examples&nbsp;of&nbsp;specific&nbsp;object&nbsp;types)_</sup>|_See&nbsp;"specific&#8239;object&#8239;type"_|_See&nbsp;"specific&#8239;object&#8239;type"_|
 
## Intrinsic data types

The following tables list the supported intrinsic [data types](../../Glossary/vbe-glossary.md#data-type), & includes information on data-type storage sizes and ranges. The last table documents the intrinsic [**Variant**](../../Glossary/vbe-glossary.md#variant) data type & its sub-types. The tables before the last table, document all the instrinsic data types that do not need the **Variant** type in order for them to be used.

### [Numeric data types](../../glossary/vbe-glossary.md#numeric-data-type)

|Intrinsic data&nbsp;type|Range|Storage size <sup>_(in bytes)_</sup>|
|:----|:----|----:|
|**[Boolean](boolean-data-type.md)**|**True** or **False**|2|
|**[Byte](byte-data-type.md)**|Integer in range from 0 to 255|1|
|**[Integer](integer-data-type.md)**|Integer in range from -_a_ to (_a_ - 1) where _a_ = 32,768|2|
|**[Long](long-data-type.md)**<BR><sup>_(Long integer)_<sup>|Integer in range from -_a_ to (_a_ - 1) where _a_ = 2,147,483,648|4|
|**[LongLong](longlong-data-type.md)**<BR><sup>_(LongLong integer)_<sup>|Integer in range from -_a_ to (_a_ - 1) where _a_ = 9,223,372,036,854,775,808<br/><br/>Valid on 64-bit platforms only.|8|
|**[LongPtr](longptr-data-type.md)**<BR><sup>_(Long&nbsp;integer&nbsp;on 32&#8209;bit&nbsp;systems,<BR>LongLong&nbsp;integer&nbsp;on 64&#8209;bit&nbsp;systems)_<sup>|The range for bit version of number, is every possible bit combination for the bytes specified by the storage size. This integeric range is:<BR>&thinsp;•&nbsp;&nbsp;00&thinsp;00&thinsp;00&thinsp;00&thinsp;<sub>16</sub>&nbsp;→&nbsp;FF&thinsp;FF&thinsp;FF&thinsp;FF&thinsp;<sub>16</sub><BR>&nbsp;&nbsp;&nbsp;on 32-bit systems, \&<BR> &thinsp;•&nbsp;&nbsp;00&thinsp;00&thinsp;00&thinsp;00&thinsp;**00&thinsp;00 00&thinsp;00**&thinsp;<sub>16</sub>&nbsp;→&nbsp;FF&thinsp;FF&thinsp;FF&thinsp;FF&thinsp;**FF&thinsp;FF FF&thinsp;FF**&thinsp;<sub>16</sub><BR>&nbsp;&nbsp;&nbsp;on 64-bit systems.|On&nbsp;32&#8209;bit&nbsp;systems,&nbsp;4.<br/><br/>On&nbsp;64&#8209;bit&nbsp;systems,&nbsp;8.|
|**[Single](single-data-type.md)**<BR><sup>_(single&#8209;precision floating&#8209;point)_</sup>|The special values +∞, -∞, NaN & negative zero, plus values from -_a_ to _a_ where _a_ = 3<sub>&#8226;</sub>402823 × 10<sup>38</sup>.<BR><BR>The value mustn't have more than 24 significant binary digits & 126 binary places, when the value is a non-special value with an absolute value ≥ 2<sup>−126</sup>.<BR><BR>When 0 < &vert;_the value_&vert; < 2<sup>−126</sup>, the value must be able to be written as ±p ÷ 2<sup>(126 + 23)</sup> for some natural number _p_ for which &vert;_p_&vert; ≤ (2<sup>23</sup> - 1), except if the value is the special negative-zero value.<BR><BR>Note that all integers having no more than 6 significant decimal digits, and having absolute values no more than 3<sub>&#8226;</sub>402823 × 10<sup>38</sup>, are in the range.|4|
|**[Double](double-data-type.md)** <BR><sup>_(double&#8209;precision floating-point)_</sup>|-_a_ to (_a_ + 10<sup>-14</sup>) where _a_ = 1<sub>&#8226;</sub>79769313486231 × 10<sup>308</sup>|8|
|**[Date](date-data-type.md)**|Date & time in range from&nbsp;&nbsp;January&thinsp;1,&thinsp;100&thinsp;AD,&thinsp;00:00.00&thinsp;AM,&nbsp;&nbsp;to&nbsp;&nbsp;December&thinsp;31,&thinsp;9999&thinsp;AD,&thinsp;23:59.59&thinsp;PM&nbsp;&nbsp;whenever the time can be accurately expressed purely as an integeric number of seconds after midnight; certain dates & times within the upper- & lower-bounds of this range, where the times are not expressible as such, are also within the range of this data type.|8|
|**[Currency](currency-data-type.md)**<BR><sup>_(scaled integer)_</sup>|-_a_&nbsp;&nbsp;to&nbsp;&nbsp;(_a_&thinsp;-&thinsp;0<sub>&#8226;</sub>0001)&nbsp;&nbsp;where<BR>&nbsp;&nbsp;&nbsp;_a_ ⪆ 9 × 10<sup>14</sup> &<BR>&nbsp;&nbsp;&nbsp;value must have no more than 4 decimal places,<p align="right">_a_ = 922,337,203,685,477<sub>&#8226;</sub>5808 (≈ 922 trillion).</p>|8|

> [!NOTE] 
> [LongPtr](longptr-data-type.md) is not a true data type because it transforms to a [Long](long-data-type.md) in 32-bit environments, or a [LongLong](longlong-data-type.md) in 64-bit environments. **LongPtr** should be used to represent pointer and handle values in [Declare statements](declare-statement.md) and enables writing portable code that can run in both 32-bit and 64-bit environments.

<BR><BR>

### [String](../../glossary/vbe-glossary.md#string-data-type) data types

|Kind of intrinsic String&nbsp;data&nbsp;type|Range|Storage size <sup>_(in bytes)_</sup>|
|:----|:----|----:|
|_Variable&#8209;length_|Each character can be any Unicode character; length of string can be changed to any non-negative integer up to approximately 2 billion.|4<BR>+&nbsp;&#8239;_**LongPtr**&#8239;storage&#8239;size_<BR>+&nbsp;&#8239;((_length&#8239;of&#8239;string_&nbsp;&#8239;+&nbsp;&#8239;1)&nbsp;&#8239;&times;&nbsp;&#8239;2)|
|_Fixed&#8209;length_|Each character can be any Unicode character; length of string can be set to any positive integer up to approximately 65,400; string length can only be set during execution of the variable's declaration|_String&#8239;length_&nbsp;&#8239;&times;&nbsp;&#8239;2|

<BR><BR>

### Family of [array](../../concepts/getting-started/using-arrays.md) data types

|Array<BR>data&#8209;type<BR>family|Range|Storage size <sup>_(in bytes)_</sup>|
|:----|:----|----:|
|For arrays in general|Each element must have the same data type. The element data type in VBA, is chosen when executing the [variable's declaration](../../concepts/getting-started/declaring-variables.md)&mdash;any data type other than an array data type, can be chosen for the element data type. Once chosen, another cannot be chosen whilst the code is running. Each element has the same range as the chosen data type. The element configuration can have up to 60 dimensions, and has a maximum size limited by your operating system & amount of available RAM. The index range for each dimension is some contiguous set of integers or just one particular integer, specified in the element configuration.<BR><BR>Variable must be declared as either a fixed-size array or a dynamic (re-sizeable / variable-size) array. During execution, if the variable is a fixed-size array, the variable cannot accommodate a dynamic array, & vice versa.<BR><BR>Range is further restricted according to whether the array is a fixed-size or dynamic array (see below).|16<BR>**+**&nbsp;&#8239;_**LongPtr**&#8239;storage&#8239;byte&#8239;size_<BR>**+**&nbsp;&#8239;(8&nbsp;&#8239;**&times;**&nbsp;&#8239;_no.&#8239;of&#8239;array&#8239;dimensions_)<BR>**+**&nbsp;&#8239;**&Sigma;**&nbsp;&#8239;(_each&#8239;element's<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#8239;storage&#8239;byte&#8239;size_)<br><sup>[*](#asteriskfootnote "Read this footnote for an example showing calculation of the storage size for a particular array: the data in a single-dimension array consisting of 4 Integer data elements of 2 bytes each occupies 8 bytes; the 8 bytes required for the data plus the 24 bytes of overhead in 32-bit environments brings the total memory requirement for the array to 32 bytes in 32-bit environments.")</sup>|
|For<BR>fixed&#8209;size<BR>arrays|Fixed-size arrays have for their entire run-time life-times, immutable element configurations. Their element configurations are programatically specified within VBA through VBA variable declarations.<BR><BR>Range specification for _"For&thinsp;arrays&thinsp;in&thinsp;general"_ array data&#8209;type family also applies.|See _"For&thinsp;arrays&thinsp;in&thinsp;general"_|
|For<BR>dynamic<BR>arrays|Dynamic arrays, in contrast to fixed-size arrays, have no initial element configuration, & their element configurations can be re-specified during run-time as many times as is needed. Their element configurations are programatically specified within VBA through [**ReDim**](../../reference/user-interface-help/redim-statement.md) statements.<BR><BR>Range specification for _"For&thinsp;arrays&thinsp;in&thinsp;general"_ array data&#8209;type family also applies.|See _"For&thinsp;arrays&thinsp;in&thinsp;general"_|

<BR><BR>

### The [**Variant**](variant-data-type.md) type

|Kind of '**Variant**&nbsp;data&#8209;type'&nbsp;data|Range|Storage&nbsp;size <sup>_(in&nbsp;bytes)_</sup>|
|:--------|:-----------|----:|
|Any|Data of:<br><ul><li>any non-array data type except the fixed-length **String** type & user-defined types declared in the VBE using VBA's **Type** statement,</li><li>any array data type except if the element type is the fixed-length **String** type or a user-defined type declared in the VBE using VBA's **Type** statement,</li><li>the **Variant** [**Error**](../../reference/user-interface-help/cverr-function.md) sub-type, _or_</li><li>the **Variant** [**Decimal**](../../reference/user-interface-help/decimal-data-type.md) sub-type,</li></ul>_or_ one of the following special values:<ul><li>[**Empty**](../../Glossary/vbe-glossary.md#empty) <sup>_(**Variant** value)_.</sup></li><li>[**Null**](../../Glossary/vbe-glossary.md#null) <sup>_(**Variant** value)_.</sup></li><li>[**Nothing**](../../Reference/User-Interface-Help/nothing-keyword.md) <sup>_(object-based value)._</sup></li><li>the special value representing a [missing procedure argument](../../Reference/User-Interface-Help/ismissing-function.md) (a **Variant** value).</li></ul>|≥ 16<BR>(exact amount depends on specific kind of data)|
|Arrays|Any array-data-type array.|(16&nbsp;&#8239;-&nbsp;&#8239;_**LongPtr**&#8239;storage&#8239;size_)&thinsp; more than when stored in variable declared as an array.|
|Characters|Same range as for variable-length **String**|22<BR>+&nbsp;&#8239;(_string&#8239;length_&nbsp;&#8239;&times;&nbsp;&#8239;2)|
|Number within the range of one of VBA's (non-composite) [numeric types](../../glossary/vbe-glossary.md#numeric-type)|Any such number.|16|
|Number&nbsp;of&nbsp;**[Decimal](decimal-data-type.md)**&nbsp;data&nbsp;type|Integer with an absolute value less than 79,228,162,514,264,337,593,543,950,336, or a number from this range after it has been scaled down by 10<sup>a</sup> where _a_ can be any natural number under 29.|16|
|**Boolean** data-type value|**True** or **False**.|16|
|**Date** data-type value|Same as the **Date** data type.|16|
|Objects of the intrinsic [**Object**](object-data-type.md) data type|Same range as the **Object** data type.|(16&nbsp;&#8239;-&nbsp;&#8239;_**LongPtr**&#8239;storage&#8239;size_) &thinsp;more than when stored in variable declared as having the **Object**&nbsp;data&nbsp;type.|
|_"Specific&nbsp;[object&nbsp;type](../../glossary/vbe-glossary.md#object-type)"_&nbsp;objects<BR><sup>_(objects not of the **Object** data type)_</sup>|Range for _"specific&#8239;object&#8239;type"_ from earlier section applies.|(16&nbsp;&#8239;-&nbsp;&#8239;_**LongPtr**&#8239;storage&#8239;size_) &thinsp;more than when stored in variable declared as having _"specific&#8239;object&#8239;type"_ data type.|
|[User-defined type](../../How-to/user-defined-data-type.md)|**User-defined type** must be accessed through a [VBE library reference](../../reference/user-interface-help/references-dialog-box.md); range specified for the non-intrinsic **user-defined type** data type (in earlier section) also applies.|16 more than when stored in variable declared as having 'user&#8209;defined&#8239;type' data type.|
|Special [**Error** sub&#x2011;type](../../reference/user-interface-help/cverr-function.md) values|Corresponds to valid [error numbers](../../glossary/vbe-glossary.md#error-number)|16|
|Special values [**Empty**](../../Glossary/vbe-glossary.md#empty), [**Null**](../../Glossary/vbe-glossary.md#null), [**Nothing**](../../Reference/User-Interface-Help/nothing-keyword.md), & the special value representing a [missing procedure argument](../../Reference/User-Interface-Help/ismissing-function.md)|Just the four special values.|16|

<br/>

> [!NOTE] 
> Use the [**StrConv**](../../reference/user-interface-help/strconv-function.md) function to convert a string to a different string 'style' (not primarily a data-type conversion). For example, with the function you can possibly change a string to & from a Unicode string, as well as possibly to a proper-case string.


## Conversion & casting between data types

### Implicit conversions & casts

#### Assignment statements _<sup>(implicit conversions & casts)</sup>_

The following tables in this section summarize several implicit type conversions & casts that always take place in variable, [property](../../glossary/vbe-glossary.md#property), & [constant](../../glossary/vbe-glossary.md#constant) [assignment statements](../../Concepts/Getting-Started/writing-assignment-statements.md) of values, whenever [by reference](../../glossary/vbe-glossary.md#by-reference) functionality hasn't been established for the identifier meant to be assigned a value in the statement, & that happen so that the assignments still assign potentially useful values.

##### Conversions

###### Variant related

|Variable/property/constant data type|Value form|
|:--------|:-----------|
|**Variant**|Any data not held in a **Variant**, that a **Variant** can directly hold.<BR><sup>_(See the [section above on the **Variant** type](#the-variant-type) for details on this.)_</sup>|
|Intrinsic data type|For each variable/property/constant data type, the form is that required by the union of all other rules in this table that apply to the particular variable/property/constant data type, except that the form applies to sub-type data of a **Variant** value where the **Variant** value is the actual value assigned.|

###### String related

|Variable/property/constant data type|Value form|
|:--------|:-----------|
|[Numeric&nbsp;data&nbsp;type](../../glossary/vbe-glossary.md#numeric-data-type)|**String** textual representation of a number that parses as a number, and that would be automatically implicitly coerced to the variable/property/constant data type in a related variable/property/constant assignment statement|
|**Date** type|**String** textual representation of a valid date, that parses as a date.|
|**Currency** type|**String** textual representation of a valid currency amount, that parses as a currency amount.|
|**String**|Any intrinsic non-object & non-array data type; not having **Error** sub-type; not special values **Null**, **Nothing**, an object or an array.|

###### 'Numeric type to numeric type' conversions

|Variable/property/constant data type|Value form|
|:--------|:-----------|
|Numeric&nbsp;data&nbsp;type apart from the **Boolean**&nbsp;type|Numeric data type apart from the **Date** type, or the **Variant** [**Decimal**](decimal-data-type.md) sub-type; not above the upper-bound, or beneath the lower-bound, of the range of the variable/property/constant data type.|
|**Byte**&nbsp;or **Integer**&nbsp;type|**Date** type;&nbsp;&nbsp;&nbsp;-32,768 &le; _value_ &le; 32,767|
|**Long**,&nbsp;**Single**,&nbsp;**Double**,&nbsp;or **Currency**&nbsp;type|**Date** type|
|**Boolean** type|Numeric data type or the **Variant** **Decimal** sub-type|

##### Casts

Even though strictly speaking these casts always take place between [object types](../../glossary/vbe-glossary.md#object-type) or between an object type & the [**Object** data type](../../glossary/vbe-glossary.md#object-data-type), the style of casting is a kind of interface casting (not object casting.) [&dagger;](#daggerfootnote "VBA doesn't provide object inheritance as a standard mechanism, meaning that conventional object-oriented programming (OOP) object casting isn't fundamentally supported.")

|Variable/property type|Value form|
|:--------|:-----------|
|The&nbsp;**Object**&nbsp;data&nbsp;type|Object reference exposing COM's **IDispatch** interface, or can be downcast to such a reference|
|A&nbsp;specific&nbsp;object&nbsp;type<BR><sup>_(not the **Object** data type)_</sup>|Object type defined using the **Implements** statement to specify implementation of the interface derived from the variable/property type|
|An&nbsp;[interface](../../Glossary/vbe-glossary.md#interface)&nbsp;type|Object type defined using the [**Implements**](../../reference/user-interface-help/implements-statement.md) statement to specify implementation of the interface|

##### Operations involving a cast & a conversion

If **Variant** data containing an object reference is assigned to a variable or property having either an object type or the **Object** data type, a conversion & then a cast can occur together.

<BR>

#### Procedure calls _<sup>(implicit conversions & casts)</sup>_

The two tables below summarize known implicit type conversions & casts that always take place for all arguments apart from the last argument, in [procedure calls](../../glossary/vbe-glossary.md#procedure-call). The tables also summarize known implicit type conversions & casts that always take place for the last argument in procedure calls, for calls not on the left-hand side of assignment statements. For type conversions & casts that always take place for the last argument in procedure calls for calls on the left-hand side of assignment statements (such as those that take place in **Procedure Let** & **Procedure Set** procedure calls), see the previous ['Assignment statements _(implicit conversions & casts)_'](#assignment-statements-implicit-conversions--casts) section.

For procedure arguments that aren't variables (such as for constants, literals, properties & [expressions](../../glossary/vbe-glossary.md#expression)), as well as for procedure arguments ['passed by value'](../../glossary/vbe-glossary.md#by-value), more implicit type conversions than those in the following two tables, can take place. The reason for this appears to be because no ['by reference'](../../glossary/vbe-glossary.md#by-reference) functionality needs to be maintained in such cases. The implicit type conversions for such arguments seem likely to be exactly those implicit type conversions that take place in assignment statements, where the procedure parameter is represented by the assignment-statement variable/property/constant & the procedure argument is represented by the assignment-statement value. See the previous ['Assignment statements _(implicit conversions & casts)_'](#assignment-statements-implicit-conversions--casts) section for details on the implicit type conversions that take place in assignment statements.

The implicit conversions & casts listed in the following two tables, convert or cast from start forms to end types. The start form is the form of a value passed as an [argument](../../glossary/vbe-glossary.md#argument) in a standard procedure call. The end type is the type of the internal [parameter](../../glossary/vbe-glossary.md#parameter) for the related argument (parameters are variables accessed by the contents of procedures). Note that arguments that are passed 'by reference' will maintain their 'by reference' functionality even though they be converted or cast in the ways described in the following tables (doesn't apply to last argument of **Procedure Let** & **Procedure Set** procedure calls).

##### Conversions

|Parameter&nbsp;data&nbsp;type|Argument form|
|:---------|:-----------|
|**Variant**|Any data not held in a **Variant**, that a **Variant** can directly hold.<BR><sup>_(See the [section above on the **Variant** type](#the-variant-type) for details on this.)_</sup>|

##### Casts

Even though strictly speaking these casts always take place between [object types](../../glossary/vbe-glossary.md#object-type) or between an object type & the [**Object** data type](../../glossary/vbe-glossary.md#object-data-type), the style of casting is a kind of interface casting (not object casting.) [&dagger;](#daggerfootnote "VBA doesn't provide object inheritance as a standard mechanism, meaning that conventional object-oriented programming (OOP) object casting isn't fundamentally supported.")

|Parameter&nbsp;data&nbsp;type|Argument form|
|:---------|:-----------|
|The&nbsp;**Object**&nbsp;data&nbsp;type|Object reference exposing COM's **IDispatch** interface, or can be downcast to such a reference|
|A&nbsp;specific&nbsp;object&nbsp;type<BR><sup>_(not the **Object** data type)_</sup>|Object type defined using the **Implements** statement where the statement specifies implementation of the interface derived from the parameter type|
|An&nbsp;interface&nbsp;type|Object type defined using the **Implements** statement to specify implementation of the interface|

##### Operations involving a cast & a conversion

If a **Variant** argument containing an object reference is assigned to a parameter having either an object type or the **Object** data type, a conversion & then a cast can occur together.

<BR>
  
### Explicit conversions

See [Type conversion functions](../../concepts/getting-started/type-conversion-functions.md) for examples of how to use the following functions to convert an expression to a specific data type: **CBool**, **CByte**, **CCur**, **CDate**, **CDbl**, **CDec**, **CInt**, **CLng**, **CLngLng**, **CLngPtr**, **CSng**, **[CStr](#returns-for-cstr)**, and **CVar**.

The [**Fix**, and **Int** functions](int-fix-functions.md) provide other forms of integeric conversion.

**[CVErr](cverr-function.md)** can be used to create **Variant** special values of the **Variant** sub-type **Error**, from an error number.

The [**LSet**](../../reference/user-interface-help/lset-statement.md) statement can be used to convert a value in one [user-defined type](../../How-to/user-defined-data-type.md), to a value of another user-defined type.

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
