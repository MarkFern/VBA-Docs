---
title: VarType constants (VBA)
keywords: vblr6.chm1012527
f1_keywords:
- vblr6.chm1012527
ms.prod: office
ms.assetid: 169a159e-7494-56cf-e7ca-01da5bd9705d
ms.date: 12/26/2018
localization_priority: Normal
---


# VarType constants

The following [constants](../../Glossary/vbe-glossary.md#constant) specified in VBA's **VbVarType** [enumeration](../../reference/user-interface-help/enum-statement.md), can be used anywhere in your code in place of the actual values. At the time of writing, all of them are used by the [**VarType**](../../Reference/User-Interface-Help/vartype-function.md) function.

|Constant|Value|Description|Corresponding [VARENUM](https://docs.microsoft.com/en-us/windows/desktop/api/wtypes/ne-wtypes-varenum)&nbsp;constant in [OLE Automation Protocol specification](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f)|
|:-----|:-----|:-----|:-----|
|**vbEmpty**|0|[**Empty**](../../Glossary/vbe-glossary.md#empty) value <sup>_(represents uninitialized [variable](../../glossary/vbe-glossary.md#variable))_</sup> <sup>[*](#asteriskfootnote "Default.")</sup> <sup>[&dagger;](#daggerfootnote "Variant special value.")</sup>|VT_EMPTY|
|**vbNull**|1|[**Null**](../../Glossary/vbe-glossary.md#null) value <sup>_(contains no valid data)_</sup> <sup>[&dagger;](#daggerfootnote "Variant special value.")</sup>|VT_NULL|
|**vbInteger**|2|Integer of data type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|VT_I2|
|**vbLong**|3|[Long](../../reference/User-Interface-Help/long-data-type.md) integer|VT_I4|
|**vbSingle**|4|[Single](../../Glossary/vbe-glossary.md#single-data-type) value <sup>_(single-precision floating-point number)_</sup>|VT_R4|
|**vbDouble**|5|[Double](../../Glossary/vbe-glossary.md#double-data-type) value <sup>_(double-precision floating-point number)_</sup>|VT_R8|
|**vbCurrency**|6|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|VT_CY|
|**vbDate**|7|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|VT_DATE|
|**vbString**|8|[String](../../Glossary/vbe-glossary.md#string-data-type)|VT_BSTR|
|**vbObject**|9|Has either of the following forms:<br><table><tr><td>i) A (VBA) object-based type with a particular [_interface_](../../Glossary/vbe-glossary.md#interface) chosen[<sup>&para;</sup>](#paragraphfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, & that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, & the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts."), where the chosen interface directly exposes COM's **IDispatch** interface. <sup>[&Dagger;](#doubledaggerfootnote "That the chosen interface exposes IDispatch means that the particular object-based type of the data, can be directly used with (OLE) Automation late-binding technology. Object references of such types can be cast to the Object data type if they are not already of the type.")</sup></td></tr><tr><td>ii) [**Nothing**](../../reference/user-interface-help/nothing-keyword.md) value <sup>_(special value)_</sup> as a literal.</sup></td></tr></table>||VT_DISPATCH|
|**vbError**|10|Has either of the following forms:<br><table><tr><td>i) An [**Error**](../../reference/user-interface-help/cverr-function.md) value.</td></tr><tr><td>ii) The [parameter](../../glossary/vbe-glossary.md#parameter) for a [_missing_](../../reference/user-interface-help/ismissing-function.md) [_optional_](../../concepts/getting-started/understanding-named-arguments-and-optional-arguments.md) **Variant** [argument](../../glossary/vbe-glossary.md#argument) of some procedure, that hasn't yet had a conventional value assignment (the "missing" flag bit will have been set), or a variable or [property](../../glossary/vbe-glossary.md#property) holding the value of such a parameter. At the time of writing, such values are also **Error** values of the [error number 448](../../reference/user-interface-help/named-argument-not-found-error-448.md).</td></tr></table><sup>[&dagger;](#daggerfootnote "Variant special value.")</sup>|VT_ERROR|
|**vbBoolean**|11|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|VT_BOOL|
|**vbVariant**|12|[**Variant**](../../Glossary/vbe-glossary.md#variant-data-type)<BR><sup>_(used for return value only when added to **vbArray** constant to signify an [array](../../Glossary/vbe-glossary.md#array) of variants)_</sup>|VT_VARIANT|
|**vbDataObject**|13|A (VBA) object-based type with a particular _interface_ chosen[<sup>&para;</sup>](#paragraphfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, & that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, & the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts."), that is not represented by the **vbObject** constant documented in this table. <sup>[&sect;](#sectionfootnote "An object of such an object-based type, like all VBA objects, is still a COM object. Like all COM objects and interfaces, such objects expose COM's IUnknown interface. Not to be confused with ActiveX Data Objects (ADO) which is a database technology.")</sup>|VT_UNKNOWN|
|**vbDecimal**|14|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|VT_DECIMAL|
|**vbByte**|17|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) integer|VT_UI1|
|**vbLongLong**|20|[LongLong](../../reference/User-Interface-Help/long-data-type.md) integer <sup>_(valid on 64-bit platforms only)_</sup>|VT_I8|
|**vbUserDefinedType**|36|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type)|VT_RECORD|
|**vbArray**|8192|Array|VT_ARRAY|

<table>
 <tr><td><a name="asteriskfootnote"><sup>*</sup></a></td><td> 
   
   Default.<BR></td></tr>
 <tr><td><a name="daggerfootnote"><sup>&dagger;</sup></a></td><td>
  
   **Variant** special value.</td></tr>
 <tr><td><a name="doubledaggerfootnote"><sup>&Dagger;</sup></a></td><td>  

 That the **chosen** _interface_[<sup>&para;</sup>](#paragraphfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, & that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, & the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts.") exposes **IDispatch** means that the particular object-based type of the data, can be directly used with (OLE) Automation late-binding technology. <sup>[&Vert;](#doubleverticalbarfootnote)</sup><BR>Object references of such types can be [cast](../../Reference/User-Interface-Help/data-type-summary.md#implicit-conversions--casts) to the [**Object**](../../reference/user-interface-help/object-data-type.md) data type if they are not already of the type.</td></tr>
 <tr><td><a name="sectionfootnote"><sup>&sect;</sup></a></td><td>

 An object of such an object-based type, like all VBA objects, is still a COM object. Like all COM objects and _interfaces_, such objects expose COM's **IUnknown** interface. Not to be confused with [ActiveX Data Objects (ADO)](../../../access/concepts/activex-data-objects/set-properties-of-activex-data-objects-in-visual-basic.md) which is a database technology. <sup>[&Vert;](#doubleverticalbarfootnote)</sup></td></tr>
<tr><td><a name="doubleverticalbarfootnote"><sup>&Vert;</sup></a></td><td>
  
   The glossary definition of [ActiveX object](../../Glossary/vbe-glossary.md#activex-object) in the VBA documentation on 7th April 2019 (current date), indicates that ActiveX objects are (OLE) Automation objects. However, various developers instead use ActiveX as a synonym for the COM technology, meaning that those developers also class non-OLE-Automation COM objects as being a certain type of ActiveX object.</td></tr>
<tr><td><a name="paragraphfootnote"><sup>&para;</sup></a></td><td>

 Chosen _interface_ is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's [class](../../Glossary/vbe-glossary.md#class) definition as being _implemented_ by the class (object's interfaces), that is chosen during run-time, & that has to be chosen before conventional execution of any of the object's [methods](../../glossary/vbe-glossary.md#method) or conventional access of any of the object's [properties](../../glossary/vbe-glossary.md#property), can take place. If an interface needs to be chosen for an object, & the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of [implicit type casts](../../Reference/User-Interface-Help/data-type-summary.md#implicit-conversions--casts).</td></tr>
</table>

## See also

- [Data type summary](../../reference/user-interface-help/data-type-summary.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
