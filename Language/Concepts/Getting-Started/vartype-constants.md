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

The following [constants](../../Glossary/vbe-glossary.md#constant) specified in VBA's **VbVarType** [enumeration](../../reference/user-interface-help/enum-statement.md), can be used anywhere in your code in place of the actual values.

|Constant|Value|Description|Corresponding [VARENUM](https://docs.microsoft.com/en-us/windows/desktop/api/wtypes/ne-wtypes-varenum)&nbsp;constant in [OLE Automation Protocol specification](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f)|
|:-----|:-----|:-----|:-----|
|**vbEmpty**|0|[**Empty**](../../Glossary/vbe-glossary.md#empty) value <sup>_(represents uninitialized variable)_</sup> <sup>*</sup> <sup>&dagger;</sup>|VT_EMPTY|
|**vbNull**|1|[**Null**](../../Glossary/vbe-glossary.md#null) value <sup>_(contains no valid data)_</sup> <sup>&dagger;</sup>|VT_NULL|
|**vbInteger**|2|Integer of data type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|VT_I2|
|**vbLong**|3|[Long](../../reference/User-Interface-Help/long-data-type.md) integer|VT_I4|
|**vbSingle**|4|[Single](../../Glossary/vbe-glossary.md#single-data-type) value <sup>_(single-precision floating-point number)_</sup>|VT_R4|
|**vbDouble**|5|[Double](../../Glossary/vbe-glossary.md#double-data-type) value <sup>_(double-precision floating-point number)_</sup>|VT_R8|
|**vbCurrency**|6|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|VT_CY|
|**vbDate**|7|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|VT_DATE|
|**vbString**|8|[String](../../Glossary/vbe-glossary.md#string-data-type)|VT_BSTR|
|**vbObject**|9|A (VBA) [object](../../glossary/vbe-glossary.md#object) with a particular interface chosen, where the chosen interface directly exposes COM's **IDispatch** interface. <sup>&Dagger;</sup>|VT_DISPATCH|
|**vbError**|10|Has either of the following forms:<br><table><tr><td>i) An [**Error**](../../reference/user-interface-help/cverr-function.md) value.</td></tr><tr><td>ii) The [parameter](../../glossary/vbe-glossary.md#parameter) for a [_missing_](../../reference/user-interface-help/ismissing-function.md) [_optional_](../../concepts/getting-started/understanding-named-arguments-and-optional-arguments.md) **Variant** [argument](../../glossary/vbe-glossary.md#argument) of some procedure, that hasn't yet had a conventional value assignment (the "missing" flag bit will have been set), or a variable holding the value of such a parameter. <sup>&dagger;</sup></td></tr></table>|VT_ERROR|
|**vbBoolean**|11|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|VT_BOOL|
|**vbVariant**|12|[**Variant**](../../Glossary/vbe-glossary.md#variant-data-type) <sup>_(used for return value only when added to **vbArray** constant to signify an [array](../../Glossary/vbe-glossary.md#array) of variants)_</sup>|VT_VARIANT|
|**vbDataObject**|13|A (VBA) object not represented by the **vbObject** constant documented in this table. <sup>&sect;</sup>|VT_UNKNOWN|
|**vbDecimal**|14|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|VT_DECIMAL|
|**vbByte**|17|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) integer|VT_UI1|
|**vbLongLong**|20|[LongLong](../../reference/User-Interface-Help/long-data-type.md) integer <sup>_(valid on 64-bit platforms only)_</sup>|VT_I8|
|**vbUserDefinedType**|36|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type)|VT_RECORD|
|**vbArray**|8192|Array|VT_ARRAY|

<table>
 <tr><td><sup>*</sup></td><td> 
   
   Default.<BR></td></tr>
 <tr><td><sup>&dagger;</sup></td><td>
  
   **Variant** special value.</td></tr>
 <tr><td><sup>&Dagger;</sup></td><td>  

   That the object implements **IDispatch** means that the VBA COM object that encompasses the passed object reference, is an (OLE) [Automation object](../../Glossary/vbe-glossary.md#automation-object-1). That the **chosen** interface exposes **IDispatch** means that the particular object reference that has been passed, can be directly used with (OLE) Automation technology. <sup>&brvbar;</sup><BR>Such object references can be cast to the [**Object**](../../reference/user-interface-help/object-data-type.md) data type. <br> </td></tr>
 <tr><td><sup>&sect;</sup></td><td>

   Such an object, like all VBA objects, is still a COM object. Like all COM objects and interfaces, such an object exposes COM's **IUnknown** interface. Not to be confused with [ActiveX Data Objects (ADO)](../../../access/concepts/activex-data-objects/set-properties-of-activex-data-objects-in-visual-basic.md) which is a database technology. <sup>&brvbar;</sup></td></tr>
 <tr><td><sup>&brvbar;</sup></td><td>
  
   The glossary definition of [ActiveX object](../../Glossary/vbe-glossary.md#activex-object) in the VBA documentation on 7th April 2019 (current date), indicates that ActiveX objects are Automation objects. However, various developers instead use ActiveX as a synonym for the COM technology, meaning that those developers also class non-Automation COM objects as being a certain type of ActiveX object.</td></tr>
</table>
  
## See also

- [Data type summary](../../reference/user-interface-help/data-type-summary.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
