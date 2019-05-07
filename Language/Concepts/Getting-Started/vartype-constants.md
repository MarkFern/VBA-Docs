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

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbEmpty**|0|[**Empty**](../../Glossary/vbe-glossary.md#empty) value _(variable uninitialized)_ * &#8224;|
|**vbNull**|1|[**Null**](../../Glossary/vbe-glossary.md#null) value _(contains no valid data)_ &#8224;|
|**vbInteger**|2|Integer of data type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|
|**vbLong**|3|[Long](../../reference/User-Interface-Help/long-data-type.md) integer|
|**vbSingle**|4|[Single](../../Glossary/vbe-glossary.md#single-data-type) value _(single-precision floating-point number)_|
|**vbDouble**|5|[Double](../../Glossary/vbe-glossary.md#double-data-type) value _(double-precision floating-point number)_|
|**vbCurrency**|6|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|
|**vbDate**|7|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|
|**vbString**|8|[String](../../Glossary/vbe-glossary.md#string-data-type)|
|**vbObject**|9|A (VBA) [object](../../glossary/vbe-glossary.md#object) with a particular interface chosen, where the chosen interface directly exposes the **IDispatch** interface. That the object implements **IDispatch** means that the VBA COM object that encompasses the passed object reference, is an (OLE) [Automation object](../../Glossary/vbe-glossary.md#automation-object-1). That the **chosen** interface implements **IDispatch** means that the particular object reference that has been passed, can be directly used with (OLE) Automation technology. &Dagger;|
|**vbError**|10|An [**Error**](../../reference/user-interface-help/cverr-function.md) value|
|**vbBoolean**|11|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|
|**vbVariant**|12|[**Variant**](../../Glossary/vbe-glossary.md#variant-data-type) _(used for return value only when added to **vbArray** constant to signify an [array](../../Glossary/vbe-glossary.md#array) of variants)_|
|**vbDataObject**|13|A (VBA) object not represented by the **vbObject** constant documented in this table. Not to be confused with [ActiveX Data Objects (ADO)](../../../access/concepts/activex-data-objects/set-properties-of-activex-data-objects-in-visual-basic.md) which is a database technology. &Dagger;|
|**vbDecimal**|14|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|
|**vbByte**|17|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) integer|
|**vbLongLong**|20|[LongLong](../../reference/User-Interface-Help/long-data-type.md) integer _(valid on 64-bit platforms only)_|
|**vbUserDefinedType**|36|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type)|
|**vbArray**|8192|Array|

<sup>* Default.</sup><BR>
<sup>&#8224; **Variant** special value.</sup><BR>
<sup>&Dagger; The glossary definition of [ActiveX object](../../Glossary/vbe-glossary.md#activex-object) in the VBA documentation on 7th April 2019 (current date), indicates that ActiveX objects are Automation objects. However, various developers instead use ActiveX as a synonym for the COM technology, meaning that those developers also class non-Automation COM objects as being a certain type of ActiveX object.</sup>
  
## See also

- [Data type summary](../../reference/user-interface-help/data-type-summary.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
