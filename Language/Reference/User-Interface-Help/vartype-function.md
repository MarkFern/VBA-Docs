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

Returns an **Integer** where the returned value will indicate one of the following things, the choice of which depends upon the [argument](../../Glossary/vbe-glossary.md#argument) passed:
1) The subtype or type of a [variable](../../Glossary/vbe-glossary.md#variable), [expression](../../glossary/vbe-glossary.md#expression) or other kind of value.
2) The type or lack of type for an object's default member's return value.
3) The [**Variant**](../../Glossary/vbe-glossary.md#variant-data-type) special value that a **Variant** variable, expression or other kind of value, evaluates to.

## Syntax

**VarType**(_arg_)

The required _arg_ argument must be of the **Variant** type, or be able to be [coerced](../../Reference/User-Interface-Help/data-type-summary.md#implicit-conversions--casts) to it.
 
## Return values

Return value is either:

- just one of the following constants excluding the **vbArray** constant & the **vbVariant** constant, _or_
- the **vbArray** constant added to any of the other constants from the following list.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbEmpty**|0|[**Empty**](../../Glossary/vbe-glossary.md#empty) value <sup>_(represents uninitialized variable)_</sup> [\*](#asteriskfootnote "Variant special value.")|
|**vbNull**|1|[**Null**](../../Glossary/vbe-glossary.md#null) value <sup>_(represents no valid data)_</sup> [\*](#asteriskfootnote  "Variant special value.")|
|**vbInteger**|2|Integer of data type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|
|**vbLong**|3|[Long](../../Glossary/vbe-glossary.md#long-data-type) integer|
|**vbSingle**|4|[Single](../../Glossary/vbe-glossary.md#single-data-type) value <sup>_(single-precision floating-point number)_</sup>|
|**vbDouble**|5|[Double](../../Glossary/vbe-glossary.md#double-data-type) value <sup>_(double-precision floating-point number)_</sup>|
|**vbCurrency**|6|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|
|**vbDate**|7|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|
|**vbString**|8|[String](../../Glossary/vbe-glossary.md#string-data-type)|
|**vbObject**|9|A (VBA) [object](../../glossary/vbe-glossary.md#object) with a particular interface chosen, where the chosen interface directly exposes COM's **IDispatch** interface. [&dagger;](#singledagger "That the object implements IDispatch means that the VBA COM object that encompasses the passed object reference, is an (OLE) Automation object. That the chosen interface exposes IDispatch means that the particular object reference that has been passed, can be directly used with (OLE) Automation technology. If this constant is returned, it is possible to cast the argument to the Object data type.")|
|**vbError**|10|Has either of the following forms:<br><table><tr><td>i) An [**Error**](../../reference/user-interface-help/cverr-function.md) value.</td></tr><tr><td>ii) The [parameter](../../glossary/vbe-glossary.md#parameter) for a [_missing_](../../reference/user-interface-help/ismissing-function.md) [_optional_](../../concepts/getting-started/understanding-named-arguments-and-optional-arguments.md) **Variant** argument of some procedure, that hasn't yet had a conventional value assignment (the "missing" flag bit will have been set), or a variable holding the value of such a parameter. At the time of writing, such values are also **Error** values of the [error number 448](../../reference/user-interface-help/named-argument-not-found-error-448.md).</td></tr></table>[\*](#asteriskfootnote "Variant special value.")|
|**vbBoolean**|11|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|
|**vbVariant**|12|**Variant** <sup>_(used for return value only when added to **vbArray** constant)_</sup>|
|**vbDataObject**|13|A (VBA) object not represented by the **vbObject** constant documented in this table. [&Dagger;](#doubledaggerfootnote "Such an object, like all VBA objects, is still a COM object. Like all COM objects and interfaces, such an object exposes COM's IUnknown interface. Not to be confused with ActiveX Data Objects (ADO) which is a database technology.")|
|**vbDecimal**|14|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|
|**vbByte**|17|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) integer|
|**vbLongLong**|20|[LongLong](longlong-data-type.md) integer <sup>_(valid on 64-bit platforms only)_</sup>|
|**vbUserDefinedType**|36|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type)|
|**vbArray**|8192|[Array](../../Glossary/vbe-glossary.md#array) <sup>_(always added to another constant when returned by this function)_</sup>|


<table>
 <tr><td><a name="asteriskfootnote"><sup>*</sup></a></td><td>
  
  **Variant** special value.</td></tr>
 <tr><td><a name="singledagger"><sup>&dagger;</sup></a></td><td>
 
 That the object implements **IDispatch** means that the VBA COM object that encompasses the passed object reference, is an (OLE) [Automation object](../../Glossary/vbe-glossary.md#automation-object-1). That the **chosen** interface exposes **IDispatch** means that the particular object reference that has been passed, can be directly used with (OLE) Automation technology. <sup>[&sect;](#sectionfootnote)</sup><BR>If this constant is returned, it is possible to cast the argument to the [**Object**](../../reference/user-interface-help/object-data-type.md) data type.</td></tr>
 <tr><td><a name="doubledaggerfootnote"><sup>&Dagger;</sup></a></td><td>
 
 Such an object, like all VBA objects, is still a COM object. Like all COM objects and interfaces, such an object exposes COM's **IUnknown** interface. Not to be confused with [ActiveX Data Objects (ADO)](../../../access/concepts/activex-data-objects/set-properties-of-activex-data-objects-in-visual-basic.md) which is a database technology. <sup>[&sect;](#sectionfootnote)</sup></td></tr>
 <tr><td><a name="sectionfootnote"><sup>&sect;</sup></a></td><td>
 
 The glossary definition for [ActiveX object](../../Glossary/vbe-glossary.md#activex-object) in the VBA documentation on 7th April 2019 (current date), indicates that ActiveX objects are Automation objects. However, various developers instead use ActiveX as a synonym for the COM technology, meaning that those developers also class non-Automation COM objects as being a certain type of ActiveX object.</td></tr>
</table>

> [!NOTE] 
> These [constants](../../Glossary/vbe-glossary.md#constant) are specified by Visual Basic for Applications. The names can be used anywhere in your code in place of the actual values.

## Remarks

If an object represented by the **vbObject** constant (constant documented in the above table) is passed, and has a parameterless default member (either a default property or default function), **VarType**(_object_) returns a value indicating the type of the default member's return value in the case that there is a return value, and the value of the **vbEmpty** constant when there is no return value. If an object is passed that doesn't fulfill this criteria, the constant **vbObject** or the constant **vbDataObject** is returned, the constant representing the object type.

When passing the default object reference for objects/instances of [classes](../../Glossary/vbe-glossary.md#class) defined through [class modules](../../Glossary/vbe-glossary.md#class-module), **VarType** returns **vbObject** - this means such references directly support (OLE) Automation. Such references use the default interface of the respective object.

The **VarType** function never returns the value for **vbArray** by itself. It is always added to some other value to indicate an array of a particular type. For example, the value returned for an array of integers is calculated as **vbInteger** + **vbArray**, or 8194. 

The constant **vbVariant** is only returned in conjunction with **vbArray** to indicate that the argument to the **VarType** function is an array of type **Variant**.

When the function's argument evaluates to a **Variant** special value, the constant associated with the special value is returned.

## Example

This example uses the **VarType** function to determine: the subtypes of different **Variant** variables; the type of a particular non-**Variant** object variable; the type of an object's default member's return value; and the **Variant** special values that certain variables hold.

```vb
Dim MyCheck
Dim IntVar, StrVar, DateVar, AppVar, WorkbookVar

Dim ArrayVar
Dim UninitVar
Dim NullVar: NullVar = Null            ' Assign Null value.
IntVar = 459: StrVar = "Hello World": DateVar = #2/12/1969#
Set AppVar = Excel.Application
Set WorkbookVar = ActiveWorkbook       ' Workbook object.

' `stdole` is a library reference to the OLE Automation library.
' The IUnknown interface is the most basic COM interface & is 
' implemented by all COM objects, coming first before all other 
' interfaces in the interface order of a COM object. It was used before 
' (OLE) Automation was available.
Dim IUnknownVar As stdole.IUnknown
Dim ObjectVarWithInterfaceNotCompatibleWithAutomation

' Casting the Workbook object to an IUnknown object means that the 
' object reference is changed such that a different interface is chosen
' (fundamentally it is still the same object.)
Set IUnknownVar = WorkbookVar
Set ObjectVarWithInterfaceNotCompatibleWithAutomation = IUnknownVar

ArrayVar = Array("1st Element", "2nd Element")

' Run VarType function on different types.
MyCheck = varType(IntVar)              ' Returns 2.
MyCheck = varType(DateVar)             ' Returns 7.
MyCheck = varType(StrVar)              ' Returns 8.

' Assuming 'Microsoft Excel 16.0 Object Library' reference is being 
' used, return values for AppVar & WorkbookVar are as follows.
MyCheck = varType(AppVar)              ' Returns 8 (vbString) even 
                                       ' though AppVar is an object.
MyCheck = varType(WorkbookVar)         ' Returns 9 (vbObject) because 
                                       ' it's an object without a
                                       ' default member, & because the
                                       ' interface chosen is Automation 
                                       ' compatible.

MyCheck = varType(ObjectVarWithInterfaceNotCompatibleWithAutomation)
                                       ' Returns 13 (vbDataObject) even
                                       ' though object when considered
                                       ' as the broader object that
                                       ' encompasses this object reference, 
                                       ' does actually support Automation
                                       ' via the Workbook interface.
                                       
MyCheck = varType(IUnknownVar)         ' Returns 13 (vbDataObject)
                                       ' in respect of a non-Variant 
                                       ' variable.
                                       
MyCheck = varType(ArrayVar)            ' Returns 8204 which is
                                       ' `8192 + 12`, the computation of
                                       ' `vbArray + vbVariant`.

' Run VarType function on Variant special values.
MyCheck = varType(UninitVar) ' Returns 0 (vbEmpty).
MyCheck = varType(NullVar)   ' Returns 1 (vbNull).
```

## See also

- [VarType constants](../../Concepts/Getting-Started/vartype-constants.md)
- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
