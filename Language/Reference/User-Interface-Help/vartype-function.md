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

Returns an [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type) where the returned value will indicate one of the following things, the choice of which depends upon the [argument](../../Glossary/vbe-glossary.md#argument) passed:
1) The subtype or type of a [variable](../../Glossary/vbe-glossary.md#variable), [property](../../glossary/vbe-glossary.md#property), [expression](../../glossary/vbe-glossary.md#expression), [constant](../../Glossary/vbe-glossary.md#constant) or literal.
2) The type or lack of type for an [object](../../glossary/vbe-glossary.md#object)'s default member's return value.
3) The [**Variant**](../../Glossary/vbe-glossary.md#variant-data-type) special value that a **Variant** variable, property, expression or literal, evaluates to.

## Syntax

**VarType**(_arg_)

The required _arg_ argument must be of the **Variant** type, or be able to be [coerced](../../Reference/User-Interface-Help/data-type-summary.md#implicit-conversions--casts) to it.
 
## Return values

Return value is either:

- just one of the following constants excluding the **vbArray** constant and the **vbVariant** constant, _or_
- the **vbArray** constant added to any of the other constants from the following list.

|Constant|Value|Description|
|:-----|-----:|:-----|
|**vbEmpty**|0|[**Empty**](../../Glossary/vbe-glossary.md#empty) value <sup>_(represents uninitialized variable)_ [\*](#asteriskfootnote "Variant special value.")</sup>|
|**vbNull**|1|[**Null**](../../Glossary/vbe-glossary.md#null) value <sup>_(represents no valid data)_ [\*](#asteriskfootnote  "Variant special value.")</sup>|
|**vbInteger**|2|Integer of data type [**Integer**](../../Glossary/vbe-glossary.md#integer-data-type)|
|**vbLong**|3|[Long](../../Glossary/vbe-glossary.md#long-data-type) integer|
|**vbSingle**|4|[Single](../../Glossary/vbe-glossary.md#single-data-type) value <sup>_(single-precision floating-point number)_</sup>|
|**vbDouble**|5|[Double](../../Glossary/vbe-glossary.md#double-data-type) value <sup>_(double-precision floating-point number)_</sup>|
|**vbCurrency**|6|[Currency](../../Glossary/vbe-glossary.md#currency-data-type) value|
|**vbDate**|7|[Date](../../Glossary/vbe-glossary.md#date-data-type) value|
|**vbString**|8|[String](../../Glossary/vbe-glossary.md#string-data-type)|
|**vbObject**|9|Has either of the following forms:<br><ol type="i"><li>A (VBA) [object](../../glossary/vbe-glossary.md#object)-based type with a particular [_interface_](../../Glossary/vbe-glossary.md#interface) chosen[<sup>&Vert;</sup>](#doubleverticalbarfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, and that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, and the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts."), where the chosen interface directly exposes COM's **IDispatch** interface. <sup>[&dagger;](#singledagger "That the chosen interface exposes IDispatch means that the particular object-based type of the argument that has been passed, can be directly used with (OLE) Automation late-binding technology. If this constant is returned, it is possible to cast the argument to the Object data type, or the argument already has the Object data type.")</sup></li><li>[**Nothing**](../../reference/user-interface-help/nothing-keyword.md) value <sup>_(special value)_</sup> as a literal.</sup></li></ol>|
|**vbError**|10|Has either of the following forms:<br><ol type="i"><li>An [**Error**](../../reference/user-interface-help/cverr-function.md) value.</li><li>The [parameter](../../glossary/vbe-glossary.md#parameter) for a [_missing_](../../reference/user-interface-help/ismissing-function.md) [_optional_](../../concepts/getting-started/understanding-named-arguments-and-optional-arguments.md) **Variant** argument of some procedure, that hasn't yet had a conventional value assignment (the "missing" flag bit will have been set), or a variable or property holding the value of such a parameter. At the time of writing, such values are also **Error** values of the [error number 448](../../reference/user-interface-help/named-argument-not-found-error-448.md).</li></ol><sup>[\*](#asteriskfootnote "Variant special value.")</sup>|
|**vbBoolean**|11|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value|
|**vbVariant**|12|**Variant** <sup>_(used for return value only when added to **vbArray** constant)_</sup>|
|**vbDataObject**|13|A (VBA) object-based type with a particular _interface_ chosen[<sup>&Vert;</sup>](#doubleverticalbarfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, and that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, and the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts."), that is not represented by the **vbObject** constant documented in this table. <sup>[&Dagger;](#doubledaggerfootnote "An object of such an object-based type, like all VBA objects, is still a COM object. Like all COM objects and interfaces, such objects expose COM's IUnknown interface. Not to be confused with ActiveX Data Objects (ADO) which is a database technology.")</sup>|
|**vbDecimal**|14|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) value|
|**vbByte**|17|[Byte](../../Glossary/vbe-glossary.md#byte-data-type) integer|
|**vbLongLong**|20|[LongLong](longlong-data-type.md) integer <sup>_(valid on 64-bit platforms only)_</sup>|
|**vbUserDefinedType**|36|A value of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type)|
|**vbArray**|8192|[Array](../../Glossary/vbe-glossary.md#array) <sup>_(always added to another constant when returned by this function)_</sup>|


<table>
 <tr><td><a name="asteriskfootnote"><sup>*</sup></a></td><td>
  
  **Variant** special value.</td></tr>
 <tr><td><a name="singledagger"><sup>&dagger;</sup></a></td><td>
 
That the **chosen** _interface_[<sup>&Vert;</sup>](#doubleverticalbarfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, and that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, and the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts.") exposes **IDispatch** means that the particular object-based type of the argument that has been passed, can be directly used with (OLE) Automation late-binding technology. <sup>[&sect;](#sectionfootnote)</sup><BR>If this constant is returned, it is possible to [cast](../../Reference/User-Interface-Help/data-type-summary.md#implicit-conversions--casts) the argument to the [**Object**](../../reference/user-interface-help/object-data-type.md) data type, or the argument already has the **Object** data type.</td></tr>
 <tr><td><a name="doubledaggerfootnote"><sup>&Dagger;</sup></a></td><td>
 
 An object of such an object-based type, like all VBA objects, is still a COM object. Like all COM objects and _interfaces_, such objects expose COM's **IUnknown** interface. Not to be confused with [ActiveX Data Objects (ADO)](../../../access/concepts/activex-data-objects/set-properties-of-activex-data-objects-in-visual-basic.md) which is a database technology. <sup>[&sect;](#sectionfootnote)</sup></td></tr>
 <tr><td><a name="sectionfootnote"><sup>&sect;</sup></a></td><td>
 
 The glossary definition for [ActiveX object](../../Glossary/vbe-glossary.md#activex-object) in the VBA documentation on 7th April 2019 (current date), indicates that ActiveX objects are (OLE) Automation objects. However, various developers instead use ActiveX as a synonym for the COM technology, meaning that those developers also class non-OLE-Automation COM objects as being a certain type of ActiveX object.</td></tr>
</table>

> [!NOTE] 
> These [constants](../../Glossary/vbe-glossary.md#constant) are specified by Visual Basic for Applications. The names can be used anywhere in your code in place of the actual values.

## Remarks

#### Objects

If an object with a particular _interface_ chosen[<sup>&Vert;</sup>](#doubleverticalbarfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, and that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, and the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts."), represented by the **vbObject** constant (constant documented in the above table) is passed, and has a parameterless default member (either a default property or default function), **VarType**(_object_) returns a value indicating the type of the default member's return value in the case that there is a return value, and the value of the **vbEmpty** constant when there is no return value. If an object with a particular interface chosen, is passed that doesn't fulfill this criteria, the constant **vbObject** or the constant **vbDataObject** is returned, the constant representing the object type.

When passing data of an [object type](../../Glossary/vbe-glossary.md#object-type) corresponding to a [class](../../Glossary/vbe-glossary.md#class) defined through a [class module](../../Glossary/vbe-glossary.md#class-module), **VarType** returns **vbObject**&mdash;this means that when such data takes the form of an object reference referring to an actual object, the interface of the reference will directly support (OLE) Automation late-binding technology. Such references use the default interface[<sup>&Vert;</sup>](#doubleverticalbarfootnote "Chosen interface is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being implemented by the class (object's interfaces), that is chosen during run-time, and that has to be chosen before conventional execution of any of the object's methods or conventional access of any of the object's properties, can take place. If an interface needs to be chosen for an object, and the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of implicit type casts.") of the respective COM object.

#### Arrays

The **VarType** function never returns the value for **vbArray** by itself. It is always added to some other value to indicate an array of a particular type. For example, the value returned for an array of integers is calculated as **vbInteger** + **vbArray**, or 8194. 

#### Variant data

The constant **vbVariant** is only returned in conjunction with **vbArray** to indicate that the argument to the **VarType** function is an array whose element type is the **Variant** type.

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
' The IUnknown interface is the most basic COM interface and is 
' implemented by all COM objects, coming first before all other 
' interfaces in the interface order of a COM object. It was used before 
' (OLE) Automation was available.
Dim IUnknownVar As stdole.IUnknown
Dim ObjectVarWithNonIDispatchExposingInterfaceChosen

' Casting the Workbook object to an IUnknown object means that the 
' object reference is changed such that a different interface is chosen
' (fundamentally it is still the same COM object.)
Set IUnknownVar = WorkbookVar
Set ObjectVarWithNonIDispatchExposingInterfaceChosen = IUnknownVar

ArrayVar = Array("1st Element", "2nd Element")

' Run VarType function on different types.
MyCheck = varType(IntVar)              ' Returns 2.
MyCheck = varType(DateVar)             ' Returns 7.
MyCheck = varType(StrVar)              ' Returns 8.

' Assuming 'Microsoft Excel 16.0 Object Library' reference is being 
' used, return values for AppVar and WorkbookVar are as follows.
MyCheck = varType(AppVar)              ' Returns 8 (vbString) even 
                                       ' though AppVar is an object.
MyCheck = varType(WorkbookVar)         ' Returns 9 (vbObject) because 
                                       ' it's an object without a
                                       ' default member, and because the
                                       ' chosen interface exposes
                                       ' COM's IDispatch interface
                                       ' that is used in (OLE) 
                                       ' Automation.

MyCheck = varType(ObjectVarWithNonIDispatchExposingInterfaceChosen)
                                       ' Returns 13 (vbDataObject) even
                                       ' though object when considered
                                       ' as the broader COM object that
                                       ' encompasses this object reference, 
                                       ' can be referenced as a vbObject
                                       ' object via the Workbook interface.
                                       
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
<br><br><hr>

<a name="doubleverticalbarfootnote"><sup>&Vert;</sup></a> Chosen _interface_ is meant to refer to a defined interface, of the defined interfaces referenced in a COM object's class definition as being _implemented_ by the class (object's interfaces), that is chosen during run-time, and that has to be chosen before conventional execution of any of the object's [methods](../../glossary/vbe-glossary.md#method) or conventional access of any of the object's [properties](../../glossary/vbe-glossary.md#property), can take place. If an interface needs to be chosen for an object, and the interface isn't the object's default interface, it must be chosen by programmatically selecting (choosing) it. Programmatic interface selection is fundamentally supported by the VBA language grammar through the use of [implicit type casts](../../Reference/User-Interface-Help/data-type-summary.md#implicit-conversions--casts).

## See also

- [VarType constants](../../Concepts/Getting-Started/vartype-constants.md)
- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
