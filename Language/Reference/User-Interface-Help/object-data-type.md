---
title: Object data type
keywords: vblr6.chm1008829
f1_keywords:
- vblr6.chm1008829
ms.prod: office
ms.assetid: cffe448d-29dd-52aa-4a5c-2155c07b5bf3
ms.date: 11/19/2018
localization_priority: Normal
---


# Object data type

[Variables](../../Glossary/vbe-glossary.md#variable) of the [**Object** data type](../../Glossary/vbe-glossary.md#object-data-type) are stored as 32-bit (4-byte) addresses that refer to [Automation objects](../../Glossary/vbe-glossary.md#automation-object). Using the **Set** statement, a variable declared as having the **Object** type, can have any object reference assigned to it, as long as either:
- the reference's chosen interface is directly Automation compatible, _or_
- the reference's chosen interface has an Automation-compatible (inheriting) child interface that is also selected in the object reference <sup>[*](#asteriskfootnote "The child interface would not be such that it has the primary focus of the passed object reference. This happens, for example, if you assign a Collection object reference to an IUnknown variable, and then try to assign the IUnknown variable's value to a variable of the Object type. ...")</sup>.

> [!NOTE] 
> Although a variable declared with the **Object** type is flexible enough to contain a reference to any Automation object, binding to the object referenced by that variable is always late ([run-time](../../Glossary/vbe-glossary.md#run-time) binding). 
> 
> To force early binding ([compile-time](../../Glossary/vbe-glossary.md#compile-time) binding), assign the object reference to a variable declared with a specific [class](../../Glossary/vbe-glossary.md#class) name.

<a name="asteriskfootnote"><sup>*</sup></a> The child interface would not be such that it has the primary focus of the passed object reference. This happens, for example, if you assign a **Collection** object reference to an **IUnknown** variable, and then try to assign the **IUnknown** variable's value to a variable of the **Object** type. The unadorned **IUnknown** type isn't Automation compatible however, the **IUnknown** variable's value also has the child interface for the **Collection** class somehow also selected in the **IUnknown** variable's value (although it doesn't have primary focus.) VBA performs a kind of 'downcasting' of interfaces, and then assigns the 'downcast' reference, which is directly Automation compatible, to the **Object** variable.

## See also

- [Data type summary](data-type-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
