---
title: = operator
keywords: vblr6.chm1009738
f1_keywords:
- vblr6.chm1009738
ms.prod: office
ms.assetid: d0140aaf-7475-97e4-da7d-630c3f562b30
ms.date: 11/19/2018
localization_priority: Normal
---


# = operator

Used to [assign](../../Concepts/Getting-Started/writing-assignment-statements.md) or [coerce](../../Reference/User-Interface-Help/data-type-summary.md#assignment-statements-implicit-conversions-and-casts) a value to a [variable](../../Glossary/vbe-glossary.md#variable) or [property](../../Glossary/vbe-glossary.md#property).

## Syntax

_variable_=_value_

The **=** operator syntax has these parts:

|Part|Description|
|:-----|:-----|
| _variable_|Variable or writable property; cannot be [array](../../glossary/vbe-glossary.md#array)-data-type variable; if variable, can be array element, [user-defined-type](../../glossary/vbe-glossary.md#user-defined-type) element, or standard variable; can only be variable of [object](../../glossary/vbe-glossary.md#object)-based type if variable holds object having a default writable property.|
| _value_|A numeric or string literal, [constant](../../Glossary/vbe-glossary.md#constant), [expression](../../Glossary/vbe-glossary.md#expression), variable or readable property.|

## Remarks

Properties on the left side of the equal sign can only be those properties that are writable at [run time](../../Glossary/vbe-glossary.md#run-time).

This operator is more fully documented under the [Let statement](../../Reference/User-Interface-Help/let-statement.md) documentation.

## See also

- [Operator summary](operator-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
