---
title: Implements statement (VBA)
keywords: vblr6.chm1103517
f1_keywords:
- vblr6.chm1103517
ms.prod: office
ms.assetid: 9d0fe592-e945-649a-d277-fe882fe8cf67
ms.date: 12/03/2018
localization_priority: Normal
---

# Implements statement

Specifies an [interface](../../Glossary/vbe-glossary.md#interface) or [class](../../Glossary/vbe-glossary.md#class) that will be implemented in the [class module](../../Glossary/vbe-glossary.md#class-module) in which it appears.

## Syntax

**Implements** { _InterfaceName_ | _ClassName_ }

The required _InterfaceName_ or _ClassName_ is the name of an interface or class whose public methods & public properties will be implemented by the corresponding methods in the Visual Basic class in which the **Implements** statement has been used.

## Remarks

When a Visual Basic class implements an interface or class, the implementing Visual Basic class provides its own versions of all the **Public** [procedures](../../Glossary/vbe-glossary.md#procedure) & [properties](../../Glossary/vbe-glossary.md#property) specified in the referenced interface or class. In addition to providing a mapping between the prototypes (from the referenced interface or class) and the procedures of the implementing class, the **Implements** statement causes the implementing class to accept COM QueryInterface calls for the specified interface ID of the implemented interface or referenced class.

> [!NOTE] 
> In respect to object-oriented programming (OOP), Visual Basic for Applications (VBA) doesn't support implementation inheritence & also doesn't support multilevel interface inheritence. VBA supports single-level interface inheritence that can be multiple single-level interface inheritence, through the use of the **Implements** statement.

When you implement an interface or class, you must code for all the **Public** properties & methods. For each **Public** property of the implemented interface or referenced class, you need to implement both a [**Property Get** method](../../reference/user-interface-help/property-get-statement.md) & an appropriate value-setting method (either a [**Property Let** method](../../reference/user-interface-help/property-let-statement.md) or a [**Property Set** method](../../reference/user-interface-help/property-set-statement.md)) for the property. For non-property methods, you simply code functions for functions & sub procedures for sub procedures.

The name of each implementing member needs to be text composed of the following three parts in the order specified: _name of the referenced interface or class_; an underscore character ('\_'); _implemented-member name_.

A missing member in an implementation of an interface or class causes an error. If you don't place code in one of the procedures in a class you are implementing, you can raise the appropriate error (**Const** E_NOTIMPL = &H80004001) so a user of the implementation understands that a member is not implemented.

The **Implements** statement can be used more than once in a class module, but can't be used in a [standard module](../../Glossary/vbe-glossary.md#standard-module).


## Example

The following example shows how to use the **Implements** statement to make a set of declarations available to multiple classes. By sharing the declarations through the **Implements** statement, neither class has to make any declarations itself. The example also shows how use of an interface allows abstraction: a strongly-type variable can be declared by using the interface type. It can then be assigned objects of different class types that implement the interface.

Assume there are two forms, SelectorForm and DataEntryForm. The selector form has two buttons, **Customer Data** and **Supplier Data**. To enter name and address information for a customer or a supplier, the user clicks the customer button or the supplier button on the selector form, and then enters the name and address by using the data entry form. The data entry form has two text fields, **Name** and **Address**.

The following code for the interface declarations is in a class called **PersonalData**:

```vb
Public Name As String 
Public Address As String 
```

<br/>

The code supporting the customer data is in a class module called **Customer**. Note that the PersonalData interface is implemented with members that are named with the interface name `PersonalData_` as a prefix.

```vb
Implements PersonalData

'For PersonalData implementation
Private m_name As String
Private m_address As String

'Customer specific
Public CustomerAgentId As Long

'PersonalData implementation
Private Property Let PersonalData_Name(ByVal RHS As String)
    m_name = RHS
End Property
 
Private Property Get PersonalData_Name() As String
    PersonalData_Name = m_name
End Property


Private Property Let PersonalData_Address(ByVal RHS As String)
    m_address = RHS
End Property

Private Property Get PersonalData_Address() As String
    PersonalData_Address = m_address
End Property


'Initialize members
Private Sub Class_Initialize()
    m_name = "[customer name]"
    m_address = "[customer address]"
    CustomerAgentID = 0
End Sub

```


<br/>

The code supporting the supplier data is in a class module called **Supplier**:

```vb
Implements PersonalData

'for PersonalData implementation
Private m_name As String
Private m_address As String

'Supplier specific
Public NumberOfProductLines As Long


'PersonalData implementation
Private Property Let PersonalData_Name(ByVal RHS As String)
    m_name = RHS
End Property
 Private Property Get PersonalData_Name() As String
    PersonalData_Name = m_name
End Property


Private Property Let PersonalData_Address(ByVal RHS As String)
    m_address = RHS
End Property

Private Property Get PersonalData_Address() As String
    PersonalData_Address = m_address
End Property


'initialize members
Private Sub Class_Initialize()
    m_name = "[supplier name]"
    m_address = "[supplier address]"
    NumberOfProductLines = 15
End Sub


```

<br/>

The following code supports the **Selector** form:

```vb
Private cust As New Customer 
Private sup As New Supplier 
 
Private Sub Customer_Click() 
Dim frm As New DataEntryForm 
 Set frm.PD = cust 
 frm.Show 1 
End Sub 
 
Private Sub Supplier_Click() 
Dim frm As New DataEntryForm
 Set frm.PD = sup 
 frm.Show 1 
End Sub
```

<br/>

The following code supports the **Data Entry** form:

```vb
Private m_pd As PersonalData

Private Sub SetTextFields()
    With m_pd
        Text1 = .Name
        Text2 = .Address
    End With
End Sub

Public Property Set PD(Data As PersonalData) 
    Set m_pd = Data
    SetTextFields
End Property

Private Sub Text1_Change()
    m_pd.Name = Text1.Text
End Sub

Private Sub Text2_Change()
    m_pd.Address = Text2.Text
End Sub

```

<br/>

Note how, in the data entry form, the *m_pd* variable is declared by using the PersonalData interface, and it can be assigned objects of either the **Customer** or **Supplier** class because both classes implement the PersonalData interface.

Also note that the *m_pd* variable can only access the members of the PersonalData interface. If a **Customer** object is assigned to it, the **Customer-specific member CustomerAgentId** is not available. Similarly, if a **Supplier** object is assigned to it, the Supplier-specific member **NumberOfProductLines** is not available. Assigning an object to variables declared by using different interfaces provides a polymorphic behavior.

Also note that the **Customer** and **Supplier** classes, as defined earlier, do not expose the members of the PersonalData interface. The only way to access the PersonalData members is to assign a **Customer** or **Supplier** object to a variable declared as _PersonalData_. If an inheritance-like behavior is desired, with the **Customer** or **Supplier** class exposing the PersonalData members, public members must be added to the class. These can be implemented by delegating to the PersonalData interface implementations. 

For example, the **Customer** class could be extended with the following:

```vb
'emulate PersonalData inheritance
Public Property Let Name(ByVal RHS As String)
    PersonalData_Name = RHS
End Property

Public Property Get Name() As String
    Name = PersonalData_Name
End Property

Public Property Let Address(ByVal RHS As String)
    PersonalData_Address = RHS
End Property

Public Property Get Address() As String
    Address = PersonalData_Address
End Property

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
