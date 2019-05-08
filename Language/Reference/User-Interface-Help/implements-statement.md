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

Specifies an [interface](../../Glossary/vbe-glossary.md#interface) that will be implemented in the [class module](../../Glossary/vbe-glossary.md#class-module) in which it appears.

## Syntax

**Implements** { _InterfaceName_ | _ClassName_ }

The required _InterfaceName_ or _ClassName_ is the name of an interface, or of a class from which an interface is automatically derived, where the interface's **Public** prototypes for determining [methods](../../Glossary/vbe-glossary.md#method) & [properties](../../Glossary/vbe-glossary.md#property), will be implemented by the corresponding methods in the Visual Basic class module in which the **Implements** statement has been used.

## Remarks

When a Visual Basic class implements an interface, the implementing Visual Basic class provides its own versions of all the methods & properties specified by the **Public** prototypes of the interface. An interface automatically derived from a class, in the context of the **Implements** statement, is simply an interface whose prototypes correspond to all the **Public** methods & properties of the referenced class. In addition to providing a mapping between the prototypes of the interface and the procedures of the implementing class module, the **Implements** statement causes instances of the class represented by the class module, to accept COM QueryInterface calls for the interface ID of the implemented interface.

> [!NOTE]
> In respect to object-oriented programming (OOP), Visual Basic for Applications (VBA) doesn't support implementation inheritance (in accordance with COM's specification). Interface inheritance including multiple interface inheritance can, of a sort, be mimicked using the VBA language through the use of the **Implements** statement (although strictly speaking it is interface implementation rather than interface inheritance). Even though multi-level interface inheritance is supported within type libraries, VBA only recognises type-library interfaces that directly inherit from **IDispatch** or **IUnknown**. This, coupled with the limitations of the **Implements** statement, renders multi-level interface inheritance support within VBA not so strong.

When implementing an interface where the implementation is enforced using the **Implements** statement, you must code for all the **Public** interface prototypes that determine properties & methods. For each **Public** property prototype of the interface, you need to implement both a [**Property Get** method](../../reference/user-interface-help/property-get-statement.md) & an appropriate value-setting method (either a [**Property Let** method](../../reference/user-interface-help/property-let-statement.md) or a [**Property Set** method](../../reference/user-interface-help/property-set-statement.md)) for the related property. For **Public** non-property method prototypes of the interface, you simply code functions for the **Public** function prototypes of the interface, & sub procedures for the **Public** sub-procedure prototypes of the interface, in the implementation.

The name of each implementing member needs to be text composed of the following three parts in the order specified: _interface name_; an underscore character ('\_'); _name of corresponding prototype member being implemented_. The parameters & return value of each implementing member must match those of the corresponding interface prototype member being implemented.

A missing member in an interface implementation causes an error. If you don't place code in one of the implementing procedures, you can raise the appropriate error (**Const** E_NOTIMPL = &H80004001) so a user of the implementation understands that a member is not properly implemented.

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

Also note that the **Customer** and **Supplier** classes, as defined earlier, do not expose the members of the PersonalData interface. The only way to access the PersonalData members is to assign a **Customer** or **Supplier** object to a variable declared as _PersonalData_. If an inheritance-like behavior is desired, with the **Customer** or **Supplier** class exposing the PersonalData members, **Public** members must be added to the class. These can be implemented by delegating to the PersonalData interface implementations. 

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
