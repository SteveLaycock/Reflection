# Reflection (Work in progress)
A module that provides methods that extracts information from VBA source code at run time.
## How it works  
Reflection uses Application.VBE.vbProject.Item(\<name\>).vbComponent.Item(\<name\>).CodeModule to access the text for VBA code in Components that are available for editing at runtime.  Consequently, anything that can be achieved by 'interpreting' the raw text can be achieved by suitable VBA code.  

At some time in the future, Rubberduck, will be able to provide the information that is currently obtained by using the raw text in code modules.

## Methods
### IsFactory  

&nbsp;&nbsp;&nbsp;&nbsp;**Syntax**: IsFactory  
&nbsp;&nbsp;&nbsp;&nbsp;**Returns**: Boolean  

Returns 'True' if a Class contains the Rubberduck '@PredeclaredId annotation and a 'Make' method.  

### IsNotFactory  

&nbsp;&nbsp;&nbsp;&nbsp;**Syntax**: IsNotFactory  
&nbsp;&nbsp;&nbsp;&nbsp;**Returns**: Boolean  

Returns  'Not IsFactory ' 

### IsStatic  

&nbsp;&nbsp;&nbsp;&nbsp;**Syntax**: IsStatic  
&nbsp;&nbsp;&nbsp;&nbsp;**Returns**: Boolean  

Returns True if a Class contains the Rubberduck '@PredeclaredId annotation but does not contain a 'Make' method.

### IsNotFactory  

&nbsp;&nbsp;&nbsp;&nbsp;**Syntax**: IsNotFactory  
&nbsp;&nbsp;&nbsp;&nbsp;**Returns**: Boolean  

Determines Not IsStatic