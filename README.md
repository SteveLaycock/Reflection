# Reflection (Work in progress)
A module that provides methods that extracts information from VBA source code at run time.

### IsFactory  

>**Syntax**: IsFactory  
>**Returns**: Boolean  

>Determines if a Class contains the Rubberduck '@PredeclaredId annotation and a 'Make' method.  

### IsNotFactory  

>**Syntax**: IsNotFactory  
>**Returns**: Boolean  

>Determines Not IsFactory  

### IsStatic  

>**Syntax**: IsStatic  
>**Returns**: Boolean  

>Determines if a Class contains the Rubberduck '@PredeclaredId annotation but does not contain a 'Make' method.

### IsNotFactory  

>**Syntax**: IsNotFactory  
>**Returns**: Boolean  

>Determines Not IsStatic