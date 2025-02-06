Late Binding and Type Mismatches: VBScript's late binding can lead to runtime errors if an object's method or property doesn't exist or if a type mismatch occurs during an operation.  For example, attempting to access a property on a null object or performing arithmetic with incompatible data types can cause unexpected crashes or incorrect results.

Example:
```vbscript
Dim obj
Set obj = CreateObject("SomeObject.Class") 'Object might not exist
If obj.DoesNotExist = True Then
  MsgBox "This will cause an error if DoesNotExist doesn't exist"
end if
```