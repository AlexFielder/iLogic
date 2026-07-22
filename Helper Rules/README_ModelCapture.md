# Model-Independent iLogic Rules

This collection provides a solution for capturing all model information (parameters, features, operations) from any Inventor model without needing to know specific parameter names in advance.

## Files

### 1. `ModelIndependentCapture.iLogicVb`
The main rule that captures all model information and stores it in the requested array format.

**Key Features:**
- Works with Parts, Assemblies, and Drawings
- Captures all parameters regardless of name
- Stores data in 3-column array: [Operation, Name, Value]
- Provides helper functions for programmatic access

**Usage:**
```vb
' Simply run the rule - it will capture everything automatically
' Data is stored in array format as requested:
' Column 0: Operation/Feature type (Parameter, Feature, Sketch, etc.)
' Column 1: Name of the parameter/feature
' Column 2: Value of the parameter/feature
```

### 2. `ModelInfoCapture.iLogicVb`
Advanced version with additional features like data export, filtering, and detailed analysis.

**Key Features:**
- Complete model information capture
- Data filtering by category
- Export to CSV
- Search functionality
- Detailed reporting

### 3. `ModelCaptureExamples.iLogicVb`
Examples showing how to use the captured data programmatically.

**Examples Include:**
- Basic array access (as requested in the issue)
- Parameter searching without knowing names
- Generic parameter processing
- Model analysis and statistics

## How to Use

### Basic Usage (Addresses the Original Request)

1. Open any Inventor model (Part or Assembly)
2. Run the `ModelIndependentCapture` rule
3. The rule will automatically:
   - Read all operations and parameters
   - Store them in a 3-column array [Operation, Name, Value]
   - Display a summary of captured information

### Programmatic Access

```vb
' Get the model information array
Dim modelArray(,) As String = GetModelInfoArray()

' Loop through all rows to access each parameter
For i As Integer = 1 To modelArray.GetLength(0) - 1
    Dim operation As String = modelArray(i, 0)  ' 1st column: operations
    Dim paramName As String = modelArray(i, 1)  ' 2nd column: name of parameter  
    Dim paramValue As String = modelArray(i, 2) ' 3rd column: value of parameter
    
    ' Now you can access each parameter without using the name given by the user
    If operation = "Parameter" Then
        ' Process parameter without hardcoding names
        MessageBox.Show($"Found: {paramName} = {paramValue}")
    End If
Next
```

### Finding Parameters by Pattern

```vb
' Search for parameters containing specific text
Dim modelArray(,) As String = GetModelInfoArray()

For i As Integer = 1 To modelArray.GetLength(0) - 1
    Dim operation As String = modelArray(i, 0)
    Dim paramName As String = modelArray(i, 1)
    Dim paramValue As String = modelArray(i, 2)
    
    If operation = "Parameter" AndAlso paramName.ToLower().Contains("length") Then
        ' Found a length parameter - can modify it without knowing the exact name
        Parameter(paramName) = CDbl(paramValue) * 1.1 ' Scale by 10%
    End If
Next
```

## Problem Solved

This solution addresses the original issue request:

> "How can I do with ilogic to create a rule that is model independent? When the model is opened and the rule is executed, it will automatically read the operations and parameters for me and start the rule. The idea I have in mind is that the names are stored in an array 1st column operations 2nd column name of the parameter and 3rd value of the parameter. So going through all the rows of the array I can access each parameter without having to use the name given by the user."

✅ **Model Independent**: Works with any Part or Assembly  
✅ **Automatic Reading**: Captures all operations and parameters when executed  
✅ **Array Format**: Stores data in exactly the requested 3-column format  
✅ **No Hardcoded Names**: Access parameters without knowing their names in advance  

## Installation

1. Copy the `.iLogicVb` files to your Inventor iLogic rules location
2. Open any Inventor model
3. Run `ModelIndependentCapture` to start capturing model information
4. Run `ModelCaptureExamples` to see usage examples

## Notes

- The rules handle errors gracefully and will continue processing even if some parameters can't be read
- Captured information is also stored in custom iProperties for reference
- The solution works with Inventor's standard parameter types (User, Model, Reference, Derived)
- For assemblies, it also captures component and constraint information