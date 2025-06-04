# Inspector

A Windows Forms application that leverages the Windows Automation API to inspect UI elements at runtime. By pressing a configurable hotkey (Ctrl+F by default), users can capture information about the control under the mouse cursor and display it within the Inspector Tree interface. The application shows details about the control (e.g., Name, ClassName, AutomationId) and generates code snippets (e.g., selectors for automated testing).

---

## Features

### Live UI Inspection
- Press `Ctrl+F` to capture the UI element under your cursor.
- Retrieves the top-level window and scans its child controls.
- Displays a hierarchical view in a TreeView; expand or collapse nodes to explore nested controls.

### Control Information
- Displays essential properties:  
    - Name  
    - ClassName  
    - AutomationId  
    - ControlType  
- Simple text-based output in the details panel.

### Generated Automation Selectors
- Suggests code lines (Python-like syntax) to locate the control by its properties.
- Useful for automated test scripts or custom UI interactions.

### Context Menu Actions
Right-click a node in the TreeView to perform:
- Click  
- Right Click  
- Double Click  
- Type Text  

These actions move the mouse programmatically and perform clicks or keystrokes on the control’s bounding rectangle.

---

## Requirements
- Windows OS  
- .NET 9.0 (Windows Desktop) or later

---

## Usage
1. Start the application.  
2. Press `Ctrl+F` anywhere on your desktop.  
3. The tree builds, showing the UI element at the cursor.  
4. Click a node to view details on the right.  
5. “Suggestions” panel shows code lines for selectors.  
6. Right-click a node for context actions.

---

## Searching
- Use the search bar in the top-left to filter nodes by partial matches.  
- Press **Enter** to jump to the first match.