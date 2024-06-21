## SAP Automation Framework - Python

This framework simplifies and streamlines the development of automations for SAP GUI using Python. It's designed to enhance code readability, documentation, and maintainability.

### Key Features:

- **Object-Oriented Structure:** Utilizes a class-based approach (`SAP` class) for better organization and abstraction.
- **Simplified Interactions:** Provides easy-to-use methods for common SAP GUI actions:
    - Connecting and interacting with SAP sessions
    - Navigating menus, tabs, and fields
    - Writing and reading data from fields, tables, and grids
    - Handling variants, buttons, and selections
    - Saving files and retrieving session information
- **Robust Error Handling:** Incorporates error handling mechanisms and informative messages.
- **Language Support:** Leverages a `Language` class (refer to `language_dict.py`) for potential multi-language support.

### Example Usage:

The provided code snippet demonstrates a basic example of using the framework to interact with an SAP session. 

```python
from sap_functions import SAP

default_language = 'PT'
login = open('sap_login.txt', 'r').readline().strip().split(',')
scheduled_execution = {'scheduled?': False, 'username': login[0], 'password': login[1], 'principal': '100'}
sap_window = 0

class Work:
    def __init__(self):
        self.sap = SAP(sap_window, scheduled_execution, default_language)
        self.orders = []

    def COHV(self):
        self.sap.select_transaction('COHV')
        self.sap.write_text_field('Layout', '/usin_exce')
        self.sap.write_text_field('Elemento PEP', '120-2300923-21')
        self.sap.run_actual_transaction()

        my_grid = self.sap.get_my_grid()
        rows = self.sap.get_my_grid_count_rows(my_grid)
        for i in range(rows):
            self.orders.append(my_grid.getCellValue(i, 'AUFNR'))


if __name__ == '__main__':
    work = Work()
    work.COHV()
    print(work.orders)
```

This example outlines the initialization of the framework and the execution of actions within a `Work` class, which likely utilizes the `SAP` class for interacting with SAP GUI.

### How to Use:

1. **Prerequisites:**
   - Python 3.x
   - `win32com` library (install via `pip install pywin32`)
   - SAP GUI for Windows installed and configured
2. **Project Setup:**
   - Copy the provided code into your Python script (e.g., `sap_functions.py`).
   - Create a separate script (e.g., `main.py`) to utilize the `SAP` class.
3. **Customization:**
   - Adapt the `language_dict.py` file for your desired language translations.
   - Extend or modify the `SAP` class with additional methods or functionalities.
4. **Implementation:**
   - Import the `SAP` class into your main script.
   - Create an instance of the `SAP` class, providing the necessary parameters (window index, scheduled execution details, language).
   - Utilize the methods of the `SAP` object to automate your desired SAP processes.

### Notes:

- This framework relies on the SAP GUI Scripting API provided by `win32com`.
- Ensure that scripting is enabled in your SAP GUI settings.
- Refer to the SAP Scripting documentation for comprehensive information on available objects, methods, and properties.
- The `language_dict.py` file and its usage are not fully illustrated in the provided code; you might need to adapt or extend this aspect.