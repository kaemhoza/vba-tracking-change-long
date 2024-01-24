# vba-tracking-change-long
This VBA code is designed to monitor changes in specific columns of an Excel worksheet and log these changes, along with additional information, into another worksheet. Here's a detailed description of how this code operates:

    Declaration of OldValue Variable:
        Dim OldValue As Variant: This line declares a variable named OldValue with a data type of Variant. This variable is used to store the value of a cell before it is changed.

    Worksheet_SelectionChange Event:
        This event is triggered every time a different cell or range of cells is selected in the worksheet.
        The code inside this event checks if the selected cell (Target) is within the specified columns (B, E, I, J, K, L, P, U) using the Application.Intersect method. If the selected cell is within these columns, the current value of the cell is stored in the OldValue variable.

    Worksheet_Change Event:
        This event occurs when cells on the worksheet are changed by the user or by an external link.
        A range object KeyCells is defined to represent the monitored columns (B, E, I, J, K, L, P, U).
        The Application.Intersect method is used again to check if the changed cell (Target) is within the KeyCells range. If it is, the code proceeds.
        A message box prompts the user to confirm if they wish to save the change made to the cell. It displays the address of the changed cell and provides Yes and No options.
        If the user chooses Yes, the code performs the following actions:
            It identifies the next available row in a designated archive sheet ("A2") to log the change.
            It sets a reference to another source sheet ("A1") to retrieve additional information.
            It assigns the value from column 5 of the source sheet ("A1") corresponding to the row of the changed cell to the variable SourceValue.
            In the archive sheet ("A2"), it logs the sequential number (calculated as the next row number minus one), the current date and time (Now), the SourceValue from column 5 of the source sheet, and the old value of the changed cell (OldValue).
            A message box confirms that the change has been logged.

This code is useful for tracking changes in specific columns of a worksheet and logging these changes, along with the date, time, and related information from another sheet, into an archive sheet for record-keeping or auditing purposes.
