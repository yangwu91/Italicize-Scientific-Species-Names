# Bulk Italicize Scientific Names in Word using a VBA Macro

English | [中文](https://github.com/yangwu91/Italicize-Scientific-Species-Names/blob/main/README_zh.md)

## Step 1: Create and Save the VBA Macro

1. Open the Microsoft Word application.
2. Press `Alt + F11` to open the Visual Basic for Applications (VBA) Editor.
3. In the "Project" explorer pane on the left, find and double-click on the **`Normal`** (or `Normal.dotm`) project. This represents Word's global template.
   > **Important:** Modifying the `Normal` template ensures the macro is available in all your documents. If you select a project specific to an open document, the macro will only work in that single file.
4. Right-click on the `Normal` project, then select ​**Insert > Module**​. This will create a new code module.
5. Copy and paste the entire VBA code `ItalicizeSpecies.vba` into the code window that appears on the right.
6. **Customize your list of scientific names:**
   * In the code, locate the `speciesList = Array(...)` section. You can freely add, delete, or modify the names here.
   * Each name must be enclosed in double quotes (`""`).
   * Names must be separated by commas (`,`).
   * For long lists, you can use an underscore (`_`) to break a single line of code into multiple lines for better readability.
7. Press `Ctrl + S` to save the macro, then close the VBA Editor.

## Step 2: Add a Quick Access Button for the Macro

For easy access, you can add a button to the Quick Access Toolbar (the small icons at the very top-left of the Word window). This allows you to run the macro with a single click.

1. In the main Word window, click on ​**File > Options**​.
2. In the "Word Options" dialog box, select ​**Quick Access Toolbar**​.
3. From the "Choose commands from:" dropdown menu, select ​**Macros**​.
4. You should see the macro you just created, likely named something like `Normal.Module1.ItalicizeScientificNames`. Select it.
5. Click the **Add >>** button in the middle to move it to the list on the right.
6. **(Optional but Recommended) Modify the button icon and name:**
   * With the macro selected in the right-hand list, click the **Modify...** button below.
   * In the new window, you can choose a more intuitive icon (like an italic ​*I*​) and set a user-friendly "Display name," such as "Italicize Names."
   * Click **OK** to save the changes.
7. Finally, click **OK** to close the "Word Options" dialog box.

## Step 3: Start Using the Macro!

You will now see the new icon you set up in the Quick Access Toolbar at the top-left of your Word window. Whenever you need to format the scientific names in your document, simply click this button to run the macro.
