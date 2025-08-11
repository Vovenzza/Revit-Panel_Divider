# Revit Analytical Panel Divider

Autodesk Revit 2023 add-in designed to split analytical panels using other panels as cutting guides. This tool is perfect for structural engineers and BIM modelers who need to break down large analytical surfaces into smaller, constructible panels while preserving important data.

![https://github.com/Vovenzza/Revit-Panel_Divider/blob/956d967507b7dff7ec5459dcf7fd343481b779ad/Revit_NWifBwiRWs-ezgif.com-optimize.gif]
---

### What It Does

The script provides a simple workflow within Revit:

1.  **Select Panels to Split:** You first select the main analytical panels that you want to divide.
2.  **Select Cutting Panels:** You then select one or more "cutter" panels. The script uses the planes of these cutters to determine where to slice the main panels.
3.  **Automatic Splitting:** The script calculates the intersection lines and intelligently splits the original panels into new, smaller ones.
4.  **Data Preservation:** The original panel is deleted, and its properties (like material, thickness, and other parameters) and any analytical openings (like doors or windows) are automatically copied to the newly created panels.

### Key Features

*   **Intuitive Workflow:** Simple two-step selection process right inside the Revit UI.
*   **Multi-Panel Support:** Split multiple panels in a single operation.
*   **Parameter Copying:** Automatically transfers all writable parameters from the source panel to the new panels, ensuring data continuity.
*   **Opening Preservation:** Intelligently copies analytical openings to the new panels they fall within.
*   **Robust Geometry Engine:** Includes advanced logic to handle complex contours, clean up vertices, and ensure the created panels are valid.
*   **Detailed Logging:** Creates a log file in your temp folder (`RevitScriptLog.txt`) for easy debugging and troubleshooting.

### How to Use

1.  **Load the Add-in:** Compile the project and load the resulting `.dll` and `.addin` files into Revit.
2.  **Run the Command:** Find the "Divide Analytical Panels" command in the Revit Add-ins tab.
3.  **Select Panels:** Follow the on-screen prompts to first select the panels you want to split, and press `Enter`.
4.  **Select Cutters:** Select the panels that will act as cutting planes, and press `Enter`.
5.  **Done!** The script will process the panels inside a single transaction.

### AI-Assisted Development ("Vibecoding")

This script was developed with significant assistance from AI. The process, which could be described as **"Vibecoding,"** involved translating a clear functional vision and complex geometric logic into robust C# code through collaboration with an AI partner. The AI helped structure the code, implement complex mathematical functions, and debug the workflow, acting as a powerful tool to bring the initial idea to life.

### Installation & Setup

To use this script, you need to load it as a Revit Add-in.

1.  **Compile the Code:** Open the project in Visual Studio 2022 and build the solution. This will create a `PanelDivider.dll` file in the `bin/Debug` or `bin/Release` folder.
2.  **Create a Manifest File:** Create a file named `PanelDivider.addin` in Revit's add-in directory (`%appdata%\Autodesk\Revit\Addins\2023`). Paste the following into it, making sure the `<Assembly>` path points to your `.dll` file:

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
    <RevitAddIns>
      <AddIn Type="Command">
        <Name>Divide Analytical Panels</Name>
        <FullClassName>PanelDivider.DivideAnalyticalPanels</FullClassName>
        <Text>Divide Analytical Panels</Text>
        <Description>Splits analytical panels using other panels as cutting planes.</Description>
        <VisibilityMode>AlwaysVisible</VisibilityMode>
        <Assembly>"C:\Path\To\Your\Project\bin\Debug\PanelDivider.dll"</Assembly>
        <AddInId>{GENERATED-GUID-HERE}</AddInId>
        <VendorId>YourName</VendorId>
        <VendorDescription>Your Company or Website</VendorDescription>
      </AddIn>
    </RevitAddIns>
    ```
    *   **Important:** Replace `"C:\Path\To\Your\Project\bin\Debug\PanelDivider.dll"` with the actual path to your DLL.
    *   You should also generate a new GUID (you can do this in Visual Studio via `Tools > Create GUID`) and paste it into the `<AddInId>` field.

3.  **Run Revit:** Start Revit 2023, and the command should appear in your "Add-Ins" tab.

---

**Disclaimer:** This is a utility script. Always back up your Revit models before running commands that modify geometry. Use at your own risk.
