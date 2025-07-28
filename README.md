# 🎨 Excel-Style Color Picker for WinForms

A comprehensive, Excel-style color picker component for C# WinForms applications. This user control provides a modern and feature-rich interface for color selection, designed to be easily integrated as a drop-down from a button.

## ✨ Demo

 
*(This is a placeholder for a real GIF/screenshot of your control in action)*

## 🚀 Features

-   **🌈 Circular Color Wheel**: An intuitive wheel for selecting hue and saturation.
-   **🎚️ Brightness Slider**: A vertical slider to control the selected color's brightness (HSV Value).
-   **▦ Standard Palette**: A pre-defined grid of theme colors, similar to Microsoft Office applications.
-   **🔢 Live Value Display**: Instantly see the selected color's values in **HEX**, **RGB**, and **CMYK** formats.
-   **🛡️ Safety Modes**:
    -   **Web-Safe Mode**: Snaps all color selections to the nearest web-safe color.
    -   **Print-Safe Mode**: Snaps all selections to their nearest CMYK-safe equivalent, ensuring better color fidelity in print.
-   **⚠️ Safety Warnings**: Visual cues indicate if the current color is not web-safe or print-safe, with a one-click option to fix it.
-   **👀 Original vs. New Color Preview**: Easily compare the newly selected color with the original one.
-   **⌨️ Full Keyboard Navigation**: All elements of the control are fully accessible and navigable using the keyboard.
-   **🧩 Easy Integration**: Simply add the `ExcelColorPopupButton` to your form. No other setup is required.
-   **📦 Self-Contained**: The control is written entirely in C# and relies only on the standard .NET `System.Windows.Forms` and `System.Drawing` libraries. No external dependencies are needed.

## 🛠️ Getting Started

###📋 Prerequisites

You need a .NET project with Windows Forms support. This component has been designed with .NET Framework 4.x or .NET 5/6/7/8 and later in mind.

### ⚙️ Installation

1.  Clone this repository or download the source code.
2.  Add the C# source files (`ExcelColorPopupButton.cs`, `ExcelColorDropDownControl.cs`, etc.) to your WinForms project.
3.  Rebuild your project.
4.  The `ExcelColorPopupButton` control should now appear in the Visual Studio Toolbox, ready to be dragged onto your forms.

## 💡 How to Use

The easiest way to use the color picker is with the `ExcelColorPopupButton` control.

1.  **🖱️ Add the Button to Your Form**:
    -   After rebuilding your project, find the `ExcelColorPopupButton` in your Visual Studio Toolbox.
    -   Drag and drop it onto your form.

2.  **⚡ Handle the `ColorChanged` Event**:
    -   Select the button on your form and go to the Events tab in the Properties window.
    -   Double-click the `ColorChanged` event to create a handler.
    -   Use this event to get the newly selected color and apply it to other controls or objects in your application.

```csharp
// Example: Change a Panel's background color when a new color is selected.
private void excelColorPopupButton1_ColorChanged(object sender, EventArgs e)
{
    // Get the selected color from the button's property
    Color newColor = excelColorPopupButton1.SelectedColor;

    // Apply it to another control
    panelToChange.BackColor = newColor;

    // You can also get the value as a string for display
    this.Text = $"Selected Color: {newColor}";
}

// You can also set the initial color of the button in your Form's Load event
private void Form1_Load(object sender, EventArgs e)
{
    excelColorPopupButton1.SelectedColor = Color.CornflowerBlue;
}
```

## 🔍 Code Overview

The project is structured into several key classes:

-   **`🔘 ExcelColorPopupButton`**: The main `Button` control that you add to your form. It handles painting the color preview and the drop-down arrow, and it manages the display of the `ExcelColorDropDown`.

-   **`🔽 ExcelColorDropDown`**: A custom `ToolStripDropDown` that acts as a borderless window to host the main color picker user control.

-   **`🎛️ ExcelColorDropDownControl`**: The core `UserControl` that contains all the UI elements: the circular picker, brightness slider, color grid, text boxes, and checkboxes. It manages all the internal logic for color updates and state changes.

-   **`🌈 CircularColorPicker`**: A custom control that renders the HSL color wheel and handles user input (mouse and keyboard) for selecting hue and saturation.

-   **`🎚️ BrightnessSlider`**: A custom control for the vertical brightness (value) slider.

-   **`🔧 ColorUtils`**: A static helper class containing methods for color model conversions (RGB, HSV, HSL, CMYK) and color safety checks.

-   **`📝 VerticallyCenteredTextBox`**: A small enhancement of the standard `TextBox` to vertically center its text.

-   **`🖌️ DoubleBufferedPanel`**: A simple `Panel` with double-buffering enabled to prevent flicker in the color grid.

## 📜 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
