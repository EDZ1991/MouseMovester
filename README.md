Dependencies:
- pyautogui: For controlling mouse movements and clicks.
- pywin32: For Windows-specific API calls to detect primary monitor dimensions
           and potentially query window information (though direct UI element
           detection is limited).

Usage:
1. Run `setup.bat` to install necessary Python packages.
2. Run `start_mouse_mover.bat` to start the program.
3. To stop the program, move the mouse cursor to the top-left corner of the screen
   (pyautogui's failsafe) or press Ctrl+C in the console.

Defensive Programming Notes:
- Error handling: `try-except` blocks are used for API calls and mouse operations.
- Failsafe: `pyautogui.FAILSAFE` is enabled for emergency program termination.
- Threading: Uses a separate thread for the mouse movement loop to keep the main
  program responsive.
- UI Avoidance (Best Effort): Implements a basic heuristic to avoid clicking
  in areas commonly occupied by window control buttons and the Start menu.
  This is not foolproof due to the dynamic nature of UI elements but aims
  to reduce unintended interactions.
