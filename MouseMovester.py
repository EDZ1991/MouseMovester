#!/usr/bin/env python3
"""
Random Mouse Mover for Windows 11

This program moves the mouse cursor to a random position on the primary monitor
every 5 seconds and performs a left-click. It is designed to run on Windows 11
and handles multi-monitor setups by ensuring the cursor stays within the bounds
of the primary display. It also includes basic logic to try and avoid clicking
on common UI elements like window control buttons (minimize, maximize, close)
and the Start menu area.

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
"""

import pyautogui  # Used for controlling mouse movements and clicks
import random     # Used for generating random coordinates
import time       # Used for delays and timing
import sys        # Used for system-specific parameters and functions (e.g., sys.exit)
import threading  # Used for running the mouse movement loop in a separate thread
from typing import Tuple # Used for type hinting

try:
    # Attempt to import win32api and win32con for Windows-specific monitor detection.
    # These modules are part of the pywin32 package.
    import win32api
    import win32con
    WINDOWS_AVAILABLE = True
except ImportError:
    # If pywin32 is not installed or not on Windows, set flag to False
    WINDOWS_AVAILABLE = False
    print("Warning: win32api not available. Will use pyautogui screen detection, which may not accurately detect primary monitor in multi-monitor setups.")


class RandomMouseMover:
    """A class to encapsulate the functionality of moving the mouse randomly and clicking."""

    def __init__(self):
        """Initializes the RandomMouseMover with default states and screen dimensions."""
        self.running = False  # Flag to control the main loop of the mouse mover
        self.thread = None    # Thread object for running the mouse movement loop
        
        # Enable pyautogui's FAILSAFE feature. Moving the mouse to the top-left
        # corner (0,0) of the screen will raise a pyautogui.FailSafeException
        # and stop the program, providing an emergency stop mechanism.
        pyautogui.FAILSAFE = True
        
        # Determine the dimensions of the primary monitor.
        # This is crucial for ensuring the mouse stays within a single screen.
        self.screen_width, self.screen_height = self.get_primary_monitor_size()
        print(f"Primary monitor size detected: {self.screen_width}x{self.screen_height}")
        
        # Define regions to avoid clicking. These are approximate and based on
        # typical Windows 11 UI layouts. These values might need adjustment
        # based on screen resolution, scaling, and taskbar position.
        # This is a best-effort attempt at defensive clicking.
        self.avoid_regions = []
        self._define_avoid_regions()

    def _define_avoid_regions(self):
        """
        Defines approximate screen regions to avoid clicking.
        These are heuristics and may not be perfect for all setups.
        Regions are defined as (left, top, right, bottom) pixel coordinates.
        """
        # Top-right corner for window control buttons (close, maximize, minimize)
        # Assuming a typical button size and padding.
        button_width = 50 # Approximate width of a control button
        top_bar_height = 30 # Approximate height of a window title bar
        
        # Avoid top-right corner (e.g., Close, Maximize, Minimize buttons)
        self.avoid_regions.append((
            self.screen_width - (3 * button_width), # Left: 3 buttons from right edge
            0,                                     # Top: from very top
            self.screen_width,                     # Right: to very right edge
            top_bar_height                         # Bottom: height of title bar
        ))

        # Bottom-left corner for Start Menu button and taskbar icons
        # Assuming taskbar is at the bottom and Start button is on the left.
        taskbar_height = 40 # Approximate height of the taskbar
        start_menu_width = 100 # Approximate width of the Start button area

        # Avoid bottom-left corner (e.g., Start button, search, task view)
        self.avoid_regions.append((
            0,                                     # Left: from very left edge
            self.screen_height - taskbar_height,  # Top: taskbar top edge
            start_menu_width,                      # Right: width of start menu area
            self.screen_height                     # Bottom: to very bottom edge
        ))
        
        # Avoid the entire taskbar area (assuming bottom taskbar)
        self.avoid_regions.append((
            0,                                     # Left: from very left edge
            self.screen_height - taskbar_height,  # Top: taskbar top edge
            self.screen_width,                     # Right: to very right edge
            self.screen_height                     # Bottom: to very bottom edge
        ))

        print("Defined avoidance regions:")
        for region in self.avoid_regions:
            print(f"  {region}")

    def is_position_in_avoid_region(self, x: int, y: int) -> bool:
        """
        Checks if a given (x, y) coordinate falls within any defined avoidance region.
        """
        for region in self.avoid_regions:
            left, top, right, bottom = region
            if left <= x <= right and top <= y <= bottom:
                return True
        return False

    def get_primary_monitor_size(self) -> Tuple[int, int]:
        """ 
        Determines and returns the width and height of the primary monitor.
        Prioritizes `win32api` for accurate multi-monitor detection on Windows.
        Falls back to `pyautogui.size()` if `win32api` is not available or fails.
        """
        if WINDOWS_AVAILABLE:
            try:
                # Get information about the primary monitor.
                # MonitorFromPoint((0,0)) returns the handle to the monitor
                # that contains the point (0,0), which is typically the primary monitor.
                primary_monitor = win32api.GetMonitorInfo(win32api.MonitorFromPoint((0, 0)))
                
                # The 'Monitor' key in the returned dictionary contains a tuple
                # (left, top, right, bottom) representing the monitor's coordinates.
                monitor_area = primary_monitor["Monitor"]
                width = monitor_area[2] - monitor_area[0]  # right - left
                height = monitor_area[3] - monitor_area[1] # bottom - top
                return width, height
            except Exception as e:
                print(f"Error getting primary monitor info via Windows API: {e}. Falling back to pyautogui.size().")
        
        # Fallback if win32api is not available or an error occurs.
        # pyautogui.size() returns the size of the entire virtual screen,
        # which spans across all monitors in a multi-monitor setup.
        return pyautogui.size()
    
    def get_random_position(self) -> Tuple[int, int]:
        """ 
        Generates a random (x, y) coordinate within the bounds of the primary monitor.
        A small margin is applied to prevent the cursor from going to the very edges,
        which can sometimes trigger unintended system behaviors or failsafe.
        """
        margin = 50  # Pixels from the edge to avoid
        
        # Generate random X coordinate within the screen width, respecting the margin.
        x = random.randint(margin, self.screen_width - margin)
        
        # Generate random Y coordinate within the screen height, respecting the margin.
        y = random.randint(margin, self.screen_height - margin)
        
        return x, y
    
    def move_mouse_and_click(self):
        """ 
        Moves the mouse cursor to a newly generated random position and performs a left-click.
        The movement is performed smoothly over a short duration. It attempts to avoid
        clicking in predefined UI sensitive regions.
        """
        try:
            x, y = self.get_random_position()
            
            # Check if the generated position is in an avoidance region
            if self.is_position_in_avoid_region(x, y):
                print(f"Skipping click at ({x}, {y}) as it's in an avoidance region.")
                return # Do not click if in an avoidance region

            # Move the mouse cursor to the generated (x, y) position.
            # The `duration` parameter makes the movement visible and smooth.
            pyautogui.moveTo(x, y, duration=0.5)
            
            # Perform a left-click at the new position
            pyautogui.click()
            
            current_time = time.strftime("%H:%M:%S")
            print(f"[{current_time}] Mouse moved to: ({x}, {y}) and clicked.")
            
        except Exception as e:
            print(f"Error moving mouse or clicking: {e}")
    
    def mouse_mover_loop(self):
        """ 
        The main loop that continuously moves the mouse cursor and clicks.
        It runs as long as the `self.running` flag is True.
        Includes error handling for graceful termination.
        """
        print("\nRandom mouse mover started. To stop: move mouse to top-left corner or press Ctrl+C.")
        
        while self.running:
            try:
                self.move_mouse_and_click()
                
                # Wait for 5 seconds before the next move.
                # The sleep is broken into smaller intervals to allow for quicker
                # response to the `self.running` flag change (e.g., when stopping).
                for _ in range(50):  # 50 iterations * 0.1 seconds/iteration = 5 seconds total
                    if not self.running:
                        break  # Exit the inner loop if stop is requested
                    time.sleep(0.1)
                    
            except pyautogui.FailSafeException:
                # This exception is raised when the mouse is moved to the top-left corner.
                print("\nEmergency stop triggered (mouse moved to top-left corner).")
                self.stop() # Call stop method to clean up and exit
                break      # Exit the main loop
            except KeyboardInterrupt:
                # This exception is raised when Ctrl+C is pressed in the console.
                print("\nKeyboard interrupt received.")
                self.stop() # Call stop method to clean up and exit
                break      # Exit the main loop
            except Exception as e:
                # Catch any other unexpected errors during execution.
                print(f"An unexpected error occurred: {e}")
                self.stop() # Attempt to stop gracefully
                break      # Exit the main loop
    
    def start(self):
        """ 
        Starts the mouse movement and clicking loop in a new daemon thread.
        A daemon thread runs in the background and automatically terminates
        when the main program exits.
        """
        if self.running:
            print("Mouse mover is already running.")
            return
        
        self.running = True
        # Create and start the thread. `daemon=True` ensures the thread exits
        # when the main program exits, preventing it from hanging.
        self.thread = threading.Thread(target=self.mouse_mover_loop, daemon=True)
        self.thread.start()
    
    def stop(self):
        """ 
        Stops the mouse movement and clicking loop.
        Sets the `self.running` flag to False, which signals the `mouse_mover_loop`
        to terminate. It then waits for the thread to finish.
        """
        self.running = False
        if self.thread and self.thread.is_alive():
            # Wait for the thread to complete its current operation and exit.
            # A timeout is provided to prevent indefinite waiting.
            self.thread.join(timeout=1)
        print("Random mouse mover stopped.")


def main():
    """ 
    The entry point of the program.
    Initializes the mouse mover and handles the main program flow,
    including initial checks and graceful shutdown.
    """
    print("=" * 50)
    print("Random Mouse Mover for Windows 11")
    print("=" * 50)
    
    # Basic check to ensure pyautogui is installed.
    # Although setup.bat handles this, it's good practice for direct execution.
    try:
        import pyautogui
    except ImportError:
        print("Error: The 'pyautogui' library is required but not found.")
        print("Please install it using: pip install pyautogui")
        sys.exit(1) # Exit the program if a critical dependency is missing
    
    # Create an instance of the RandomMouseMover class.
    mover = RandomMouseMover()
    
    try:
        # Start the mouse movement.
        mover.start()
        
        # Keep the main thread alive while the mouse mover thread runs.
        # This prevents the program from exiting immediately after starting the thread.
        while mover.running:
            time.sleep(1) # Sleep to reduce CPU usage while waiting
            
    except KeyboardInterrupt:
        # Handle Ctrl+C gracefully.
        print("\nCtrl+C detected. Stopping the mouse mover...")
        mover.stop()
    except Exception as e:
        # Catch any unexpected errors in the main thread.
        print(f"An unhandled error occurred in the main program: {e}")
        mover.stop()
    
    print("Program finished.")


if __name__ == "__main__":
    main()


