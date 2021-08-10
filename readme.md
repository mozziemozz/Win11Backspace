# Description

Missing the back button from Windows 10 Tablet Mode on the Taskbar in Windows 11?
Use this simple shortcut and PowerShell script to add a back button to your Windows 11 Taskbar.
It is primarily designed for small touchscreens and apps which have small back buttons. Having a back button on the taskbar makes it easier and faster to press it instead of using in app buttons.

![Taskbar](taskbar.png)

# How it works
The shortcut will launch a PowerShell script in minimized mode which does nothing more than just alt + tab to your last active Window and send emulate a backspace keystroke to that app.

# Installation
Place the folder `win11backspace` in your `C:\Temp` folder and right click the shortcut. Choose, `show more options` and then `pin to taskbar`.

## Customization
You may want to place the files in a different folder. Open the shortcuts properties and adjust the file path to your needs.

# Notes / Issues
This method only works with apps which support going back by backspace like Windows Explorer or Microsoft Photos. Based on my tests, it does not work in Edge. 