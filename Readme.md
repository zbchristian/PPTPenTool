# Powerpoint Add-in: Pencil Toolbar
Add a hovering toolbar to the slide show to directly select the color of the annotation pen, an eraser, 
laser pointer and insert a new empty slide.
![Toolbar to select th epen color](https://github.com/zbchristian/PenTool/raw/master/images/Screenshot.png)

## How to Install
Copy the file `PenTool.ppam` file to your Add-in directory (e.g. `C:\Users\<username>\AppData\Roaming\Microsoft\AddIns` ). Open an empty Powerpoint presentation and 
goto  `File -> Options -> Add-Ins -> Manage "Powerpoint Add-Ins" -> Insert new` and select the file `PenTool.ppam`

It might be, that a security warning appears to enable macros. You will only be able to use the Add-in, when macros are enabled.

A new entry in the Menu appears called `Pen Tool`. 

When starting Powerpoint, the toolbar is disabled. To enable it, you need to click the `Init` button in the `Pen Tool` menu. 
The toolbar appears once the slide show is started.

## Customization
Open the file `PenTool.pptm` and start the VBA console (`ALT+F11`). 
After you did your modifications, save the pptm file AND do a ´save as to `PenTool.ppam`. The latter requires some additions. 
Install the `CustomUIEditor`, load the ppam fie and right click on the name in the left hand pane. Select `Office 2010 ...` and 
paste the content of the file `PenTool.xml` into the right pane. Customize the XML content.

Thats it! Now you can load the Add-In again. If you did not change the name, the modifications will be visible at the next start 
of Powerpoint.

Enjoy!
