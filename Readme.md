# Powerpoint Add-in: Quick access pen toolbar for slideshows
Add a hovering quick access toolbar to the slide show to directly select the color of the annotation pen, a highlighter, an eraser, 
laser pointer, insert a new empty slide and display the slide overview. 

The toolbar can be used in duplicated screen view as well as on the presenter view.

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Screenshot.png" alt="Toolbar to select pen properties and more" width="500">

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Screenshot_vert.png" alt="Toolbar to select pen properties and more" width="150">

As you can see, there are two versions of the toolbar. At startup you will see the horizontal version. Clicking the turn left/right icon allows to switch between the two versions. 

The vertical version is, due to restrictions for the minimal width of `userforms` in office, quite wide.

**The toolbar has been tested with Office 2016 and 2019 with 64bit versions**

## How to Install
Test the Add-in by double clicking the the file `PenTool.ppam` and Powerpoint should start and run the Add-in. To install the toolbar, you should copy it to your Add-in directory (e.g. `C:\Users\<username>\AppData\Roaming\Microsoft\AddIns` ). Open an empty Powerpoint presentation and 
goto  `File -> Options -> Add-Ins -> Manage "Powerpoint Add-Ins" -> Insert new` and select the file `PenTool.ppam`

It might be, that a security warning appears to enable macros. You will only be able to use the Add-in, when macros are allowed to be executed.

A new entry in the Menu appears called `Pen Tool`. 

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Screenshot_Ribbon.png" alt="Pen Tool Ribbon to enable/disable the toolbar">

## How to use it?
When starting Powerpoint, the toolbar is disabled. To enable it, you need to click the `Enable` button in the `Pen Tool` menu. 
The toolbar appears once the slide show is started.

Annotations are usually done with a second screen/projector attached to the laptop or tablet. Per default, Powerpoint starts with the "Presenter View" on the main screen. The toolbar will appear on the "Presenter View" screen in the top left corner. You can use it to annotate the current slide in the presenter view. Usually you would increase the size of the current slide to make this convenient.

To draw directly on the projected screen, you need to switch to "Mirrored Screens" AND disable the "Presenter View" in the "Slide Show" settings. Goto Monitor group and uncheck "Use Presenter View".

### Buttons

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Move_256.bmp" width="30" alt="Move button"> : Move the toolbar to a different location

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Turn_right_256.bmp" width="30" alt="Turn button"> : Turn the toolbar to vertical layout

<img src="https://github.com/zbchristian/PenTool/raw/master/images/SelectColor.png" width="30" alt="Select pen color buttons"> : Select the color of the pen

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Eraser_256.bmp" width="30" alt="Eraser button"> : Switch to eraser tool

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Highlighter_256.bmp" width="30" alt="Highlighter button"> : Switch to highlighter pen

<img src="https://github.com/zbchristian/PenTool/raw/master/images/LaserPointer_256.bmp" width="30" alt="Laser pointer button"> : Use pen as laser pointer

<img src="https://github.com/zbchristian/PenTool/raw/master/images/NewSlide_256.bmp" width="30" alt="New slide button"> : Add an empty slide 

<img src="https://github.com/zbchristian/PenTool/raw/master/images/AllSlides_256.bmp" width="30" alt="All slides button"> : Show slide overview

<img src="https://github.com/zbchristian/PenTool/raw/master/images/PrevSlide_256.bmp" width="15" alt="Goto previous slide button"> : Goto previous slide

<img src="https://github.com/zbchristian/PenTool/raw/master/images/NextSlide_256.bmp" width="15" alt="Goto next slide button"> : Goto next slide

<img src="https://github.com/zbchristian/PenTool/raw/master/images/Exit_256.bmp" width="30" alt="Exit slide show"> : Exit the slide show

**Esc**: Send Escape key

## Customization
Open the file `PenTool.pptm` and start the VBA console (`ALT+F11`). To run the code, you have to execute `InitializeApp` in 
the module `PenTool_Init`: select the module, place the cursor at the end of the code and hit `F5` to execute. Afterwards, goto the 
PPT window and start a slideshow. The toolbar should appear now.
 
After you did your modifications, save the pptm file AND do a `save as` to `PenTool.ppam`. The latter requires some additions. 
Install the `CustomUIEditor`, load the ppam fie and right click on the name in the left hand pane. Select `Office 2010 ...` and 
paste the content of the file `PenTool.xml` into the right pane. Customize the XML content.

Thats it! Now you can load the Add-In again. If you did not change the name, the modifications will be visible at the next start 
of Powerpoint.

# Simple Version
The standard version `PenTool.ppam` sends key shortcuts to Powerpoint to switch to the Eraser, Highlighter and Laser Pointer. 
This might fail, depending on the Office version. The `PenTool_simple.ppam` contains a pure Visual Basic macro based version. This 
version misses the Highlighter and the icon of the Eraser is the laser pointer.

![Toolbar to select th epen color](https://github.com/zbchristian/PenTool/raw/master/images/Screenshot_simple.png)

Enjoy!
