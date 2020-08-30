3 projects in one (work in progress as of 2020/08/29 - Does not work yet)
- **Macro Recorder for PowerPoint**
  - It calculates the delta of all objects between start and stop, and generates the delta VBA code
- Generate VBA code for a given PowerPoint presentation, which recreates the same PowerPoint presentation from scratch
- Create a Powerpoint element manually, select it, run the macro which will re-create the same element (without the methods copy/paste). Useful for understanding how to create the element with VBA by debugging the macro. Maybe superseded by the Macro Recorder above.

# PowerPoint Macro Recorder

TODO

# PPT-VBA-pseudo-recorder
Microsoft has removed the macro recorder in PowerPoint 2007. The following macros can help you determine what VBA code is needed to produce given Powerpoint elements.

# Macro recreate_selected_object
Create a Powerpoint element manually, select it, run the macro which will re-create the same element (without the methods copy/paste).
If the element is correct (i.e. your element doesn't have characteristics that VBA cannot generate), select the element again, run and debug the macro to see what VBA code is used to create the element.
If you just want to generate the VBA code to create the element, use the macro **module** below.

# Macro module
Select any element in your PowerPoint presentation, run the macro, and it will create a file on your laptop containing the VBA code to generate the element.
