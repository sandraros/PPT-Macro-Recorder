# PPT-VBA-pseudo-recorder
Microsoft has removed the macro recorder in PowerPoint 2007. The following macros can help you determine what VBA code is needed to produce given Powerpoint elements.

# Macro recreate_selected_object
Create a Powerpoint element manually, select it, run the macro and it will create the same element.
If the element is correct (it means that VBA can), select the element again, run and debug the macro to see what VBA code is used to create the element.
If you just want to generate the VBA code to create the element, use the macro **module** below.

# Macro module
Select any element in your PowerPoint presentation, run the macro, and it will create a file on your laptop containing the VBA code to generate the element.
