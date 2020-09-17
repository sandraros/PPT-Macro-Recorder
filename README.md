**Work in progress as of 2020/09/08 - Does not work yet**

This is a Microsoft Powerpoint add-in to "record" your actions while you change the PowerPoint presentation and generate a corresponding VBA macro, since Microsoft has removed the official Macro Recorder from PowerPoint since version 2007 (see [Microsoft MVP answer here](https://answers.microsoft.com/en-us/msoffice/forum/all/macro-recorder-for-powerpoint-2007) and Stack Overflow [question 1](https://stackoverflow.com/questions/34143374/vba-in-powerpoint-without-macrorecorder) and [question 2](https://stackoverflow.com/questions/381206/recording-vba-code-in-power-point-2007)).

What is a Macro Recorder for?
- End users: save one or more user actions which are often used so that to repeat them in one click
- Developers: fast way to determine the VBA code corresponding to a user action

NB: Microsoft offers macro recorders in Word and Excel, only PowerPoint does not have one.

How to use the PPT Macro Recorder:
- The user presses the START Recorder button
  - A modal window opens up and asks the name of the macro
  - What happens internally: all objects of the current Powerpoint instance are saved as "V1" variables in the global memory for later comparison
- The user does some actions
- The user presses the STOP Recorder button
  - What happens internally: All objects of the current Powerpoint instance are saved as "V2" variables in the global memory. There is a comparison between the V1 and V2 variables, anything different produces VBA code to redo the action. For instance, changing the color of an object will produce this code:
    ```
    TODO
    ```
  - The code is saved to the macro in the Visual Basic Editor
- The user may edit the macro (Alt + F11) or run it

So that the add-in can save VBA code, you must enable the option Centre de gestion de la confidentialité > Paramètres > Macros > Accès approuvé au modèle d'objet du projet VBA.

# How does it work

TODO

# OLD

3 projects in one (work in progress as of 2020/08/29 - Does not work yet)
- **Macro Recorder for PowerPoint**
  - It calculates the delta of all objects between start and stop, and generates the delta VBA code
- Generate VBA code for a given PowerPoint presentation, which recreates the same PowerPoint presentation from scratch
- Create a Powerpoint element manually, select it, run the macro which will re-create the same element (without the methods copy/paste). Useful for understanding how to create the element with VBA by debugging the macro. Maybe superseded by the Macro Recorder above.

## PowerPoint Macro Recorder

TODO

## PPT-VBA-pseudo-recorder
Microsoft has removed the macro recorder in PowerPoint 2007. The following macros can help you determine what VBA code is needed to produce given Powerpoint elements.

## Macro recreate_selected_object
Create a Powerpoint element manually, select it, run the macro which will re-create the same element (without the methods copy/paste).
If the element is correct (i.e. your element doesn't have characteristics that VBA cannot generate), select the element again, run and debug the macro to see what VBA code is used to create the element.
If you just want to generate the VBA code to create the element, use the macro **module** below.

## Macro module
Select any element in your PowerPoint presentation, run the macro, and it will create a file on your laptop containing the VBA code to generate the element.
