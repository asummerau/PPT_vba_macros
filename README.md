# Useful PPT Macros for MacOS

1. Start by opening your PowerPoint file, and then navigate to  *Tools > Macros > Visual Basic Editor*.
2. In the Visual Basic Editor, under the *Project* section, choose the project you wish to apply changes to.
3. To import the necessary .bas file, go to  *File > Import File*. Once you've successfully imported the file, it will be added as a Module within your project.
4. Open this module and execute the desired Macro by selecting *Run Macro*. The output will be generated in the *Immediate Window of the Visual Basic Editor*.

## Available Macros
- **NotesCleanup.bas**
Remove all notes from a PowerPoint.

-  **HiddenSlidesCleanup.bas**
Removes all hidden slides from a PowerPoint.

- **ReplaceFooterText.bas**
Updates the years in the PowerPoint footer to match the current year, which is 2025.

- **SideAnalysis.bas**
Generates a list of Master names and the associated slides in your PowerPoint presentation using layouts from the corresponding Master.

- **SlideAnalysis2.bas**
Lists all used (unique) Slide Designs and how often the same Design Layout was imported.

- **SlideMasterCleanup.bas**
Removes all unused Slide Master. A slide Master and all its layouts are only removed if none of the layouts is used in your deck.

- **NormalizeSlideDesings.bas**
Normalizes slide designs in a PowerPoint presentation. When having the same master designs imported multiple time (e.g. "23_Blue_theme" is the same as "Blue_theme"), it will ensure that each slide from "23_Blue_theme" is moved to "Blue_theme" while still keeping the same layout.

## Useful Flows
Problem: You have duplicate Slide Masters (e.g. "23_Blue_theme","22_Blue_theme", "Blue_theme" etc.) and you want to use only one of those identical Slide masters instead.
Fix: run the following scripts  
1. Analysis1
2. Analysis2
3. NormalizeSlideDesings
4. SlideMasterCleanup
5. Analysis1
6. Analysis2