# Useful PPT Macros for MacOS

1. Start by opening your PowerPoint file, and then navigate to  *Tools > Macros > Visual Basic Editor*.
2. In the Visual Basic Editor, under the *Project* section, choose the project you wish to apply changes to.
3. To import the necessary .bas file, go to  *File > Import File*. Once you've successfully imported the file, it will be added as a Module within your project.
4. Open this module and execute the desired Macro by selecting *Run Macro*. The output will be generated in the *Immediate Window of the Visual Basic Editor*.

## Available Macros
- **NotesCleanup.bas**
Removes all notes from the presentation.

-  **HiddenSlidesCleanup.bas**
Removes all hidden slides from the presentation.

- **ReplaceFooterText.bas**
Updates the year in the PowerPoint footer to match the current year (currently 2025).

- **SideAnalysis.bas**
Generates a list of Master names and the associated slides using layouts from the corresponding Master.

    <ins>Output</ins>:
    ```
    -----START-----
    List of all slides using: Green_theme
    PPT Slide #: 3
    PPT Slide #: 4
    List of all slides using: Blue_theme
    PPT Slide #: 1
    PPT Slide #: 2
    PPT Slide #: 5
    PPT Slide #: 6
    -----END-------
    ```
- **SlideAnalysis2.bas**
Lists all unique Slide Designs and counts how often each Design Layout was imported. (e.g. If you dec contains `"23_Blue_theme"`, `"22_Blue_theme"`, and `"Blue_theme"`, the output will tell you you have `"Blue_theme"` 3x imported).

    <ins>Output</ins>:
    ```
    -----START-----
    -----------------------------------
    Design Name ---------------- Count
    -----------------------------------
    Blue_theme ----3
    Green_theme----1
    -----END-------
    ```
- **printalllayouts.bas**
Prints all custom layouts of the specified master design.

- **SlideMasterCleanup.bas**
Removes all unused Slide Masters.  
A Master (and all its layouts) is only removed if none of its layouts is used in the presentation.

- **NormalizeSlideDesings.bas**
Normalizes slide designs in the presentation.  
If the same Master Design is imported multiple times (e.g. `"23_Blue_theme"`, `"22_Blue_theme"`, `"Blue_theme"`), it ensures that all slides are moved to the canonical design (e.g. `"Blue_theme"`), while preserving the layout used on each slide.

- **ReplaceOldDesign.bas**
Replaces the layout of a slide to the layout of the specified Master Design which has to be specified in the code. A slide will only be changed, if the current layout name matches any layout name in the new Master.

- **ReplaceOldDesign2.bas**
Replaces the layout of a slide to the layout of the specified Master Design which has to be specified in the code. This macro can be used if the layout names don't match. To make a mapping, a manual mapping has to be done and specified in the code.

## Useful Flows

### Problem:
You have **duplicate Slide Masters** (e.g. `"23_Blue_theme"`, `"22_Blue_theme"`, `"Blue_theme"`), and you want to consolidate your presentation so that only one Master is used.

### Solution:
Run the following scripts in this order:

1. **SlideAnalysis**
2. **SlideAnalysis2**
3. **NormalizeSlideDesigns**
4. **SlideMasterCleanup**
5. **SlideAnalysis** (again, to verify)
6. **SlideAnalysis2** (again, to verify)
## Disclaimer

These macros were tested on **PowerPoint for MacOS**.  
Use them on copies of your presentations first, to avoid accidental data loss.

---

## Reporting Issues

If you encounter a bug, have a question, or want to suggest a feature:

1. Go to the [Issues](../../issues) tab of this repository.
2. Click **New Issue**.
3. Describe the problem or suggestion as clearly as possible.
4. I'll do my best to review and address it!

Thank you for contributing!

---
