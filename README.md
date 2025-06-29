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

- **SlideAnalysis.bas**
List all Slide Masters used in the PowerPoint and prints for each Slide Master the slides which are using layouts from the corresponding Master.

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
Lists all unique Slide Masters used in the PowerPoint and counts how often the same Slide Master was imported. (e.g. If it contains `"23_Blue_theme"`, `"22_Blue_theme"`, and `"Blue_theme"`, the output will tell you you have `"Blue_theme"` 3x imported).

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
- **Printalllayouts.bas**
Prints all custom layouts of the specified Slide Master. The name of the Slide Master needs to be specified in the macro (e.g., here `"Blue_theme"`). 
    
    <ins>Output</ins>:
    ```
    -----START-----
    Found design: Blue_theme
    Title Slide
    Agenda 
    Section, Title 
    ...
    Closing 
    -----END-----
    ```

- **NormalizeSlideDesings.bas**
Normalizes slide designs in the presentation.  
If the same Slide Master is imported multiple times (e.g. `"23_Blue_theme"`, `"22_Blue_theme"`, `"Blue_theme"`), it ensures that all slides are moved to the canonical Slide Master (e.g. `"Blue_theme"`), while preserving the layout used on each slide. After running this macro, use `SlideMasterCleanup.bas` to remove the unused (non-canonical) Slide Masters (`"23_Blue_theme"`, `"22_Blue_theme"`).

- **SlideMasterCleanup.bas**
Removes all unused Slide Masters.  
A Master (and all its layouts) is only removed if none of its layouts is used in the presentation.

- **ReplaceOldDesign.bas**
Replaces the layout of a slide with the layout of the new Slide Master which has to be specified in the code. Iterates through all slides in the powerpoint, but a slide will only be changed if the current layout name matches any layout name in the new Master.

- **ReplaceOldDesign2.bas**
Replaces the layout of a slide (from a specified old Slide Master) with the layout of the new Slide Master which has to be specified in the code. This macro can be used if the layout names don't match. To make a mapping, a manual mapping has to be done and specified in the code.

## Useful Flows

### Problem 1
You have **duplicate Slide Masters** (e.g. `"23_Blue_theme"`, `"22_Blue_theme"`, `"Blue_theme"`), and you want to consolidate your presentation so that only one Master is used.

#### Solution:
Run the following scripts in this order:

1. **SlideAnalysis**
2. **SlideAnalysis2**
3. **NormalizeSlideDesigns**
4. **SlideMasterCleanup**
5. **SlideAnalysis** (again, to verify)
6. **SlideAnalysis2** (again, to verify)
---
### Problem 2
You want to move your slides from an existing Slide Master (e.g. `"Blue_theme_2024"`) to a new Slide master (e.g. `"Blue_theme_2025"`).

#### Solution:
Run the following scripts in this order:

1. Import your new Slide Master (e.g. copy over a slide to the deck).

2. Run **SlideAnalysis2.bas** to verify that all Slide Master in your deck are imported only once. If not, go to Problem 1. 

3. Depending on layout name matching:
    - If you expect the layouts to be named the same, run `ReplaceOldDesign.bas`
    - If the layouts are named differently, you need to do a mapping of the old layout to the new layout (e.g., `Heading` from `"Blue_theme_2024"` becomes `Title` from `"Blue_theme_2025"`), paste this in `ReplaceOldDesign2.bas` before running it.
      ```vb
      layoutMapping(0, 0) = "Blue_theme_2024" 'Old Slide Master
      layoutMapping(0, 1) = "Blue_theme_2025" 'New Slide Master

      layoutMapping(1, 0)= "Heading"
      layoutMapping(1, 1)= "Title"

      ...
      ```
      Hint: you can use **Printalllayouts.bas** to print the layout names for both Slide Masters. If you choose to print to .txt it will already give you the format of `layoutMapping(0, i)= "Heading"`, for all i. 

4. You can run **SlideAnalysis.bas** to verify that no slide is using the old Slide Master.
5. Run **SlideMasterCleanup.bas** to remove the old Slide Master `"Blue_theme_2024"`
6. Run **SlideAnalysis2.bas** again, to verify.


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
