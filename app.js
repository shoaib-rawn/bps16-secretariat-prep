// BPS-16 Prep Dashboard - Core Application Script

// 1. Full 45-Day Syllabus Data with advanced explanations, keyboard shortcuts, pro-tips, and quiz questions.
const SYLLABUS = [
    {
        day: 1,
        category: "word",
        title: "Section Breaks vs. Page Breaks",
        description: "In MS Word, a **Page Break** (Ctrl+Enter) simply pushes text to the next page. A **Section Break** (found in Layout > Breaks) creates a barrier that allows you to apply different formatting (like margins, columns, headers, footers, or orientation) to specific parts of the document. For BPS-16 tasks, you must use Section Breaks (Next Page) to mix Portrait and Landscape orientation or to change header/footer content mid-document.",
        shortcut: "Ctrl + Enter (Page Break) / Alt + P, B (Layout Breaks menu)",
        quickTip: "To see where breaks are placed, turn on formatting marks (Home > Show/Hide ¶ or Ctrl+Shift+8).",
        quiz: {
            question: "Which feature should you use to make only page 3 of a document Landscape, keeping others Portrait?",
            options: [
                "Page Break",
                "Section Break (Next Page)",
                "Column Break",
                "Text Wrapping Break"
            ],
            answer: 1
        }
    },
    {
        day: 2,
        category: "word",
        title: "Custom Page Numbering & Decoupled Headers",
        description: "Official government reports often require Roman numerals (i, ii, iii) for introductory pages (Preface, TOC) and Arabic numerals (1, 2, 3) starting from the main body. To achieve this, insert a Section Break (Next Page) before the main body, double-click the header/footer of Section 2, **uncheck 'Link to Previous'** in the Header & Footer tab, and then format page numbers starting from 1.",
        shortcut: "Double-click Header/Footer > Link to Previous (Toggle off)",
        quickTip: "Unlinking headers and footers is a double-action. Unlinking the header does NOT automatically unlink the footer; they must be unlinked separately.",
        quiz: {
            question: "What is the critical step to prevent page numbers from appearing on the Title Page or in Roman numerals throughout?",
            options: [
                "Delete the header on the title page manually",
                "Use a Page Break and format numbers",
                "Insert a Section Break and deactivate 'Link to Previous' in the second section",
                "Change document views to Draft view"
            ],
            answer: 2
        }
    },
    {
        day: 3,
        category: "word",
        title: "Advanced Mail Merge with Excel Lists",
        description: "Mail Merge is essential for dispatching notices to hundreds of recipients. You can link an Excel sheet containing columns like Name, Designation, and Address. Use the 'Write & Insert Fields' section to place merge fields. You can also use **Rules (e.g., IF-THEN-ELSE)** to customize text. For instance, if 'City' equals 'Islamabad', print 'By Hand', otherwise 'By Post'.",
        shortcut: "Alt + M (Mailings Menu) > S (Start Mail Merge) > E (Letters)",
        quickTip: "Always check your Excel data: ensure row 1 contains unique column headers and there are no merged cells.",
        quiz: {
            question: "What field code allows you to insert conditional text based on a recipient's data during Mail Merge?",
            options: [
                "MERGEFIELD",
                "ASK field",
                "IF...THEN...ELSE... rule",
                "FILLIN field"
            ],
            answer: 2
        }
    },
    {
        day: 4,
        category: "word",
        title: "Document Style Sheets & Automatic TOC",
        description: "For long government files, manually styling headers is inefficient. Use Word's built-in Styles (Heading 1, Heading 2, Heading 3). This lets you globally modify fonts and paragraph spacings. Crucially, applying these styles allows MS Word to generate a dynamic, hyperlinked Table of Contents (TOC) under References > Table of Contents.",
        shortcut: "Ctrl + Alt + 1 (Apply Heading 1) / Ctrl + Alt + 2 (Heading 2)",
        quickTip: "To update an existing Table of Contents after editing, click it and press F9.",
        quiz: {
            question: "Which keyboard shortcut is used to apply Heading Style 1 to selected text?",
            options: [
                "Ctrl + Shift + 1",
                "Ctrl + Alt + 1",
                "Alt + Shift + H1",
                "Ctrl + H + 1"
            ],
            answer: 1
        }
    },
    {
        day: 5,
        category: "word",
        title: "Track Changes & File Protection",
        description: "In government secretariats, files are reviewed by multiple officers (Section Officer, Deputy Secretary, Joint Secretary). Use **Track Changes** (Review > Track Changes) to see insertions, deletions, and comments. To secure a file, use 'Restrict Editing' to limit formatting to specific styles or restrict editing to 'No changes (Read only)' except for tracked revisions.",
        shortcut: "Ctrl + Shift + E (Toggle Track Changes)",
        quickTip: "To accept all revisions at once, click the bottom half of the 'Accept' button and select 'Accept All Changes'.",
        quiz: {
            question: "How can you restrict a Word document so others can only add comments without changing the text?",
            options: [
                "Set the file property to Read-Only in Windows",
                "Use Review > Restrict Editing > Allow only 'Comments'",
                "Save as a text file (.txt)",
                "Password encrypt the document entirely"
            ],
            answer: 1
        }
    },
    {
        day: 6,
        category: "word",
        title: "Captions, Cross-references, and Table of Figures",
        description: "For official documents containing multiple charts or tables, you must number them. Use References > Insert Caption. Word automatically labels them (e.g., Figure 1, Table 1). You can then use **Cross-reference** to refer to them in text (e.g., 'as shown in Table 1 on page 4'). If page layout changes, cross-references and the Table of Figures can be updated instantly.",
        shortcut: "Alt + S, F (Insert Caption) / Alt + S, RF (Cross-reference)",
        quickTip: "Select the entire document (Ctrl+A) and press F9 to update all field values, including page numbers in cross-references.",
        quiz: {
            question: "If you move 'Table 3' to a position before 'Table 1', what is the easiest way to update their sequence numbers?",
            options: [
                "Re-type all captions manually",
                "Select all text and press F9 to update fields",
                "Restart MS Word to force auto-recalculation",
                "Delete and re-insert captions"
            ],
            answer: 1
        }
    },
    {
        day: 7,
        category: "word",
        title: "Advanced Find & Replace with Wildcards",
        description: "Using wildcards in Find and Replace (Home > Replace > More > Use Wildcards) turns MS Word into a powerful text parser. For example, search for `([0-9]{2})-([0-9]{2})-([0-9]{4})` and replace with `\\2/\\1/\\3` to convert US dates (MM-DD-YYYY) to UK style (DD/MM/YYYY) instantly. You can also match word ranges, beginning/end of words, or characters.",
        shortcut: "Ctrl + H (Open Replace Dialog) > check 'Use Wildcards'",
        quickTip: "The wildcard character `?` matches any single character, while `*` matches zero or more characters.",
        quiz: {
            question: "In Word's wildcard search, which character acts as a placeholder for a single digit?",
            options: [
                "?",
                "#",
                "[0-9]",
                "*"
            ],
            answer: 2
        }
    },
    {
        day: 8,
        category: "word",
        title: "Multi-column Layouts & Column Breaks",
        description: "Official newsletters or Gazette notifications require text divided into multiple vertical columns. Under Layout > Columns, select 2 or 3. If you want to force certain text or headings to start at the top of the next column, insert a **Column Break** instead of hitting Enter repeatedly.",
        shortcut: "Ctrl + Shift + Enter (Column Break)",
        quickTip: "To have a single-column title centered above a two-column text layout on the same page, separate them with a Continuous Section Break.",
        quiz: {
            question: "Which keyboard shortcut immediately pushes the cursor and subsequent text to the top of the next column?",
            options: [
                "Ctrl + Enter",
                "Alt + Enter",
                "Ctrl + Shift + Enter",
                "Shift + Enter"
            ],
            answer: 2
        }
    },
    {
        day: 9,
        category: "word",
        title: "Document Inspector & Metadata Removal",
        description: "Before sending official letters outside the Secretariat, it is critical to remove confidential metadata (author names, template sources, document paths, revision logs, hidden text). File > Info > Check for Issues > **Inspect Document** scans for this data and lets you purge it permanently to prevent leaks.",
        shortcut: "Alt + F, I, W (File > Info > Inspect Document shortcut sequence)",
        quickTip: "Tracked changes that are hidden but not accepted/rejected will be flagged and can be removed by the inspector.",
        quiz: {
            question: "Why should a government private secretary run the 'Document Inspector' before emailing files?",
            options: [
                "To check for spelling and grammar errors",
                "To remove hidden personal metadata, tracked changes, and comments",
                "To compress the file size for fast transmission",
                "To convert the layout into a PDF layout"
            ],
            answer: 1
        }
    },
    {
        day: 10,
        category: "word",
        title: "Advanced Text Wrapping & Anchors",
        description: "When inserting pictures, organizational charts, or signature boxes into reports, Word defaults to 'In Line with Text', which treats the image like a massive letter. For custom layouts, change Wrap Text to **Square**, **Tight**, or **Behind/In Front of Text**. An image is anchored (Show ¶ to see the anchor icon) to a paragraph. Moving that paragraph moves the image unless locked.",
        shortcut: "Right-click Image > Wrap Text (or Alt + JP, TW)",
        quickTip: "To position an image precisely in the center of the page regardless of text additions, right-click > Size and Position > Position > set Absolute Position relative to Page.",
        quiz: {
            question: "What is indicated by the anchor icon next to an image in MS Word?",
            options: [
                "The image is linked to an external web URL",
                "The image is locked to a specific page number",
                "The image is bound to a paragraph and moves with it",
                "The image cannot be deleted"
            ],
            answer: 2
        }
    },
    {
        day: 11,
        category: "word",
        title: "Master and Subdocuments for Complex Volumes",
        description: "For long administrative volumes or gazettes, a single Word file becomes slow. Use Outlining View (View > Outline) to create a **Master Document** that links several small **Subdocuments** (e.g., Chapter files). Clicking 'Expand Subdocuments' compiles them into a single continuous stream while keeping individual files editable by different personnel.",
        shortcut: "Alt + W, U (Outline View) > Show Document > Create / Insert",
        quickTip: "Save all subdocuments and the master document in the same subfolder to avoid broken file links.",
        quiz: {
            question: "Which Word document view is required to create, split, and merge subdocuments within a master file?",
            options: [
                "Print Layout View",
                "Outline View",
                "Web Layout View",
                "Read Mode"
            ],
            answer: 1
        }
    },
    {
        day: 12,
        category: "word",
        title: "Automating Formatting with MS Word Macros",
        description: "A Macro is a series of commands recorded or written to automate repetitive tasks (e.g., styling an official memo header). Go to View > Macros > Record Macro. Assign it to a button or shortcut. Complete your formatting steps, then stop recording. The file must be saved as a **Word Macro-Enabled Document (.docm)**.",
        shortcut: "Alt + F8 (Macros Dialog box)",
        quickTip: "Never download or open macro-enabled files (.docm, .xlsm) from untrusted sources, as they can contain harmful VBA scripts.",
        quiz: {
            question: "Which file extension is required to save an MS Word document containing recorded VBA macros?",
            options: [
                ".docx",
                ".docm",
                ".dotx",
                ".rtf"
            ],
            answer: 1
        }
    },
    {
        day: 13,
        category: "word",
        title: "Quick Parts & Document Building Blocks",
        description: "Secretarial posts involve sending letters with repetitive blocks (e.g., standard header, signature blocks, confidentiality notices). Instead of copy-pasting, select the text, go to Insert > Quick Parts > Save Selection to Quick Part Gallery. You can insert this block instantly anywhere by typing its name and pressing F3.",
        shortcut: "Type Quick Part Name + F3 (Inserts the block instantly)",
        quickTip: "You can also use AutoText (a subcategory of Quick Parts) for rapid insertion of official declarations.",
        quiz: {
            question: "How do you quickly insert a pre-saved Quick Part building block named 'Disclaimer' in a document?",
            options: [
                "Type 'Disclaimer' and press F3",
                "Press Ctrl + Shift + D",
                "Double-click where you want it and press F9",
                "Type 'Disclaimer' and press Enter"
            ],
            answer: 0
        }
    },
    {
        day: 14,
        category: "word",
        title: "Fields and Formula Calculations in Tables",
        description: "Word tables can compute numbers without Excel. Click a cell, go to Layout (under Table Tools) > Formula. You can use formulas like `=SUM(ABOVE)`, `=AVERAGE(LEFT)`, or reference cells (A1, B2) just like in Excel. These formulas are **Fields**. If values change, select the cell and press F9 to update.",
        shortcut: "Ctrl + F9 (Insert empty field braces `{ }`) / F9 (Update field)",
        quickTip: "To toggle between viewing a formula and its result, select the cell and press Shift+F9.",
        quiz: {
            question: "What is the standard formula to sum all numeric cells directly above the cursor in a Word table?",
            options: [
                "=TOTAL(ABOVE)",
                "=SUM(UP)",
                "=SUM(ABOVE)",
                "=ADD(ABOVE)"
            ],
            answer: 2
        }
    },
    {
        day: 15,
        category: "word",
        title: "MS Word Shortcuts & Secret Commands",
        description: "Speed is critical for secretarial work. Crucial shortcuts include: **Ctrl+BackSpace** (Delete previous word), **Ctrl+Delete** (Delete next word), **Shift+F3** (Toggle letter case: UPPER, lower, Sentence case), **Ctrl+Shift+K** (Format letters as Small Caps), and **Ctrl+K** (Insert Hyperlink).",
        shortcut: "Shift + F3 (Change Case) / Ctrl + BackSpace (Delete whole word)",
        quickTip: "To quickly select a sentence, hold Ctrl and click anywhere within that sentence.",
        quiz: {
            question: "Which shortcut toggles selected text through UPPERCASE, lowercase, and Title Case configurations?",
            options: [
                "Ctrl + F3",
                "Alt + C",
                "Shift + F3",
                "Ctrl + Shift + U"
            ],
            answer: 2
        }
    },
    {
        day: 16,
        category: "excel",
        title: "Cell Reference Modes (Relative, Absolute, Mixed)",
        description: "Understanding cell referencing is the foundation of Excel logic. **Relative references** (`A1`) change when copied. **Absolute references** (`$A$1`) lock the column and row. **Mixed references** lock only the column (`$A1`) or only the row (`A$1`). In tests, you will be asked to write formulas that can be dragged across a grid without breaking calculation logic.",
        shortcut: "F4 (Cycles through reference modes: A1 -> $A$1 -> A$1 -> $A1)",
        quickTip: "Always use mixed references when creating cross-tabulation matrix summaries.",
        quiz: {
            question: "If a formula contains the reference B$5 and is copied from cell C1 to C2, how does the reference change?",
            options: [
                "It changes to C$5",
                "It changes to B$6",
                "It remains B$5",
                "It changes to C$6"
            ],
            answer: 2
        }
    },
    {
        day: 17,
        category: "excel",
        title: "Dynamic Lookups: VLOOKUP vs. INDEX/MATCH",
        description: "VLOOKUP is standard, but has limits: it only searches left-to-right and breaks if columns are inserted. **INDEX/MATCH** is superior. `MATCH(lookup_val, lookup_range, 0)` returns the relative row position. `INDEX(return_range, row_num)` returns the value. Combined, they lookup data in any column, left or right.",
        shortcut: "Alt + M, L (Insert Lookup and Reference functions)",
        quickTip: "Always set VLOOKUP's fourth argument (Range Lookup) to FALSE or 0 to retrieve exact matches only.",
        quiz: {
            question: "What is the primary limitation of VLOOKUP that is resolved by using INDEX/MATCH?",
            options: [
                "VLOOKUP cannot handle numbers",
                "VLOOKUP can only search from left to right",
                "VLOOKUP does not support exact matching",
                "VLOOKUP requires the sheet to be sorted alphabetically"
            ],
            answer: 1
        }
    },
    {
        day: 18,
        category: "excel",
        title: "Modern Lookups: XLOOKUP Mastery",
        description: "XLOOKUP (Excel 365+) replaces VLOOKUP and HLOOKUP. Syntax: `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`. It searches left or right, defaults to exact match, allows fallback values without using IFERROR, and can return multiple cells.",
        shortcut: "Type =XL and press Tab to autocomplete XLOOKUP",
        quickTip: "Use XLOOKUP to return a horizontal row or vertical column array with a single formula.",
        quiz: {
            question: "What is the default match mode for the XLOOKUP function?",
            options: [
                "Approximate match (1)",
                "Exact match (0)",
                "Wildcard match (2)",
                "First-character match (-1)"
            ],
            answer: 1
        }
    },
    {
        day: 19,
        category: "excel",
        title: "Logical Formulas: Nested IF vs. IFS",
        description: "To evaluate multiple conditions, traditional Excel uses nested IF formulas: `=IF(A1>=90, 'A', IF(A1>=80, 'B', 'C'))`. Modern Excel provides `IFS`: `=IFS(A1>=90, 'A', A1>=80, 'B', TRUE, 'C')`. It simplifies code reading and prevents missing parenthesis errors.",
        shortcut: "Ctrl + Shift + A (Insert function arguments checklist)",
        quickTip: "In IFS, use TRUE as the final condition to act as an 'Else' default return value.",
        quiz: {
            question: "What is the benefit of the IFS function over a nested IF formula?",
            options: [
                "IFS runs faster and requires fewer brackets, writing conditions in a single flat structure",
                "IFS can evaluate text while IF only evaluates numbers",
                "IFS automatically sorts data in ascending order",
                "IFS can search columns backwards"
            ],
            answer: 0
        }
    },
    {
        day: 20,
        category: "excel",
        title: "Conditional Aggregations: SUMIFS, COUNTIFS, AVERAGEIFS",
        description: "For reporting, you must count or sum data based on multiple criteria. Syntax: `=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)`. In government listings, you might sum salaries for staff whose designation is 'Assistant' AND who reside in 'Rawalpindi'.",
        shortcut: "Alt + M, M (Insert Math & Trig formulas)",
        quickTip: "To check criteria with wildcards, use `*` (e.g., '*Assistant*' matches Private Assistant, Senior Assistant).",
        quiz: {
            question: "In SUMIFS, what is the first argument in the formula parameter list?",
            options: [
                "criteria_range1",
                "sum_range",
                "criteria1",
                "logical_test"
            ],
            answer: 1
        }
    },
    {
        day: 21,
        category: "excel",
        title: "Text Cleanup: LEFT, MID, FIND, LEN, and CONCAT",
        description: "Official listings often contain messy string values. Use **FIND** to discover a character's index (e.g., a comma separating names). Combine it with **LEFT** or **MID** to slice text. Use **LEN** to check character count, and **CONCAT** (or `&` operator) to merge columns.",
        shortcut: "Use '&' as a quick operator alternative to CONCAT (e.g. A1 & ' ' & B1)",
        quickTip: "Use TRIM(A1) around lookup formulas to strip out accidental leading/trailing spaces from cell contents.",
        quiz: {
            question: "If cell A1 contains 'President Secretariat', what does the formula =LEFT(A1, FIND(' ', A1)-1) return?",
            options: [
                "Secretariat",
                "President",
                "President Secretariat",
                "Error"
            ],
            answer: 1
        }
    },
    {
        day: 22,
        category: "excel",
        title: "Date Calculations & Workday Scheduling",
        description: "BPS-16 assistants calculate due dates, retirement schedules, or working days. `=NETWORKDAYS(start_date, end_date, [holidays])` calculates actual working days between dates, excluding weekends and a custom list of holidays. `=EDATE(start_date, months)` adds or subtracts months.",
        shortcut: "Ctrl + ; (Insert current date stamp) / Ctrl + Shift + ; (Insert current time stamp)",
        quickTip: "To format a cell as a long date instantly, press Ctrl + Shift + #.",
        quiz: {
            question: "Which function calculates the number of working days between two dates, excluding weekends and a list of specific holidays?",
            options: [
                "DAYS",
                "WORKDAY",
                "NETWORKDAYS",
                "DATEDIF"
            ],
            answer: 2
        }
    },
    {
        day: 23,
        category: "excel",
        title: "Pivot Tables & Pivot Charts",
        description: "Pivot Tables summarize vast sheets of tabular data into cross-tabs. Drag fields into Rows, Columns, and Values. You can filter data or add a **Slicer** (Insert > Slicer) for a user-friendly clickable dashboard. Pivot Charts dynamically update when the pivot table is filtered.",
        shortcut: "Alt + N, V (Insert Pivot Table)",
        quickTip: "To refresh a Pivot Table after updating source sheet data, press Alt + F5 or right-click the pivot and click Refresh.",
        quiz: {
            question: "Which visual filtering tool can you add to a Pivot Table to filter rows with simple clickable buttons?",
            options: [
                "Timeline filter",
                "Slicer",
                "Data Filter Dropdown",
                "Conditional Format"
            ],
            answer: 1
        }
    },
    {
        day: 24,
        category: "excel",
        title: "Formula-based Conditional Formatting",
        description: "Beyond coloring duplicate cells, you can write formulas in Conditional Formatting to highlight entire rows. Select the table rows, set a formatting rule using `= $B2 < 50000` (where column B is Salary). Using the absolute column marker (`$B`) ensures Excel formats the entire row based on column B's value.",
        shortcut: "Alt + H, L (Home > Conditional Formatting menu)",
        quickTip: "Always use relative row references and absolute column references (mixed) in conditional formatting formulas.",
        quiz: {
            question: "What type of cell referencing is needed to highlight an entire row based on the value in column A?",
            options: [
                "Absolute columns and absolute rows ($A$1)",
                "Relative columns and relative rows (A1)",
                "Absolute columns and relative rows ($A1)",
                "Relative columns and absolute rows (A$1)"
            ],
            answer: 2
        }
    },
    {
        day: 25,
        category: "excel",
        title: "Advanced Data Validation & Dropdowns",
        description: "Data Validation (Data > Data Validation) restricts cell inputs to prevent errors. You can allow decimals, dates, or create a drop-down list. To create **dependent dropdowns** (where selecting a Department in cell A1 changes the list of Roles in B1), use the formula `=INDIRECT(A1)` in cell B1's List Validation box.",
        shortcut: "Alt + A, V, V (Open Data Validation Dialog)",
        quickTip: "Clear data validation checks easily by selecting cells, opening validation, and clicking 'Clear All'.",
        quiz: {
            question: "Which function is key to generating dependent dropdown lists that dynamically reference another cell's choice?",
            options: [
                "DIRECT",
                "INDIRECT",
                "XLOOKUP",
                "CHOOSE"
            ],
            answer: 1
        }
    },
    {
        day: 26,
        category: "excel",
        title: "Error Handling: IFERROR, ISERROR, and IFNA",
        description: "Formulas like VLOOKUP return `#N/A` if a value is not found, which looks unprofessional. Wrap calculations in `=IFERROR(formula, value_if_error)`. For instance, `=IFERROR(VLOOKUP(A1, B:C, 2, 0), 'Record Missing')` replaces ugly error codes with readable instructions.",
        shortcut: "Alt + M, E (Insert Logical and Error control options)",
        quickTip: "Use IFNA instead of IFERROR if you only want to trap search failures, leaving syntax or division errors visible for debugging.",
        quiz: {
            question: "What does cell containing `=IFERROR(5/0, 'Divide Error')` display?",
            options: [
                "#DIV/0!",
                "Divide Error",
                "Error",
                "5/0"
            ],
            answer: 1
        }
    },
    {
        day: 27,
        category: "excel",
        title: "What-If Analysis (Goal Seek & Scenarios)",
        description: "What-If Analysis helps you run business forecasts. **Goal Seek** (Data > What-If Analysis > Goal Seek) works backward: you specify a target result (e.g., total budget of 10 million) and the cell to vary, and Excel computes the input value needed to hit that target.",
        shortcut: "Alt + A, W, G (Goal Seek Dialog)",
        quickTip: "Goal Seek can only change one input cell at a time. For multiple input variables, use Excel's Solver Add-in.",
        quiz: {
            question: "Which tool works backwards from a desired output to find the necessary input cell value?",
            options: [
                "Scenario Manager",
                "Goal Seek",
                "Data Table",
                "Pivot Table"
            ],
            answer: 1
        }
    },
    {
        day: 28,
        category: "excel",
        title: "Excel Keyboard Shortcuts Speedrun",
        description: "Essential shortcuts: **Ctrl+1** (Format Cells dialog), **F2** (Edit active cell), **Ctrl+Shift+L** (Toggle AutoFilters on/off), **Ctrl+T** (Convert selection into an official Excel Table), **Alt+=** (Insert AutoSum formula), and **Ctrl+PageUp/PageDown** (Navigate tabs).",
        shortcut: "Alt + = (AutoSum formula shortcut) / Ctrl + 1 (Format Cells dialog)",
        quickTip: "Press Ctrl + Shift + ~ to clear all number formatting and apply the General format to selected cells.",
        quiz: {
            question: "Which keyboard shortcut toggles sorting and filtering dropdown arrows on or off across headers?",
            options: [
                "Ctrl + F",
                "Ctrl + Shift + F",
                "Ctrl + Shift + L",
                "Alt + Down Arrow"
            ],
            answer: 2
        }
    },
    {
        day: 29,
        category: "excel",
        title: "Printing & Page Layouts in Excel",
        description: "Excel sheets are often wider than a piece of paper. To print correctly: set **Print Area**, adjust Orientation to Landscape, and use Page Setup > Fit to 1 page wide. In BPS-16 tests, you must repeat title headings on every printed page using **Print Titles** (Page Layout > Print Titles > 'Rows to repeat at top').",
        shortcut: "Alt + P, I (Print Titles Page Setup)",
        quickTip: "Press Ctrl+F2 to open Print Preview and inspect alignment before sending pages to physical printers.",
        quiz: {
            question: "How do you ensure the top row of a wide spreadsheet repeats on every printed page?",
            options: [
                "Copy and paste the header row manually at page boundaries",
                "Use Page Layout > Print Titles and specify 'Rows to repeat at top'",
                "Set Row Height to fit exactly",
                "Use a Page Break before every print page"
            ],
            answer: 1
        }
    },
    {
        day: 30,
        category: "excel",
        title: "Excel Macro recording and VBA introduction",
        description: "Like Word, Excel lets you record sheet tasks as Macros. Go to Developer > Record Macro. Excel writes code in VBA (Visual Basic for Applications). Save workbooks as **Excel Macro-Enabled Workbooks (.xlsm)**. You can write VBA modules to automate daily reporting templates.",
        shortcut: "Alt + F11 (Open VBA Editor Window)",
        quickTip: "If the Developer tab is missing, customize the ribbon via File > Options > Customize Ribbon and check 'Developer'.",
        quiz: {
            question: "What is the key shortcut used to open the Microsoft Visual Basic for Applications (VBA) editor?",
            options: [
                "Alt + F8",
                "Alt + F11",
                "Ctrl + F11",
                "Shift + F11"
            ],
            answer: 1
        }
    },
    {
        day: 31,
        category: "ppt",
        title: "Slide Master View & Custom Branding",
        description: "To design a uniform presentation for the Secretariat, do not format slides one by one. Use View > Slide Master. Formatting the Slide Master (fonts, logo placements, footers, background colors) applies those styles globally to every slide in the presentation. You can also build custom layouts.",
        shortcut: "Alt + W, M (Open Slide Master view)",
        quickTip: "To exit Slide Master View, click the 'Close Master View' button or switch view back to Normal.",
        quiz: {
            question: "Which view allows you to apply a global logo position or background color to all slides in PowerPoint?",
            options: [
                "Outline View",
                "Slide Master View",
                "Slide Sorter View",
                "Presenter View"
            ],
            answer: 1
        }
    },
    {
        day: 32,
        category: "ppt",
        title: "Transitions vs. Animations & Morph",
        description: "A **Transition** is the motion effect that occurs when moving from one slide to another. An **Animation** is the visual effect applied to individual elements (text, images, shapes) on a slide. The **Morph** transition is highly advanced: it looks at elements on two slides, calculates movement, and animates them smoothly.",
        shortcut: "Alt + T (Transitions Menu) / Alt + A (Animations Menu)",
        quickTip: "Use Morph by duplicating a slide, moving/resizing elements on the second slide, and then applying the Morph transition.",
        quiz: {
            question: "What is the difference between a Transition and an Animation in PowerPoint?",
            options: [
                "Transitions apply to shapes; Animations apply to text",
                "Transitions apply to the slide shift; Animations apply to slide contents",
                "Transitions work only in PDF exports; Animations work in PPT files",
                "Transitions are automatic; Animations require VBA scripting"
            ],
            answer: 1
        }
    },
    {
        day: 33,
        category: "ppt",
        title: "Action Buttons & Custom Slide Shows",
        description: "PowerPoint can act as an interactive kiosk. Insert Shapes > **Action Buttons** (like Home, Back, Next) and link them to slides, URLs, or programs. You can also create a **Custom Slide Show** (Slide Show > Custom Slide Show) to present a subset of slides to different audiences without deleting slides.",
        shortcut: "Alt + S, C (Custom Slide Show Menu)",
        quickTip: "Use Custom Shows to hold 10 slides, showing only slides 1, 3, 5 to one department and 1, 2, 4 to another.",
        quiz: {
            question: "How do you present a specific subset of slides from a single presentation file to different officers?",
            options: [
                "Delete the slides you don't want and save under a new name",
                "Hide slide commands manually before each click",
                "Create a Custom Slide Show",
                "Convert the PPT into an Excel Sheet"
            ],
            answer: 2
        }
    },
    {
        day: 34,
        category: "ppt",
        title: "SmartArt Graphics & Organization Charts",
        description: "Bullet points are boring. Convert listings into **SmartArt Graphics** (Insert > SmartArt) to represent hierarchies, processes, matrices, or cycles. For BPS-16 Secretariats, you will frequently create hierarchical **Organization Charts** showing chain of command from Secretary down to Assistants.",
        shortcut: "Alt + N, M (Insert SmartArt Graphic)",
        quickTip: "Right-click any bulleted text list and select 'Convert to SmartArt' to transform text into graphics instantly.",
        quiz: {
            question: "Which feature is best suited for visual representation of a department's hierarchy and line of reporting?",
            options: [
                "Table Grid",
                "SmartArt Organization Chart",
                "3D Chart Models",
                "Slide Transition"
            ],
            answer: 1
        }
    },
    {
        day: 35,
        category: "ppt",
        title: "Presenter View & Presentation Shortcuts",
        description: "When projecting, check 'Use Presenter View' (Slide Show > Monitor). This shows the audience the slide, while your laptop displays the current slide, your notes, a timer, and a preview of the next slide. During a presentation, press **B** to turn screen black, **W** to turn it white, or press a slide number and **Enter** to jump there.",
        shortcut: "F5 (Start Show from beginning) / Shift + F5 (Start Show from current slide)",
        quickTip: "During a slide show, hold both left and right mouse buttons down for 2 seconds to return to Slide 1.",
        quiz: {
            question: "Which key should a presenter press during a live presentation to immediately hide the slide and display a blank black screen?",
            options: [
                "B",
                "W",
                "Esc",
                "F5"
            ],
            answer: 0
        }
    },
    {
        day: 36,
        category: "ppt",
        title: "Exporting Options, Handouts, and Media Compression",
        description: "Presentations with video or audio files can have large file sizes. Go to File > Info > **Compress Media** to reduce the size without affecting visual quality. To print slides, export them as **Handouts** (File > Export > Create Handouts), which writes slides and blank lines directly to a Word file for note-taking.",
        shortcut: "Alt + F, E, H (File > Export > Create Handouts shortcut sequence)",
        quickTip: "Convert your presentation slides into an MP4 video under Export > Create a Video for display kiosks.",
        quiz: {
            question: "What is the best PowerPoint export option to print slides alongside blank lines for physical notes?",
            options: [
                "Export as JPEG images",
                "Export as Video MP4",
                "Create Handouts in Microsoft Word",
                "Save as PPSX slide show file"
            ],
            answer: 2
        }
    },
    {
        day: 37,
        category: "theory",
        title: "Ribbon and Quick Access Toolbar Customization",
        description: "The **Ribbon** is the strip of tabs at the top of Office apps. You can customize it via File > Options > Customize Ribbon to add hidden tabs (like Developer). The **Quick Access Toolbar (QAT)** is a customizable toolbar above the Ribbon that houses commands you use frequently (e.g., Save, Undo, Print Titles).",
        shortcut: "Alt (Pressing Alt highlights key tips on the Ribbon)",
        quickTip: "Right-click any command in the Ribbon and select 'Add to Quick Access Toolbar' for one-click access.",
        quiz: {
            question: "Where is the Quick Access Toolbar (QAT) located by default?",
            options: [
                "At the very bottom of the document window (Status Bar)",
                "Above or below the Ribbon in the upper left corner",
                "Inside the Review Tab",
                "Floating in the center of the viewport"
            ],
            answer: 1
        }
    },
    {
        day: 38,
        category: "theory",
        title: "Office Clipboard vs. Standard Clipboard",
        description: "The Windows clipboard holds only one item. However, the **Office Clipboard** (Home > Clipboard arrow) can hold up to **24 copied items** (text, images, shapes) from any Office program. You can select items individually or click 'Paste All' to insert multiple snippets at once.",
        shortcut: "Alt + H, FO (Opens the Office Clipboard task pane)",
        quickTip: "The Office Clipboard functions across Excel, Word, and PowerPoint, making cross-app copying seamless.",
        quiz: {
            question: "How many copied items can the Microsoft Office Clipboard hold simultaneously?",
            options: [
                "1 item",
                "12 items",
                "24 items",
                "50 items"
            ],
            answer: 2
        }
    },
    {
        day: 39,
        category: "theory",
        title: "Standard Office File Formats & Extensions",
        description: "Every private secretary must know document file extensions. **.docx** (Word), **.docm** (Word Macro), **.dotx** (Word Template), **.xlsx** (Excel), **.xlsm** (Excel Macro), **.xltx** (Excel Template), **.pptx** (PowerPoint), **.ppsx** (PowerPoint Show - opens directly in full presentation mode), and **.csv** (Comma Separated Values text table).",
        shortcut: "F12 (Open Save As Dialog Box directly)",
        quickTip: "Saving a presentation as a .ppsx file is perfect for presentations given by senior executives, as it bypasses editing view.",
        quiz: {
            question: "Which format opens a PowerPoint presentation directly in full-screen Slide Show mode, bypassing editor controls?",
            options: [
                ".pptx",
                ".ppsx",
                ".potx",
                ".pptm"
            ],
            answer: 1
        }
    },
    {
        day: 40,
        category: "theory",
        title: "Core Windows & File System Shortcuts",
        description: "Operating system shortcuts increase secretarial speed. **Win + D** (Minimize all, show desktop), **Win + L** (Lock screen), **Alt + Tab** (Toggle windows), **Shift + Delete** (Permanently delete item, bypassing recycle bin), **Win + E** (Open File Explorer), and **Ctrl + Shift + Esc** (Open Task Manager).",
        shortcut: "Win + E (File Explorer) / Ctrl + Shift + Esc (Task Manager)",
        quickTip: "Press Win + Shift + S to take custom region screenshots directly to clipboard.",
        quiz: {
            question: "Which keyboard shortcut deletes a file permanently without moving it to the Windows Recycle Bin?",
            options: [
                "Delete key",
                "Ctrl + Delete",
                "Shift + Delete",
                "Alt + Delete"
            ],
            answer: 2
        }
    },
    {
        day: 41,
        category: "theory",
        title: "Trust Center & Macro Security",
        description: "Because VBA macros run code on your system, they present security risks. MS Office controls this through File > Options > **Trust Center** > Trust Center Settings. The default safe configuration is 'Disable all macros with notification', which alerts you with a Yellow Message Bar before running code.",
        shortcut: "Alt + F, T (Open Office Options dialog)",
        quickTip: "To execute macros from a trusted directory without warning prompts, add the folder under Trust Center > Trusted Locations.",
        quiz: {
            question: "What is the default, recommended macro security setting in Microsoft Office?",
            options: [
                "Enable all macros (not recommended; potentially dangerous code can run)",
                "Disable all macros without notification",
                "Disable all macros with notification",
                "Disable all macros except digitally signed macros"
            ],
            answer: 2
        }
    },
    {
        day: 42,
        category: "theory",
        title: "BPS-16 Secretarial Protocol & Correspondence",
        description: "An Assistant Private Secretary (BPS-16) at the President's Secretariat must master official layouts. Correspondence styles include: **Office Memorandum (OM)** (used for inter-departmental policy instructions), **U.O. Notes** (Unofficial Notes for advice/comments exchange), and **Office Orders** (relating to internal staff postings, leaves). The standard font is Times New Roman, and documents must adhere to specific structural margins.",
        shortcut: "Times New Roman (12pt for body, 1.5 line spacing, 1 inch margins)",
        quickTip: "Official memos must list the department code, subject head, reference numbers, date, and a distribution list at the bottom.",
        quiz: {
            question: "Which correspondence format is officially used to communicate policy guidelines between separate government ministries?",
            options: [
                "Office Order",
                "Office Memorandum (OM)",
                "Unofficial (U.O.) Note",
                "Notification"
            ],
            answer: 1
        }
    },
    {
        day: 43,
        category: "theory",
        title: "Mock Theory Exam: Part 1 (Word & Excel)",
        description: "Test your speed and memory on Word and Excel. Review cell references, index formulas, section breaks, document protection levels, mail merge connections, and page layout setups.",
        shortcut: "F1 (Help) / Alt + Q (Tell me search field)",
        quickTip: "A quick review of all previous Word and Excel days is key to passing the theoretical portion of the skill test.",
        quiz: {
            question: "In Excel, if you want a reference cell B3 to not shift at all when dragging, how should it be written?",
            options: [
                "B3",
                "$B3",
                "B$3",
                "$B$3"
            ],
            answer: 3
        }
    },
    {
        day: 44,
        category: "theory",
        title: "Mock Theory Exam: Part 2 (PowerPoint & OS)",
        description: "Test your speed and memory on PowerPoint and Windows commands. Review Slide Master formatting, presenter tools, transitions, action keys, file extensions, and basic secretarial procedures.",
        shortcut: "Review F5, Shift+F5, Win+D, Shift+Delete, and Alt+Tab keys",
        quickTip: "Most Secretariat tests place a heavy focus on slide management, animation limitations, and keyboard control efficiency.",
        quiz: {
            question: "Which of the following slide formats is used to define global design schemes across PowerPoint presentation pages?",
            options: [
                "Slide Master",
                "Normal View",
                "Slide Sorter",
                "Outline Slide"
            ],
            answer: 0
        }
    },
    {
        day: 40,
        category: "theory",
        title: "Exam Day Protocol & Final Checklist (Part 1 Written + Part 2 Speed)",
        description: "Final 40-Day exam strategy: 1. Keep hands warm for shorthand and typing speed test. 2. In written test (100 Marks), attempt high-confidence questions first and manage time (90 mins). 3. In typing, accuracy yields speed; typing at 99% accuracy is faster than correcting mistakes. 4. In MS Office, save your file (Ctrl+S) every 30 seconds to prevent lost work during practical exam.",
        shortcut: "Ctrl + S (Save Document - use continuously during practical exams!)",
        quickTip: "Bring your original CNIC, admission letter, shorthand notepad, and HB/2B pencils to the Secretariat testing venue.",
        quiz: {
            question: "What is the single most important habit during a live MS Office practical exam?",
            options: [
                "Applying complex fonts to make the sheet colorful",
                "Using the mouse instead of keyboard shortcuts",
                "Pressing Ctrl + S frequently to safeguard your work",
                "Finishing in half the allocated time regardless of mistakes"
            ],
            answer: 2
        }
    }
];

// 2. Typing Practice Passages for Custom Drill (10 Passages Total)
const TYPING_PASSAGES = [
    "The President's Secretariat, Aiwan-e-Sadr, functions as the administrative core for the head of state, coordinating official state affairs, diplomatic receptions, and statutory declarations. As an Assistant Private Secretary, precision in official communication, speed in shorthand dictation, and mastery over file management systems are core prerequisites for managing the workflow of high-ranking government officials.",
    "Microsoft Word section breaks are fundamental tools for crafting professional reports. They allow the author to divide a document into segments that can have distinct page numbering formats, header templates, and orientation margins. Without section breaks, uncoupling a title page or inserting landscape sheets within a portrait dossier is impossible, which is why examiners focus heavily on this topic.",
    "Excel spreadsheet models are widely deployed in ministries to track budgets, personnel registers, and monthly allocations. Private secretaries are expected to know lookup functions like INDEX and MATCH, logical calculations using nested IFS, and pivot summaries. Ensuring proper print layouts with repeating title rows is critical to creating clean reports for executive reviews.",
    "Pitman shorthand, developed by Sir Isaac Pitman in eighteen thirty-seven, is a phonetic outline writing system based on geometric lines and curves. It uses thin and thick strokes representing voiced and unvoiced consonants, and specific vowel indicators. Master private secretaries memorize thousands of word forms and contraction rules to comfortably capture spoken dictations at rates exceeding one hundred and twenty words per minute.",
    "The Rules of Business, nineteen seventy-three, outline standard government operating procedures in Pakistan. Every administrative file travels from Section Officers through Assistant Secretaries, Joint Secretaries, and finally the Secretary. The Private Secretary handles scheduling, records incoming and outgoing mail in the diary register, drafts notices, and maintains absolute file confidentiality.",
    "Mail Merge field codes in MS Word are surrounded by double chevron symbols. They act as placeholders that link directly to database queries. You can execute high-volume print runs of administrative notices or certificates by compiling Excel directories with custom greeting fields. Incorporating rules like checking blank records ensures letters look manually written.",
    "Excel Pivot Tables are interactive data summarizing engines. They enable data analysts to group dates, sum financial expenditures across distinct offices, and create customized reporting dashboards. Inserting a slicer visual component provides a professional button layout to sort records, which is critical during ministerial budget reviews.",
    "Slide Master view in PowerPoint allows you to modify the global theme, background graphics, and fonts. Setting standard layouts ensures that every presenter in the Secretariat delivers slides under unified branding. The slide sorter allows the private secretary to rapidly group, reorder, or section presentation files.",
    "Advanced keyboard straight typing requires correct hand posture and tactile feedback. By resting index fingers on the F and J keys, typists build muscle memory for letters. A relaxed wrist posture, steady typing cadence, and typing without looking down at the keyboard are essential traits for exceeding sixty words per minute.",
    "Official government communications require structured document layouts. An Office Memorandum is written in the third person and is used to circulate policy details across ministries. In contrast, an Office Order deals with internal actions such as postings, leaf approvals, or transfers, and is distributed to auditing offices."
];

// 3. Motivational Quotes & Guidance
const MOTIVATION_INSIGHTS = [
    {
        quote: "Success is the sum of small efforts, repeated day in and day out. 40 days of focused discipline will separate you from the competition.",
        author: "Robert Collier"
    },
    {
        quote: "The only limit to our realization of tomorrow will be our doubts of today.",
        author: "Franklin D. Roosevelt"
    },
    {
        quote: "It always seems impossible until it's done. Stay dedicated to your daily practice.",
        author: "Nelson Mandela"
    },
    {
        quote: "Quality is not an act, it is a habit. Make your practice drills flawless.",
        author: "Aristotle"
    },
    {
        quote: "Do not wait; the time will never be 'just right.' Start where you stand, and work with whatever tools you may have at your command.",
        author: "Napoleon Hill"
    }
];

let historyStateUpdateInProgress = false;

// 4. Application State (Loaded from Local Storage or initialized)
let state = {
    currentDay: 1,
    completedDays: [],
    shorthandLogs: [],
    typingLogs: []
};

// Load state from local storage
function loadState() {
    try {
        const saved = localStorage.getItem('bps16_prep_state_v2');
        if (saved) {
            state = JSON.parse(saved);
        }
    } catch (e) {
        console.warn("LocalStorage is restricted or blocked in this browser context. Defaulting to in-memory state.");
    }
}

// Save state to local storage
function saveState() {
    try {
        localStorage.setItem('bps16_prep_state_v2', JSON.stringify(state));
    } catch (e) {
        console.warn("LocalStorage write failed. Session changes will reset upon page reload.");
    }
}

// 5. Initialize Application
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    setupTabs();
    initDashboard();
    initSyllabus();
    initAnalytics();
    initTypingDrill();
    initProtocolCenter();
    initWrittenExam();
    updateCountdown();
});

// Tab Switch Logic
function setupTabs() {
    const tabs = document.querySelectorAll('.tab-btn');
    const contents = document.querySelectorAll('.tab-content');
    
    // Set initial history state if empty
    if (!history.state) {
        history.replaceState({ type: 'tab', tabId: 'dashboard' }, '', window.location.pathname);
    }

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            contents.forEach(c => c.classList.remove('active'));
            
            tab.classList.add('active');
            const targetId = tab.dataset.tab;
            const targetEl = document.getElementById(targetId);
            targetEl.classList.add('active');
            
            // Smoothly scroll to active section on mobile/tablet screens
            if (window.innerWidth <= 1024) {
                targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
            
            // Push history state
            if (!historyStateUpdateInProgress) {
                history.pushState({ type: 'tab', tabId: targetId }, '', window.location.pathname);
            }
            
            if (targetId === 'analytics') {
                renderCharts();
            } else if (targetId === 'typing') {
                setTimeout(() => {
                    document.getElementById('typing-passage-container')?.focus();
                }, 50);
            } else if (targetId === 'written') {
                document.getElementById('prep-day-selector').classList.add('d-none');
                document.getElementById('prep-subject-grid').classList.remove('d-none');
            }
        });
    });
}

// Initialize Dashboard Overview Tab
function initDashboard() {
    // Motivation Quote
    const quoteObj = MOTIVATION_INSIGHTS[Math.floor(Math.random() * MOTIVATION_INSIGHTS.length)];
    document.getElementById('m-quote').innerText = `"${quoteObj.quote}"`;
    document.getElementById('m-author').innerText = `- ${quoteObj.author}`;
    
    // Log submission handling
    const logForm = document.getElementById('log-speed-form');
    if (logForm) {
        logForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const logType = document.getElementById('log-type').value;
            const speed = parseInt(document.getElementById('log-speed').value);
            const accuracy = parseFloat(document.getElementById('log-accuracy').value);
            const date = new Date().toISOString().split('T')[0];
            
            if (!speed || !accuracy) {
                showToast("Please enter valid speed and accuracy values!");
                return;
            }
            
            if (logType === 'shorthand') {
                state.shorthandLogs.push({ date, speed, accuracy });
                showToast("Shorthand speed logged!");
            } else {
                state.typingLogs.push({ date, wpm: speed, accuracy });
                showToast("Typing speed logged!");
            }
            
            saveState();
            logForm.reset();
            updateLogTables();
            if (document.getElementById('analytics').classList.contains('active')) {
                renderCharts();
            }
        });
    }
    
    // Today's schedule checkbox interaction
    const checkoffs = document.querySelectorAll('.schedule-chk');
    checkoffs.forEach(chk => {
        chk.addEventListener('change', () => {
            showToast("Activity updated! Keep pushing!");
        });
    });
}

// Initialize MS Office Portal (Syllabus Sidebar & Detail Viewer)
function initSyllabus() {
    const listContainer = document.getElementById('day-list-container');
    listContainer.innerHTML = '';
    
    SYLLABUS.forEach(lesson => {
        const btn = document.createElement('button');
        btn.className = `day-item-btn ${lesson.day === state.currentDay ? 'active' : ''}`;
        if (state.completedDays.includes(lesson.day)) {
            btn.classList.add('completed');
        }
        
        btn.innerHTML = `
            <span>Day ${lesson.day}: ${lesson.title}</span>
            <span class="day-badge-status"></span>
        `;
        
        btn.addEventListener('click', () => {
            document.querySelectorAll('.day-item-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderLessonDetail(lesson);
        });
        
        listContainer.appendChild(btn);
    });
    
    // Render the initial active day
    const activeLesson = SYLLABUS[state.currentDay - 1] || SYLLABUS[0];
    renderLessonDetail(activeLesson);
}

// Render selected lesson in detail panel
function renderLessonDetail(lesson) {
    document.getElementById('v-day-num').innerText = lesson.day;
    document.getElementById('v-day-title').innerText = lesson.title;
    
    const catTag = document.getElementById('v-category-tag');
    catTag.innerText = lesson.category.toUpperCase();
    catTag.className = `lesson-tag tag-${lesson.category}`;
    
    // Markdown-to-HTML helper for simple bold formatting
    let detailedHTML = lesson.description
        .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
        .replace(/\`([^`]+)\`/g, '<code style="background: rgba(0,0,0,0.3); padding: 2px 5px; border-radius: 4px; font-family: var(--font-mono); color: var(--accent-gold);">$1</code>');
    
    document.getElementById('v-description').innerHTML = `<p>${detailedHTML}</p>`;
    document.getElementById('v-shortcut-key').innerText = lesson.shortcut;
    document.getElementById('v-quick-tip').innerText = lesson.quickTip;
    
    // Complete/Mark off checkbox
    const markBtn = document.getElementById('mark-complete-btn');
    if (state.completedDays.includes(lesson.day)) {
        markBtn.innerText = "Completed (Click to Reset)";
        markBtn.style.background = "rgba(16, 185, 129, 0.2)";
        markBtn.style.border = "1px solid var(--accent-emerald)";
        markBtn.style.color = "var(--accent-emerald)";
    } else {
        markBtn.innerText = "Mark Day as Completed";
        markBtn.style.background = "var(--accent-gold)";
        markBtn.style.border = "none";
        markBtn.style.color = "#0b110e";
    }
    
    markBtn.onclick = () => {
        if (state.completedDays.includes(lesson.day)) {
            state.completedDays = state.completedDays.filter(d => d !== lesson.day);
            showToast(`Day ${lesson.day} marked incomplete.`);
        } else {
            state.completedDays.push(lesson.day);
            showToast(`Excellent job! Day ${lesson.day} checked off.`);
            if (lesson.day === state.currentDay && state.currentDay < 45) {
                state.currentDay++;
            }
        }
        saveState();
        initSyllabus();
        initDashboard();
        renderLessonDetail(lesson);
    };
    
    // Load Mini Quiz
    initMiniQuiz(lesson);
}

// Mini Quiz Renderer
function initMiniQuiz(lesson) {
    const quizArea = document.getElementById('mini-quiz-area');
    quizArea.innerHTML = '';
    
    if (!lesson.quiz) return;
    
    const qTitle = document.createElement('h4');
    qTitle.style.marginBottom = '12px';
    qTitle.style.fontSize = '1rem';
    qTitle.innerText = `Daily Skill Challenge: ${lesson.quiz.question}`;
    quizArea.appendChild(qTitle);
    
    lesson.quiz.options.forEach((option, idx) => {
        const div = document.createElement('div');
        div.className = 'quiz-option';
        div.innerText = option;
        
        div.addEventListener('click', () => {
            const options = quizArea.querySelectorAll('.quiz-option');
            options.forEach(opt => opt.style.pointerEvents = 'none'); // lock answers
            
            if (idx === lesson.quiz.answer) {
                div.classList.add('correct');
                showToast("Correct Answer! Well done.");
            } else {
                div.classList.add('incorrect');
                options[lesson.quiz.answer].classList.add('correct');
                showToast("Incorrect, study the shortcut tips again!");
            }
        });
        
        quizArea.appendChild(div);
    });
}

// 6. Analytics & Performance Tracking Tab
function initAnalytics() {
    updateLogTables();
}

function updateLogTables() {
    const sBody = document.getElementById('shorthand-log-body');
    const tBody = document.getElementById('typing-log-body');
    
    if (sBody) {
        sBody.innerHTML = '';
        state.shorthandLogs.slice().reverse().forEach(log => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${log.date}</td>
                <td><strong>${log.speed} WPM</strong></td>
                <td>${log.accuracy}%</td>
            `;
            sBody.appendChild(tr);
        });
    }
    
    if (tBody) {
        tBody.innerHTML = '';
        state.typingLogs.slice().reverse().forEach(log => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${log.date}</td>
                <td><strong>${log.wpm} WPM</strong></td>
                <td>${log.accuracy}%</td>
            `;
            tBody.appendChild(tr);
        });
    }
}

// Render dynamic charts as inline SVGs
function renderCharts() {
    renderSvgLineChart('shorthand-chart-container', state.shorthandLogs, 'speed', '#10b981');
    renderSvgLineChart('typing-chart-container', state.typingLogs, 'wpm', '#06b6d4');
}

function renderSvgLineChart(containerId, logData, valKey, strokeColor) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    container.innerHTML = '';
    
    if (logData.length < 2) {
        container.innerHTML = '<span style="color: var(--text-muted);">Log at least 2 practice sessions to render progress chart</span>';
        return;
    }
    
    // Use a fixed virtual coordinate grid for drawing, and rely on responsive viewBox scaling
    const w = 450;
    const h = 220;
    const padding = 30;
    
    // Find min and max boundaries
    const vals = logData.map(d => d[valKey]);
    let maxVal = Math.max(...vals);
    let minVal = Math.min(...vals);
    
    if (maxVal === minVal) {
        maxVal += 5;
        minVal -= 5;
    }
    const valRange = maxVal - minVal;
    
    const count = logData.length;
    const stepX = (w - (padding * 2)) / (count - 1);
    
    let points = [];
    for (let i = 0; i < count; i++) {
        const x = padding + (i * stepX);
        const ratio = (logData[i][valKey] - minVal) / valRange;
        const y = h - padding - (ratio * (h - (padding * 2)));
        points.push({ x, y, val: logData[i][valKey], date: logData[i].date });
    }
    
    let pathD = `M ${points[0].x} ${points[0].y}`;
    for (let i = 1; i < points.length; i++) {
        pathD += ` L ${points[i].x} ${points[i].y}`;
    }
    
    let gridLinesHTML = '';
    // Draw Y axis guidelines (3 markers)
    const midVal = Math.round((maxVal + minVal) / 2);
    const yMax = padding;
    const yMin = h - padding;
    const yMid = (yMax + yMin) / 2;
    
    gridLinesHTML += `
        <line x1="${padding}" y1="${yMax}" x2="${w - padding}" y2="${yMax}" stroke="rgba(255,255,255,0.05)" />
        <text x="${padding - 5}" y="${yMax + 4}" fill="var(--text-muted)" font-size="9" text-anchor="end">${maxVal}</text>
        <line x1="${padding}" y1="${yMid}" x2="${w - padding}" y2="${yMid}" stroke="rgba(255,255,255,0.05)" stroke-dasharray="4" />
        <text x="${padding - 5}" y="${yMid + 4}" fill="var(--text-muted)" font-size="9" text-anchor="end">${midVal}</text>
        <line x1="${padding}" y1="${yMin}" x2="${w - padding}" y2="${yMin}" stroke="rgba(255,255,255,0.1)" />
        <text x="${padding - 5}" y="${yMin + 4}" fill="var(--text-muted)" font-size="9" text-anchor="end">${minVal}</text>
    `;
    
    let dotsHTML = '';
    points.forEach(pt => {
        dotsHTML += `
            <circle cx="${pt.x}" cy="${pt.y}" r="4" fill="${strokeColor}" />
            <text x="${pt.x}" y="${pt.y - 8}" fill="#ffffff" font-size="8" text-anchor="middle" font-family="var(--font-mono)">${pt.val}</text>
            <text x="${pt.x}" y="${h - 10}" fill="var(--text-muted)" font-size="7" text-anchor="middle">${pt.date.substring(5)}</text>
        `;
    });
    
    const svgHTML = `
        <svg viewBox="0 0 ${w} ${h}" width="100%" height="auto" class="chart-svg" style="display: block; max-width: 100%;">
            ${gridLinesHTML}
            <path d="${pathD}" fill="none" stroke="${strokeColor}" stroke-width="2.5" />
            ${dotsHTML}
        </svg>
    `;
    
    container.innerHTML = svgHTML;
}

// 7. Interactive Typing Drill Module
let currentPassageIndex = 0;
let characterIndex = 0;
let typingErrors = 0;
let typingStartTime = null;
let drillActive = false;

function initTypingDrill() {
    const drillText = document.getElementById('typing-passage-text');
    const typingContainer = document.getElementById('typing-passage-container');
    const resetBtn = document.getElementById('btn-reset-drill');
    const wpmResult = document.getElementById('drill-wpm');
    const accResult = document.getElementById('drill-acc');
    
    if (!drillText || !typingContainer) return;
    
    let currentMode = 'passage'; // 'passage' or 'row'
    let currentRowFocus = 'homerow'; // homerow, toprow, bottomrow, allrows
    let currentDrillText = '';
    
    // Keybr-style random drills generator
    function generateRowDrillText(level) {
        const keysMap = {
            homerow: ['a', 's', 'd', 'f', 'j', 'k', 'l', ';'],
            toprow: ['q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p'],
            bottomrow: ['z', 'x', 'c', 'v', 'b', 'n', 'm', ',', '.'],
            allrows: ['a', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'l', ';', 'q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', 'z', 'x', 'c', 'v', 'b', 'n', 'm', ',', '.']
        };
        
        const keys = keysMap[level] || keysMap.homerow;
        let resultWords = [];
        
        for (let w = 0; w < 20; w++) {
            let wordLen = Math.floor(Math.random() * 3) + 3; // length 3 to 5
            let word = '';
            for (let c = 0; c < wordLen; c++) {
                const randomKey = keys[Math.floor(Math.random() * keys.length)];
                word += randomKey;
            }
            resultWords.push(word);
        }
        return resultWords.join(' ');
    }
    
    function getActiveText() {
        if (currentMode === 'passage') {
            return TYPING_PASSAGES[currentPassageIndex];
        } else {
            if (!currentDrillText) {
                currentDrillText = generateRowDrillText(currentRowFocus);
            }
            return currentDrillText;
        }
    }
    
    function getKeyIdFromEvent(e) {
        if (!e) return '';
        let code = '';
        let key = '';
        if (e && typeof e === 'object') {
            code = e.code;
            key = e.key;
        } else {
            key = e;
        }
        
        if (code) {
            if (code === 'Space') return 'key-space';
            if (code === 'Backspace') return 'key-backspace';
            if (code === 'Tab') return 'key-tab';
            if (code === 'CapsLock') return 'key-capslock';
            if (code === 'Enter') return 'key-enter';
            if (code === 'ShiftLeft') return 'key-shiftleft';
            if (code === 'ShiftRight') return 'key-shiftright';
            if (code === 'ControlLeft') return 'key-controlleft';
            if (code === 'ControlRight') return 'key-controlright';
            if (code === 'AltLeft') return 'key-altleft';
            if (code === 'AltRight') return 'key-altright';
            if (code === 'MetaLeft') return 'key-metaleft';
            if (code === 'MetaRight') return 'key-metaright';
        }
        
        const lowerKey = (key || '').toLowerCase();
        if (lowerKey === '`' || lowerKey === '~') return 'key-backtick';
        if (lowerKey === '-') return 'key-minus';
        if (lowerKey === '=') return 'key-equal';
        if (lowerKey === '[') return 'key-bracketleft';
        if (lowerKey === ']') return 'key-bracketright';
        if (lowerKey === '\\') return 'key-backslash';
        if (lowerKey === ';') return 'key-semi';
        if (lowerKey === '\'') return 'key-quote';
        if (lowerKey === ',') return 'key-comma';
        if (lowerKey === '.') return 'key-period';
        if (lowerKey === '/') return 'key-slash';
        if (lowerKey === ' ') return 'key-space';
        
        return 'key-' + lowerKey;
    }
    
    function highlightNextKey(char) {
        document.querySelectorAll('.kb-key').forEach(key => {
            key.classList.remove('highlighted');
        });
        
        if (!char) return;
        
        // Auto-highlight Shift key if next letter is capital
        const isUppercase = (char >= 'A' && char <= 'Z');
        if (isUppercase) {
            // Touch typing rules: Left hand keys -> Right Shift, Right hand keys -> Left Shift
            const leftHandCapitals = ['Q', 'W', 'E', 'R', 'T', 'A', 'S', 'D', 'F', 'G', 'Z', 'X', 'C', 'V', 'B'];
            if (leftHandCapitals.includes(char)) {
                document.getElementById('key-shiftright')?.classList.add('highlighted');
            } else {
                document.getElementById('key-shiftleft')?.classList.add('highlighted');
            }
        }
        
        const targetId = getKeyIdFromEvent(char);
        const targetKey = document.getElementById(targetId);
        if (targetKey) {
            targetKey.classList.add('highlighted');
        }
    }
    
    function resetDrill() {
        characterIndex = 0;
        typingErrors = 0;
        typingStartTime = null;
        drillActive = false;
        wpmResult.innerText = '0';
        accResult.innerText = '100%';
        
        const activeText = getActiveText();
        
        // Render characters
        drillText.innerHTML = activeText.split('').map((char, index) => {
            return `<span class="${index === 0 ? 'current' : ''}">${char}</span>`;
        }).join('');
        
        highlightNextKey(activeText[0]);
    }
    
    // Toggle Modes Listeners
    const modePassageBtn = document.getElementById('btn-mode-passage');
    const modeRowBtn = document.getElementById('btn-mode-row');
    const rowControls = document.getElementById('row-drill-controls');
    const rowSelect = document.getElementById('row-drill-level');
    const drillModeLabel = document.getElementById('drill-mode-label');
    const nextPassageBtn = document.getElementById('btn-next-passage');
    
    modePassageBtn?.addEventListener('click', () => {
        if (currentMode === 'passage') return;
        currentMode = 'passage';
        modePassageBtn.classList.add('active');
        modeRowBtn.classList.remove('active');
        rowControls.classList.add('d-none');
        drillModeLabel.innerText = "Mode: Passage Practice";
        nextPassageBtn.innerText = "Change Passage";
        resetDrill();
        typingContainer.focus();
    });
    
    modeRowBtn?.addEventListener('click', () => {
        if (currentMode === 'row') return;
        currentMode = 'row';
        modeRowBtn.classList.add('active');
        modePassageBtn.classList.remove('active');
        rowControls.classList.remove('d-none');
        drillModeLabel.innerText = "Mode: Row Touch Typing";
        nextPassageBtn.innerText = "New Row Drill";
        currentDrillText = '';
        resetDrill();
        typingContainer.focus();
    });
    
    rowSelect?.addEventListener('change', () => {
        currentRowFocus = rowSelect.value;
        currentDrillText = '';
        resetDrill();
        typingContainer.focus();
    });
    
    // Next/Change Passage Button
    nextPassageBtn?.addEventListener('click', () => {
        if (characterIndex >= 10 && typingStartTime) {
            const timeDiffMin = (new Date() - typingStartTime) / 60000;
            const grossWpm = Math.round((characterIndex / 5) / (timeDiffMin || 0.01));
            const netWpm = Math.max(0, grossWpm - Math.round(typingErrors / (timeDiffMin || 0.01)));
            const accuracy = Math.round(((characterIndex - typingErrors) / characterIndex) * 100);
            
            const date = new Date().toISOString().split('T')[0];
            state.typingLogs.push({ date, wpm: netWpm, accuracy });
            saveState();
            updateLogTables();
            showToast(`Logged attempt before changing: ${netWpm} WPM | Acc: ${accuracy}%`);
        }
        
        if (currentMode === 'passage') {
            currentPassageIndex = (currentPassageIndex + 1) % TYPING_PASSAGES.length;
        } else {
            currentDrillText = generateRowDrillText(currentRowFocus);
        }
        resetDrill();
        showToast(currentMode === 'passage' ? "Loaded new drill passage!" : "Regenerated row drill characters!");
        typingContainer.focus();
    });
    
    // Auto-focus container when clicking inside the test card
    document.querySelector('.typing-test-box')?.addEventListener('click', () => {
        typingContainer.focus();
    });
    
    // Listen to keypresses on the container (type directly)
    typingContainer.addEventListener('keydown', (e) => {
        // Caps Lock sync
        if (e.getModifierState) {
            const capsEl = document.getElementById('key-capslock');
            if (capsEl) {
                if (e.getModifierState('CapsLock')) {
                    capsEl.classList.add('pressed');
                } else {
                    capsEl.classList.remove('pressed');
                }
            }
        }

        // Prevent default actions for standard keys (like space bar scroll)
        if (e.key === ' ' || e.key === 'Backspace' || (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey)) {
            e.preventDefault();
        } else {
            // Shift, CapsLock, Ctrl, Tab etc. visual keypress highlights
            const targetId = getKeyIdFromEvent(e);
            const keyEl = document.getElementById(targetId);
            if (keyEl) {
                keyEl.classList.add('pressed');
            }
            return;
        }
        
        const targetId = getKeyIdFromEvent(e);
        const keyEl = document.getElementById(targetId);
        if (keyEl) {
            keyEl.classList.add('pressed');
        }
        
        const text = getActiveText();
        const textSpans = drillText.querySelectorAll('span');
        
        if (e.key === 'Backspace') {
            if (characterIndex > 0) {
                characterIndex--;
                textSpans[characterIndex].className = 'current';
                for (let i = characterIndex + 1; i < textSpans.length; i++) {
                    textSpans[i].className = '';
                }
                
                typingErrors = 0;
                for (let i = 0; i < characterIndex; i++) {
                    if (textSpans[i].className === 'incorrect') {
                        typingErrors++;
                    }
                }
                highlightNextKey(text[characterIndex]);
            }
            return;
        }
        
        // Printable character
        if (e.key.length === 1 && characterIndex < text.length) {
            if (!typingStartTime && characterIndex === 0) {
                typingStartTime = new Date();
                drillActive = true;
            }
            
            const typedChar = e.key;
            const expectedChar = text[characterIndex];
            
            if (typedChar === expectedChar) {
                textSpans[characterIndex].className = 'correct';
            } else {
                textSpans[characterIndex].className = 'incorrect';
                typingErrors++;
            }
            
            characterIndex++;
            
            if (characterIndex < text.length) {
                textSpans[characterIndex].className = 'current';
                highlightNextKey(text[characterIndex]);
            } else {
                highlightNextKey(null);
            }
            
            // Calculate stats
            if (characterIndex > 0) {
                const timeDiffMin = (new Date() - typingStartTime) / 60000;
                const grossWpm = Math.round((characterIndex / 5) / (timeDiffMin || 0.01));
                const netWpm = Math.max(0, grossWpm - Math.round(typingErrors / (timeDiffMin || 0.01)));
                const accuracy = Math.round(((characterIndex - typingErrors) / characterIndex) * 100);
                
                wpmResult.innerText = netWpm;
                accResult.innerText = `${accuracy}%`;
                
                if (characterIndex >= text.length) {
                    drillActive = false;
                    showToast(`Drill complete! net speed: ${netWpm} WPM | Acc: ${accuracy}%`);
                    
                    const date = new Date().toISOString().split('T')[0];
                    state.typingLogs.push({ date, wpm: netWpm, accuracy });
                    saveState();
                    updateLogTables();
                }
            }
        }
    });
    
    typingContainer.addEventListener('keyup', (e) => {
        const targetId = getKeyIdFromEvent(e);
        const keyEl = document.getElementById(targetId);
        if (keyEl) {
            keyEl.classList.remove('pressed');
        }
    });
    
    resetBtn?.addEventListener('click', () => {
        if (characterIndex >= 10 && typingStartTime) {
            const timeDiffMin = (new Date() - typingStartTime) / 60000;
            const grossWpm = Math.round((characterIndex / 5) / (timeDiffMin || 0.01));
            const netWpm = Math.max(0, grossWpm - Math.round(typingErrors / (timeDiffMin || 0.01)));
            const accuracy = Math.round(((characterIndex - typingErrors) / characterIndex) * 100);
            
            const date = new Date().toISOString().split('T')[0];
            state.typingLogs.push({ date, wpm: netWpm, accuracy });
            saveState();
            updateLogTables();
            showToast(`Logged attempt before reset: ${netWpm} WPM | Acc: ${accuracy}%`);
        }
        resetDrill();
        typingContainer.focus();
    });
    
    resetDrill();
}

// 8. BPS-16 Document Protocols & Layout Center
function initProtocolCenter() {
    const layoutSelector = document.getElementById('protocol-select');
    const previewContainer = document.getElementById('protocol-preview');
    
    if (!layoutSelector || !previewContainer) return;
    
    const templates = {
        om: `
<div style="text-align: center; margin-bottom: 20px;">
    <strong>GOVERNMENT OF PAKISTAN<br>
    PRESIDENT'S SECRETARIAT (AIWAN-E-SADR)</strong><br>
    ISLAMABAD
</div>
<div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
    <span>No. F. 4(12)/2026-Admn</span>
    <span>Dated: July 18, 2026</span>
</div>
<hr style="border: 0.5px solid #000; margin-bottom: 20px;">
<div style="text-align: center; margin-bottom: 20px;">
    <u><strong>OFFICE MEMORANDUM</strong></u>
</div>
<p style="text-align: justify; margin-bottom: 15px; text-indent: 40px;">
    Subject: <strong>AMENDMENT IN INSTRUCTIONS REGARDING ADMINISTRATIVE SECURITY COMPLIANCE</strong>
</p>
<p style="text-align: justify; margin-bottom: 15px; text-indent: 40px;">
    The undersigned is directed to refer to the subject noted above and to state that the competent authority has been pleased to approve the updated security protocols for all private secretaries and assistants operating terminal stations in the Secretariat.
</p>
<p style="text-align: justify; margin-bottom: 25px; text-indent: 40px;">
    All administrative departments are requested to bring these instructions to the attention of all staff concerned for strict compliance.
</p>
<div style="float: right; text-align: center; width: 250px; margin-top: 30px;">
    <strong>( ABDUL HALEEM ASGHAR )</strong><br>
    Section Officer (Admn)<br>
    Ph: 051-9200000
</div>
<div style="clear: both; margin-top: 50px;">
    To,<br>
    All Ministries / Divisions of Federal Government.
</div>
        `,
        uo: `
<div style="text-align: center; margin-bottom: 20px;">
    <strong>PRESIDENT'S SECRETARIAT (AIWAN-E-SADR)</strong>
</div>
<div style="margin-bottom: 20px;">
    Subject: <u>Opinion on File No. F. 8(3)-Law/2026</u>
</div>
<hr style="border: 0.5px solid #000; margin-bottom: 20px;">
<p style="text-align: justify; margin-bottom: 15px;">
    Reference Ministry of Law and Justice U.O. Note No. 12/2026-Law, dated 10th July 2026 on the subject mentioned above.
</p>
<p style="text-align: justify; margin-bottom: 20px;">
    2. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; The case file has been examined at this Secretariat. The administrative proposals are in order. It is suggested that approval from the Joint Secretary (Legal affairs) be obtained before finalizing the notification parameters.
</p>
<div style="float: right; text-align: center; width: 250px; margin-top: 40px;">
    <strong>( MUHAMMAD ASHRAF )</strong><br>
    Assistant Private Secretary<br>
    Tel: 051-9211111
</div>
<div style="clear: both; margin-top: 60px;">
    Ministry of Law and Justice, (Mr. Jameel Ahmad, Deputy Secretary), Islamabad.<br>
    -------------------------------------------------------------------------------------<br>
    President's Secretariat (Aiwan-e-Sadr) U.O. No. F.1(4)-APS/2026 &nbsp;&nbsp;&nbsp;&nbsp; Dated: 18-07-2026
</div>
        `,
        oo: `
<div style="text-align: center; margin-bottom: 20px;">
    <strong>GOVERNMENT OF PAKISTAN<br>
    PRESIDENT'S SECRETARIAT (AIWAN-E-SADR)</strong><br>
    ISLAMABAD
</div>
<div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
    <span>No. F. 9(22)/2026-Estt</span>
    <span>Dated: July 18, 2026</span>
</div>
<hr style="border: 0.5px solid #000; margin-bottom: 20px;">
<div style="text-align: center; margin-bottom: 20px;">
    <u><strong>OFFICE ORDER</strong></u>
</div>
<p style="text-align: justify; margin-bottom: 15px; text-indent: 40px;">
    The Competent Authority has been pleased to order the transfer of Mr. Zaheer Ahmad, Assistant Private Secretary (BPS-16), from the Administration Wing to the Coordination Wing with immediate effect and until further orders.
</p>
<p style="text-align: justify; margin-bottom: 25px; text-indent: 40px;">
    He will report to the Deputy Secretary (Coordination) immediately. No transfer grant or traveling allowance is admissible as this transfer is ordered on administrative grounds.
</p>
<div style="float: right; text-align: center; width: 250px; margin-top: 30px;">
    <strong>( SHAFQAT MEHMOOD )</strong><br>
    Deputy Secretary (Establishment)<br>
    Ph: 051-9222222
</div>
<div style="clear: both; margin-top: 50px;">
    Copy to:<br>
    1. Accountant General Pakistan Revenues (AGPR), Islamabad.<br>
    2. Officer Concerned.<br>
    3. PS to Joint Secretary (Admn).<br>
    4. Guard File.
</div>
        `
    };
    
    layoutSelector.addEventListener('change', () => {
        const val = layoutSelector.value;
        if (templates[val]) {
            previewContainer.innerHTML = templates[val];
        }
    });
    
    // Set initial template
    previewContainer.innerHTML = templates.om;
}

// 9. Countdown to Skills & Written Test (40-Day Target)
function updateCountdown() {
    const completedCount = state.completedDays ? state.completedDays.length : 0;
    const daysLeft = Math.max(0, 40 - completedCount);
    
    const countdownEl = document.getElementById('countdown-days');
    if (countdownEl) {
        countdownEl.innerText = daysLeft;
    }
}

// Helper to show modern status notifications
function showToast(message) {
    let toast = document.getElementById('app-toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'app-toast';
        toast.className = 'toast-msg';
        document.body.appendChild(toast);
    }
    
    toast.innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent-gold)" stroke-width="2">
            <circle cx="12" cy="12" r="10"></circle>
            <line x1="12" y1="16" x2="12" y2="12"></line>
            <line x1="12" y1="8" x2="12.01" y2="8"></line>
        </svg>
        <span>${message}</span>
    `;
    
    toast.classList.add('show');
    
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3500);
}

/* ==========================================================================
   100 MARKS WRITTEN EXAM SIMULATOR & QUESTION DATABASE (BPS-16)
   ========================================================================== */

const WRITTEN_EXAM_DB = [
    // ----------------------------------------------------------------------
    // 1. MS OFFICE & COMPUTER KNOWLEDGE (30 MARKS: Questions 1 to 30)
    // ----------------------------------------------------------------------
    {
        id: 1,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which type of break should be inserted in MS Word to change page orientation from Portrait to Landscape for a single page only?",
        options: ["Page Break (Ctrl+Enter)", "Section Break (Next Page)", "Column Break", "Continuous Break"],
        answer: 1,
        explanation: "Section Breaks (Next Page) create independent document formatting zones, allowing different page orientations, margins, or header/footer configurations on specific pages."
    },
    {
        id: 2,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "In MS Excel, which character is used to lock row or column references (creating an absolute cell reference)?",
        options: ["#", "$", "&", "@"],
        answer: 1,
        explanation: "The dollar sign ($) locks the column letter or row number (e.g., $A$1) so references do not change when formulas are copied across cells."
    },
    {
        id: 3,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the keyboard shortcut to trigger a manual Page Break in MS Word?",
        options: ["Shift + Enter", "Ctrl + Enter", "Alt + Enter", "Ctrl + Shift + Enter"],
        answer: 1,
        explanation: "Ctrl + Enter immediately pushes text following the cursor to the top of the next page."
    },
    {
        id: 4,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which MS Excel function counts only non-empty cells containing text, numbers, or logical values in a given range?",
        options: ["COUNT", "COUNTA", "COUNTBLANK", "COUNTIF"],
        answer: 1,
        explanation: "COUNTA counts all non-empty cells in a range regardless of data type, whereas COUNT only counts numeric cells."
    },
    {
        id: 5,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "In PowerPoint, which view is used to apply global changes to font styles, headers, and logos across all slides in a presentation?",
        options: ["Slide Sorter View", "Reading View", "Slide Master View", "Notes Page View"],
        answer: 2,
        explanation: "Slide Master View defines the master hierarchy for formatting, background graphics, and slide layouts across the entire presentation deck."
    },
    {
        id: 6,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the standard function key to launch a PowerPoint presentation slideshow from the very first slide?",
        options: ["F1", "F5", "Shift + F5", "F11"],
        answer: 1,
        explanation: "F5 starts the presentation from Slide 1. Shift + F5 starts the slideshow from the currently selected slide."
    },
    {
        id: 7,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which shortcut key sequence opens the 'Find and Replace' dialog box in MS Word?",
        options: ["Ctrl + F", "Ctrl + R", "Ctrl + H", "Ctrl + Shift + F"],
        answer: 2,
        explanation: "Ctrl + H opens the Replace tab directly in the Find and Replace dialog window."
    },
    {
        id: 8,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "In MS Word Mail Merge, what are the placeholders enclosed in double chevrons (<< Name >>) called?",
        options: ["Merge Fields", "Database Tags", "Rule Operators", "Data Anchors"],
        answer: 0,
        explanation: "Merge Fields act as dynamic data holders linked to columns in an Excel sheet or Access table during Mail Merge."
    },
    {
        id: 9,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the keyboard shortcut to toggle 'Track Changes' on or off in MS Word?",
        options: ["Ctrl + Alt + T", "Ctrl + Shift + E", "Alt + Shift + C", "Ctrl + T"],
        answer: 1,
        explanation: "Ctrl + Shift + E enables or disables tracked revisions and editorial comments in MS Word."
    },
    {
        id: 10,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "In MS Excel, what keyboard shortcut inserts the current date into the selected cell?",
        options: ["Ctrl + ;", "Ctrl + Shift + ;", "Alt + Shift + D", "Ctrl + D"],
        answer: 0,
        explanation: "Ctrl + semicolon (;) inserts the current system date as a static value. Ctrl + Shift + semicolon inserts current time."
    },
    {
        id: 11,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which memory component is non-volatile and contains firmware instructions required to boot up a computer system (BIOS)?",
        options: ["RAM", "ROM", "Cache Memory", "Virtual Memory"],
        answer: 1,
        explanation: "ROM (Read-Only Memory) retains its contents permanently even when the computer is powered off."
    },
    {
        id: 12,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What does URL stand for in computer networking and web terminology?",
        options: ["Universal Resource Locator", "Uniform Resource Locator", "Unified Routing Location", "Universal Remote Link"],
        answer: 1,
        explanation: "URL stands for Uniform Resource Locator, the global address format for web documents and resources."
    },
    {
        id: 13,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which network protocol is primarily responsible for retrieving incoming emails from a remote server to a local mail client?",
        options: ["SMTP", "POP3", "FTP", "DHCP"],
        answer: 1,
        explanation: "POP3 (Post Office Protocol 3) or IMAP is used to retrieve email. SMTP is used exclusively to send emails."
    },
    {
        id: 14,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the default paragraph alignment for official letters and notes written in government secretariats?",
        options: ["Align Left", "Center", "Justify (Ctrl+J)", "Align Right"],
        answer: 2,
        explanation: "Justify alignment distributes text evenly between left and right margins, producing clean, professional document edges."
    },
    {
        id: 15,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which key is used to edit the content of an active cell in MS Excel without overwriting existing data?",
        options: ["F2", "F4", "F7", "F12"],
        answer: 0,
        explanation: "Pressing F2 places the insertion point cursor inside the active cell at the end of existing text."
    },
    {
        id: 16,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the maximum zoom percentage supported in MS Word document window?",
        options: ["200%", "400%", "500%", "1000%"],
        answer: 2,
        explanation: "MS Word supports a zoom range from a minimum of 10% to a maximum of 500%."
    },
    {
        id: 17,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "In MS Excel formula `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`, what does `range_lookup = FALSE` specify?",
        options: ["Approximate Match", "Exact Match", "Case Sensitive Search", "Wildcard Search Only"],
        answer: 1,
        explanation: "Setting range_lookup to FALSE or 0 forces VLOOKUP to search for an exact match of the lookup value."
    },
    {
        id: 18,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which keyboard shortcut applies Subscript formatting (e.g., H₂O) to selected text in MS Word?",
        options: ["Ctrl + =", "Ctrl + Shift + +", "Alt + Shift + S", "Ctrl + Alt + B"],
        answer: 0,
        explanation: "Ctrl + = toggles Subscript formatting. Ctrl + Shift + + toggles Superscript formatting."
    },
    {
        id: 19,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the total number of rows available in a single worksheet in MS Excel 2016 / 2019 / 365?",
        options: ["65,536", "104,857", "1,048,576", "2,097,152"],
        answer: 2,
        explanation: "Modern MS Excel grid contains exactly 1,048,576 rows and 16,384 columns (XFD)."
    },
    {
        id: 20,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which MS Excel function cleans extra spaces from text strings except for single spaces between words?",
        options: ["CLEAN", "TRIM", "REMOVE", "COMPACT"],
        answer: 1,
        explanation: "TRIM removes leading, trailing, and repeated spaces within text strings."
    },
    {
        id: 21,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which shortcut key turns formatting marks (Show/Hide ¶) on or off in MS Word?",
        options: ["Ctrl + Shift + 8", "Ctrl + Alt + 8", "Alt + Shift + 8", "Ctrl + F8"],
        answer: 0,
        explanation: "Ctrl + Shift + 8 (or Ctrl + *) toggles non-printing characters such as space dots, tab arrows, and paragraph marks."
    },
    {
        id: 22,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What feature in MS Excel allows keeping row headers visible while scrolling down a massive dataset?",
        options: ["Split Window", "Freeze Panes", "Lock Grid", "Pin Rows"],
        answer: 1,
        explanation: "Freeze Panes (View > Freeze Panes) locks selected rows or columns in place while navigating the rest of the worksheet."
    },
    {
        id: 23,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the standard file extension for MS PowerPoint presentation templates?",
        options: [".pptx", ".potx", ".ppsx", ".pptm"],
        answer: 1,
        explanation: ".potx is the extension for PowerPoint XML templates used to spawn new pre-styled decks."
    },
    {
        id: 24,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which keyboard shortcut repeats the last formatting action or command in MS Word / Excel?",
        options: ["F2", "F4 (or Ctrl+Y)", "F7", "F11"],
        answer: 1,
        explanation: "F4 or Ctrl + Y repeats or re-applies the previous editing action."
    },
    {
        id: 25,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "In database management systems (DBMS), what uniquely identifies each record in a relational table?",
        options: ["Foreign Key", "Primary Key", "Composite Key", "Index Tag"],
        answer: 1,
        explanation: "A Primary Key contains unique non-null values that uniquely distinguish every record row in a table."
    },
    {
        id: 26,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which feature in MS Word automatically converts typed text like (c) to © or replaces common misspellings?",
        options: ["AutoCorrect", "AutoText", "SmartTags", "Quick Parts"],
        answer: 0,
        explanation: "AutoCorrect automatically corrects typos, capitalizations, and replaces predefined symbols during typing."
    },
    {
        id: 27,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the shortcut to select non-adjacent cells or ranges in an Excel spreadsheet?",
        options: ["Shift + Left Click", "Ctrl + Left Click", "Alt + Left Click", "Tab + Click"],
        answer: 1,
        explanation: "Holding Ctrl while clicking allows selecting disjointed individual cells or separate ranges across a worksheet."
    },
    {
        id: 28,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "Which view in MS Excel displays page boundaries, header/footer text areas, and margin rulers directly on screen?",
        options: ["Normal View", "Page Layout View", "Page Break Preview", "Custom View"],
        answer: 1,
        explanation: "Page Layout View shows the sheet as it will appear on paper, including headers, footers, and page margins."
    },
    {
        id: 29,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "What is the key combination to instantly lock your Windows computer session when stepping away from your secretariat desk?",
        options: ["Win + L", "Ctrl + Alt + L", "Alt + F4", "Win + D"],
        answer: 0,
        explanation: "Windows key + L immediately locks your Windows workstation session to ensure administrative file security."
    },
    {
        id: 30,
        section: "computer",
        sectionName: "MS Office & Computer Knowledge",
        question: "In MS Word, which alignment feature aligns numbers by their decimal points in official financial statements?",
        options: ["Left Tab", "Decimal Tab", "Bar Tab", "Center Tab"],
        answer: 1,
        explanation: "Decimal Tab aligns numbers so that decimal points line up vertically regardless of digit lengths."
    },

    // ----------------------------------------------------------------------
    // 2. ENGLISH LANGUAGE & GRAMMAR (20 MARKS: Questions 31 to 50)
    // ----------------------------------------------------------------------
    {
        id: 31,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Choose the correct preposition: 'The accused was acquitted ___ all charges leveled against him.'",
        options: ["from", "of", "with", "by"],
        answer: 1,
        explanation: "The verb 'acquitted' takes the preposition 'of' (e.g., acquitted of theft/charges)."
    },
    {
        id: 32,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "What is the closest SYNONYM of the word 'Meticulous'?",
        options: ["Careless", "Painstaking / Thorough", "Hasty", "Indifferent"],
        answer: 1,
        explanation: "Meticulous means showing great attention to detail, highly careful and painstaking."
    },
    {
        id: 33,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "What is the exact ANTONYM of the word 'Transient'?",
        options: ["Fleeting", "Permanent / Enduring", "Temporary", "Momentary"],
        answer: 1,
        explanation: "Transient means lasting only for a short time; its direct opposite is permanent."
    },
    {
        id: 34,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Convert to Passive Voice: 'The Assistant Private Secretary typed the official notification.'",
        options: [
            "The official notification is typed by the Assistant Private Secretary.",
            "The official notification was typed by the Assistant Private Secretary.",
            "The official notification had been typed by the Assistant Private Secretary.",
            "The official notification was being typed by the Assistant Private Secretary."
        ],
        answer: 1,
        explanation: "Past simple active ('typed') changes to 'was typed' in passive voice."
    },
    {
        id: 35,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Convert to Indirect Speech: He said, 'I am preparing for the BPS-16 test today.'",
        options: [
            "He said that he is preparing for the BPS-16 test today.",
            "He said that he was preparing for the BPS-16 test that day.",
            "He told that he had been preparing for the BPS-16 test today.",
            "He said he will prepare for the BPS-16 test that day."
        ],
        answer: 1,
        explanation: "Present continuous ('am preparing') converts to past continuous ('was preparing'), and 'today' changes to 'that day'."
    },
    {
        id: 36,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "What is the meaning of the idiom 'To burn the midnight oil'?",
        options: ["To waste resources", "To work or study late into the night", "To start a fire", "To ruin a project"],
        answer: 1,
        explanation: "To burn the midnight oil means to study or work far into the night."
    },
    {
        id: 37,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Fill in the blank with correct subject-verb agreement: 'Neither of the candidate officers ___ present at the hearing.'",
        options: ["were", "was", "are", "have been"],
        answer: 1,
        explanation: "Indefinite pronouns like 'neither', 'either', and 'each' take singular verbs ('was')."
    },
    {
        id: 38,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Identify the correctly spelled word:",
        options: ["Bureaucracy", "Beurocracy", "Bureacracy", "Beureaucracy"],
        answer: 0,
        explanation: "'Bureaucracy' (B-U-R-E-A-U-C-R-A-C-Y) is the correct spelling."
    },
    {
        id: 39,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Choose the correct one-word substitution: 'An official post or office for which no salary is paid'",
        options: ["Honorary", "Sinecure", "Mercenary", "Autopsy"],
        answer: 0,
        explanation: "An honorary position is held without payment. (A sinecure is a paid position with minimal work)."
    },
    {
        id: 40,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Choose the correct preposition: 'He has been working in the Secretariat ___ 2018.'",
        options: ["for", "since", "from", "during"],
        answer: 1,
        explanation: "'Since' is used with a specific starting point in time (2018). 'For' is used for a duration/period of time."
    },
    {
        id: 41,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Identify the error in sentence: 'He is senior than me in service.'",
        options: ["'He is'", "'senior than'", "'me in'", "'service'"],
        answer: 1,
        explanation: "Latin comparative adjectives (senior, junior, superior, inferior, prior) take 'to', not 'than'. Correct: 'senior to me'."
    },
    {
        id: 42,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "What is the closest SYNONYM of the word 'Candid'?",
        options: ["Secretive", "Frank / Outspoken", "Deceitful", "Cautious"],
        answer: 1,
        explanation: "Candid means truthful, straightforward, and frank."
    },
    {
        id: 43,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Fill in the blank: 'If he had submitted the file on time, the Secretary ___ it.'",
        options: ["would approve", "would have approved", "will approve", "had approved"],
        answer: 1,
        explanation: "Third conditional structure: 'If + past perfect (had submitted) ... would have + past participle (would have approved)'."
    },
    {
        id: 44,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "What is the plural form of the word 'Crisis'?",
        options: ["Crisises", "Crises", "Crisii", "Crisiss"],
        answer: 1,
        explanation: "Words ending in -is derived from Greek form plurals by changing -is to -es ('Crises')."
    },
    {
        id: 45,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "What is the ANTONYM of the word 'Ameliorate'?",
        options: ["Improve", "Worsen / Aggravate", "Enhance", "Mitigate"],
        answer: 1,
        explanation: "Ameliorate means to make something better; its opposite is worsen or aggravate."
    },
    {
        id: 46,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Identify the part of speech of the underlined word: 'He typed the report extremely FAST.'",
        options: ["Adjective", "Adverb", "Noun", "Conjunction"],
        answer: 1,
        explanation: "'Fast' modifies the verb 'typed', describing how he typed, so it functions as an Adverb."
    },
    {
        id: 47,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Choose the correct meaning of the phrase 'In lieu of':",
        options: ["In spite of", "In place of / Instead of", "In favor of", "In compliance with"],
        answer: 1,
        explanation: "'In lieu of' is a prepositional phrase meaning instead of or in place of."
    },
    {
        id: 48,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Identify the correctly spelled word:",
        options: ["Accomodation", "Accommodation", "Acommodation", "Accomodasion"],
        answer: 1,
        explanation: "'Accommodation' contains double 'c' and double 'm' (A-C-C-O-M-M-O-D-A-T-I-O-N)."
    },
    {
        id: 49,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "Choose the correct preposition: 'She presided ___ the annual departmental committee meeting.'",
        options: ["at", "over", "on", "upon"],
        answer: 1,
        explanation: "The verb 'preside' takes the preposition 'over' when conducting or chairing a meeting."
    },
    {
        id: 50,
        section: "english",
        sectionName: "English Language & Grammar",
        question: "What is the meaning of the idiom 'A white elephant'?",
        options: ["A rare valuable item", "A costly but useless possession", "An ancient monument", "A symbol of peace"],
        answer: 1,
        explanation: "A white elephant refers to a possession that is very expensive to maintain and yields no practical utility."
    },

    // ----------------------------------------------------------------------
    // 3. GENERAL KNOWLEDGE & WORLD AFFAIRS (15 MARKS: Questions 51 to 65)
    // ----------------------------------------------------------------------
    {
        id: 51,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Where are the official headquarters of the United Nations Organization (UNO) located?",
        options: ["Geneva, Switzerland", "New York, USA", "London, UK", "Paris, France"],
        answer: 1,
        explanation: "UN Headquarters was established in New York City, USA, along the East River in 1952."
    },
    {
        id: 52,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Which is the largest and deepest ocean on Earth?",
        options: ["Atlantic Ocean", "Indian Ocean", "Pacific Ocean", "Arctic Ocean"],
        answer: 2,
        explanation: "The Pacific Ocean covers over 30% of the Earth's surface and contains the deepest trench (Mariana Trench)."
    },
    {
        id: 53,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "What is the official currency of Japan?",
        options: ["Yuan", "Yen", "Won", "Ringgit"],
        answer: 1,
        explanation: "The official currency unit of Japan is the Japanese Yen (¥)."
    },
    {
        id: 54,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Where is the Secretariat of the Organization of Islamic Cooperation (OIC) located?",
        options: ["Riyadh, Saudi Arabia", "Jeddah, Saudi Arabia", "Cairo, Egypt", "Istanbul, Turkey"],
        answer: 1,
        explanation: "The OIC General Secretariat is situated in Jeddah, Kingdom of Saudi Arabia."
    },
    {
        id: 55,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Which is the deepest freshwater lake in the world?",
        options: ["Lake Superior", "Lake Baikal", "Lake Victoria", "Caspian Sea"],
        answer: 1,
        explanation: "Lake Baikal in Siberia, Russia, is the world's deepest (1,642m) and most voluminous freshwater lake."
    },
    {
        id: 56,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Who authored the famous classical economic treatise 'The Wealth of Nations' (1776)?",
        options: ["Karl Marx", "Adam Smith", "John Maynard Keynes", "David Ricardo"],
        answer: 1,
        explanation: "Scottish philosopher Adam Smith wrote 'The Wealth of Nations', laying the foundations of modern free-market economics."
    },
    {
        id: 57,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Which country is popularly known as the 'Land of the Rising Sun'?",
        options: ["China", "Japan", "Thailand", "Norway"],
        answer: 1,
        explanation: "Japan is called 'Nippon' or 'Nihon' in Japanese, meaning 'sun origin' or Land of the Rising Sun."
    },
    {
        id: 58,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Where is the headquarters of the World Health Organization (WHO) situated?",
        options: ["Geneva, Switzerland", "Vienna, Austria", "Rome, Italy", "Washington D.C., USA"],
        answer: 0,
        explanation: "WHO was founded in 1948 with its global headquarters in Geneva, Switzerland."
    },
    {
        id: 59,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "The artificial Suez Canal connects which two major bodies of water?",
        options: ["Black Sea and Caspian Sea", "Mediterranean Sea and Red Sea", "Red Sea and Arabian Sea", "Atlantic Ocean and Pacific Ocean"],
        answer: 1,
        explanation: "Opened in 1869, the Suez Canal in Egypt connects the Mediterranean Sea directly to the Red Sea."
    },
    {
        id: 60,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "What is the official standard time zone of Pakistan relative to Coordinated Universal Time (UTC)?",
        options: ["UTC+4", "UTC+5", "UTC+6", "UTC+3:30"],
        answer: 1,
        explanation: "Pakistan Standard Time (PST) is 5 hours ahead of UTC (UTC+05:00)."
    },
    {
        id: 61,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "How many sovereign member nations belong to the South Asian Association for Regional Cooperation (SAARC)?",
        options: ["6", "7", "8", "10"],
        answer: 2,
        explanation: "SAARC has 8 member states: Pakistan, India, Bangladesh, Sri Lanka, Nepal, Bhutan, Maldives, and Afghanistan."
    },
    {
        id: 62,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Which is the smallest continent on Earth by total land area?",
        options: ["Europe", "Antarctica", "Australia (Oceania)", "South America"],
        answer: 2,
        explanation: "Australia is the smallest continent by land surface area (~7.69 million sq km)."
    },
    {
        id: 63,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Which gas is most abundant by volume in the Earth's atmosphere?",
        options: ["Oxygen (~21%)", "Nitrogen (~78%)", "Carbon Dioxide", "Argon"],
        answer: 1,
        explanation: "Nitrogen comprises approximately 78.08% of Earth's atmospheric volume."
    },
    {
        id: 64,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "Where is the principal judicial organ of the UN, the International Court of Justice (ICJ), seated?",
        options: ["Geneva, Switzerland", "The Hague, Netherlands", "New York, USA", "Brussels, Belgium"],
        answer: 1,
        explanation: "The ICJ is headquartered at the Peace Palace in The Hague, Netherlands."
    },
    {
        id: 65,
        section: "gk",
        sectionName: "General Knowledge & World Affairs",
        question: "What is the SI unit of electrical resistance?",
        options: ["Volt", "Ampere", "Ohm", "Watt"],
        answer: 2,
        explanation: "The Ohm (symbol Ω) is the SI derived unit of electrical resistance."
    },

    // ----------------------------------------------------------------------
    // 4. PAKISTAN STUDIES & CONSTITUTION (15 MARKS: Questions 66 to 80)
    // ----------------------------------------------------------------------
    {
        id: 66,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "On which date was the historic Lahore Resolution (later known as Pakistan Resolution) passed?",
        options: ["14th August 1947", "23rd March 1940", "12th March 1949", "3rd June 1947"],
        answer: 1,
        explanation: "Passed at Minto Park (now Iqbal Park) Lahore on 23rd March 1940 during the All-India Muslim League session."
    },
    {
        id: 67,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "Who served as the First Governor-General of independent Pakistan?",
        options: ["Liaquat Ali Khan", "Quaid-e-Azam Muhammad Ali Jinnah", "Khawaja Nazimuddin", "Malik Ghulam Muhammad"],
        answer: 1,
        explanation: "Quaid-e-Azam Muhammad Ali Jinnah took oath as the first Governor-General on 15th August 1947."
    },
    {
        id: 68,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "The current operative Constitution of the Islamic Republic of Pakistan was enacted in which year?",
        options: ["1956", "1962", "1973", "1985"],
        answer: 2,
        explanation: "Passed unanimously by the National Assembly and authenticated on 14th August 1973 under Zulfikar Ali Bhutto."
    },
    {
        id: 69,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "What is the height of K2 (Mount Godwin-Austen), the highest mountain peak in Pakistan?",
        options: ["8,126 meters", "8,611 meters", "8,848 meters", "7,788 meters"],
        answer: 1,
        explanation: "K2 in the Karakoram range stands at 8,611 meters (28,251 ft), the second highest peak on Earth."
    },
    {
        id: 70,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "When was the landmark Objective Resolution passed by the Constituent Assembly of Pakistan?",
        options: ["23rd March 1940", "12th March 1949", "14th August 1947", "23rd March 1956"],
        answer: 1,
        explanation: "Moved by Prime Minister Liaquat Ali Khan, the Objective Resolution was adopted on 12th March 1949."
    },
    {
        id: 71,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "Under the 1973 Constitution of Pakistan, what system of government is established at the federal level?",
        options: ["Presidential System", "Federal Parliamentary System", "Unitary System", "Confederation"],
        answer: 1,
        explanation: "The 1973 Constitution establishes a bicameral Federal Parliamentary democracy with Prime Minister as head of government."
    },
    {
        id: 72,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "Following the 25th Constitutional Amendment (FATA Merger), what is the total general seat membership of the National Assembly of Pakistan?",
        options: ["342", "336", "104", "100"],
        answer: 1,
        explanation: "Post-FATA merger and redistributions, total seats in National Assembly were reduced to 336 (266 general + 60 women + 10 non-Muslims)."
    },
    {
        id: 73,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "What is the national flower of Pakistan?",
        options: ["Rose", "Jasmine (Chambeli)", "Tulip", "Sunflower"],
        answer: 1,
        explanation: "Jasmine (Jasminum officinale) is the official national flower of Pakistan."
    },
    {
        id: 74,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "In which year did Pakistan adopt its first constitution, officially becoming an 'Islamic Republic'?",
        options: ["1947", "1956", "1962", "1973"],
        answer: 1,
        explanation: "The 1956 Constitution (enacted 23rd March 1956) proclaimed Pakistan as an Islamic Republic."
    },
    {
        id: 75,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "What is the name of the international border line demarcated between Pakistan and Afghanistan in 1893?",
        options: ["Radcliffe Line", "Durand Line", "McMahon Line", "Line of Control"],
        answer: 1,
        explanation: "The 2,640 km border agreement was signed between Sir Mortimer Durand and Afghan Amir Abdur Rahman Khan in 1893."
    },
    {
        id: 76,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "Who is the Chief Executive and Head of Government under the 1973 Constitution of Pakistan?",
        options: ["The President", "The Prime Minister", "The Chief Justice", "The Speaker National Assembly"],
        answer: 1,
        explanation: "Article 90 vests executive authority of the Federation in the Prime Minister as head of government."
    },
    {
        id: 77,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "Who was appointed as the first Chief Justice of Pakistan upon independence?",
        options: ["Justice M.R. Kayani", "Justice Sir Abdul Rashid", "Justice Alvin Robert Cornelius", "Justice Muhammad Munir"],
        answer: 1,
        explanation: "Sir Mian Abdul Rashid served as the inaugural Chief Justice of Pakistan from June 1949 to June 1954."
    },
    {
        id: 78,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "Tarbela Dam, one of the world's largest earth-filled dams, is constructed across which river?",
        options: ["Jhelum River", "Chenab River", "Indus River", "Sutlej River"],
        answer: 2,
        explanation: "Tarbela Dam is located on the Indus River in Haripur District, Khyber Pakhtunkhwa."
    },
    {
        id: 79,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "In response to which constitutional report did Quaid-e-Azam present his famous '14 Points' in March 1929?",
        options: ["Simon Commission Report", "Nehru Report (1928)", "Cabinet Mission Plan", "Cripps Proposal"],
        answer: 1,
        explanation: "Quaid-e-Azam's 14 Points were formulated at Delhi in March 1929 to safeguard Muslim rights against the Nehru Report."
    },
    {
        id: 80,
        section: "pakstudy",
        sectionName: "Pakistan Studies & Constitution",
        question: "In the Federal Secretariat structure, under the Rules of Business 1973, what is the primary role of BPS-16 cadre (Assistant / Assistant Private Secretary)?",
        options: [
            "Policy approval and signing international treaties",
            "Supervising section diary, drafting routine correspondence, taking dictation, and managing file records",
            "Auditing provincial expenditures",
            "Appointing judicial magistrates"
        ],
        answer: 1,
        explanation: "BPS-16 personnel form the administrative baseline responsible for shorthand dictation, typing, file movement registers, and drafting."
    },

    // ----------------------------------------------------------------------
    // 5. ISLAMIAT & ISLAMIC HISTORY (10 MARKS: Questions 81 to 90)
    // ----------------------------------------------------------------------
    {
        id: 81,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "Which were the first verses of the Holy Quran revealed to Prophet Muhammad (PBUH) in Cave Hira?",
        options: ["First 5 verses of Surah Al-Fatiha", "First 5 verses of Surah Al-Alaq (Iqra)", "First 10 verses of Surah Al-Baqarah", "Surah Yaseen"],
        answer: 1,
        explanation: "The first divine revelation brought by Angel Jibrael (A.S.) comprised the first 5 verses of Surah Al-Alaq."
    },
    {
        id: 82,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "How many total Surahs (chapters) are contained in the Holy Quran?",
        options: ["86", "114", "6666", "30"],
        answer: 1,
        explanation: "The Holy Quran consists of 114 Surahs (86 Makki and 28 Madani)."
    },
    {
        id: 83,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "In which Hijri year did the decisive Ghazwa-e-Badr take place?",
        options: ["1 A.H.", "2 A.H.", "3 A.H.", "5 A.H."],
        answer: 1,
        explanation: "The Battle of Badr took place on 17th Ramadan, 2 A.H. (624 AD), resulting in victory for Muslims."
    },
    {
        id: 84,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "Who was chosen as the First Caliph of Islam following the demise of Prophet Muhammad (PBUH)?",
        options: ["Hazrat Umar Farooq (R.A.)", "Hazrat Abu Bakr Siddique (R.A.)", "Hazrat Usman Ghani (R.A.)", "Hazrat Ali Murtaza (R.A.)"],
        answer: 1,
        explanation: "Hazrat Abu Bakr Siddique (R.A.) served as the 1st Rashidun Caliph from 632 to 634 AD."
    },
    {
        id: 85,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "Which fundamental Pillar of Islam is referred to in Hadith as a 'Shield against sin and hellfire'?",
        options: ["Salah (Prayer)", "Sawm / Roza (Fasting)", "Zakat (Almsgiving)", "Hajj (Pilgrimage)"],
        answer: 1,
        explanation: "Prophet Muhammad (PBUH) stated: 'Fasting is a shield (Junnah) with which a servant protects himself from the Fire.'"
    },
    {
        id: 86,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "Which is the longest Surah in the Holy Quran?",
        options: ["Surah Al-Imran", "Surah Al-Baqarah", "Surah An-Nisa", "Surah Yaseen"],
        answer: 1,
        explanation: "Surah Al-Baqarah (Chapter 2) contains 286 verses and is the longest Surah in the Holy Quran."
    },
    {
        id: 87,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "In which Hijri year was the historic Treaty of Hudaibiya signed between Muslims and Quraish?",
        options: ["4 A.H.", "6 A.H.", "8 A.H.", "10 A.H."],
        answer: 1,
        explanation: "Sulh al-Hudaibiyyah was concluded in 6 A.H. (628 AD), recognized by Quran as 'Fath-um-Mubeen' (a manifest victory)."
    },
    {
        id: 88,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "What was the name of the respected mother of Holy Prophet Muhammad (PBUH)?",
        options: ["Hazrat Halima (R.A.)", "Hazrat Amina (R.A.)", "Hazrat Khadija (R.A.)", "Hazrat Fatima (R.A.)"],
        answer: 1,
        explanation: "Hazrat Amina bint Wahb (R.A.) was the mother of Prophet Muhammad (PBUH)."
    },
    {
        id: 89,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "In which Hijri year did the bloodless Conquest of Makkah (Fath Makkah) occur?",
        options: ["6 A.H.", "8 A.H.", "9 A.H.", "10 A.H."],
        answer: 1,
        explanation: "The Conquest of Makkah took place in Ramadan, 8 A.H. (630 AD)."
    },
    {
        id: 90,
        section: "islamiat",
        sectionName: "Islamiat & Islamic History",
        question: "What is the minimum threshold (Nisab) of Gold for Zakat obligation?",
        options: ["5 Tolas", "7.5 Tolas (87.48 grams)", "52.5 Tolas", "10 Tolas"],
        answer: 1,
        explanation: "Nisab for Gold is 7.5 Tolas (87.48 grams). For Silver, it is 52.5 Tolas (612.36 grams)."
    },

    // ----------------------------------------------------------------------
    // 6. GENERAL MATHEMATICS & ANALYTICAL ABILITY (10 MARKS: Questions 91 to 100)
    // ----------------------------------------------------------------------
    {
        id: 91,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "If 20% of a number is equal to 40, what is 50% of that same number?",
        options: ["80", "100", "120", "200"],
        answer: 1,
        explanation: "Let number be X. 0.20 * X = 40 => X = 200. Therefore, 50% of 200 = 100."
    },
    {
        id: 92,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "What is the arithmetic mean (average) of the numbers: 10, 20, 30, 40, and 50?",
        options: ["25", "30", "35", "40"],
        answer: 1,
        explanation: "Sum = 10+20+30+40+50 = 150. Count = 5. Average = 150 / 5 = 30."
    },
    {
        id: 93,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "Two numbers are in the ratio 2 : 3. If their sum is 50, what is the larger number?",
        options: ["20", "25", "30", "35"],
        answer: 2,
        explanation: "Total parts = 2 + 3 = 5 parts. 1 part = 50 / 5 = 10. Larger number = 3 parts = 3 * 10 = 30."
    },
    {
        id: 94,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "Complete the number series pattern: 2, 4, 8, 16, 32, ___?",
        options: ["48", "64", "56", "128"],
        answer: 1,
        explanation: "Each number is multiplied by 2 (geometric sequence). 32 * 2 = 64."
    },
    {
        id: 95,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "Calculate the Simple Interest on a principal amount of Rs. 1,000 at 10% per annum for 2 years.",
        options: ["Rs. 100", "Rs. 150", "Rs. 200", "Rs. 250"],
        answer: 2,
        explanation: "Simple Interest formula SI = (P * R * T) / 100 = (1000 * 10 * 2) / 100 = Rs. 200."
    },
    {
        id: 96,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "Solve the linear equation for x: `3x + 5 = 20`",
        options: ["x = 3", "x = 5", "x = 6", "x = 15"],
        answer: 1,
        explanation: "3x = 20 - 5 => 3x = 15 => x = 5."
    },
    {
        id: 97,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "Which of the following fractions is largest in value: 1/2, 3/4, 2/3, or 5/8?",
        options: ["1/2 (0.50)", "3/4 (0.75)", "2/3 (0.66)", "5/8 (0.625)"],
        answer: 1,
        explanation: "Converting to decimals: 1/2 = 0.50, 3/4 = 0.75, 2/3 = 0.667, 5/8 = 0.625. Largest is 3/4."
    },
    {
        id: 98,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "If Worker A can complete a task in 6 days and Worker B in 12 days, how many days will they take working together?",
        options: ["3 days", "4 days", "5 days", "9 days"],
        answer: 1,
        explanation: "Combined rate = 1/6 + 1/12 = 3/12 = 1/4 per day. Total time taken = 4 days."
    },
    {
        id: 99,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "What is the sum of interior angles of any Euclidean triangle?",
        options: ["90 degrees", "180 degrees", "270 degrees", "360 degrees"],
        answer: 1,
        explanation: "The interior angles of any triangle always add up to 180 degrees."
    },
    {
        id: 100,
        section: "math",
        sectionName: "General Math & Analytical Ability",
        question: "If the price of a official handbook increases from Rs. 80 to Rs. 100, what is the percentage increase?",
        options: ["20%", "25%", "30%", "15%"],
        answer: 1,
        explanation: "Increase = 100 - 80 = 20. Percentage Increase = (20 / 80) * 100 = 25%."
    }
];

// Active Written Exam Runtime State
let writtenExamState = {
    active: false,
    mode: 'full',
    subjectKey: null,
    questions: [],
    userAnswers: [],
    flagged: new Set(),
    currentIndex: 0,
    timerSeconds: 5400,
    timerInterval: null,
    startTime: null
};

// Initialize Written Exam Event Listeners
function initWrittenExam() {
    const btnStartFull = document.getElementById('btn-start-full-exam');
    if (btnStartFull) {
        btnStartFull.addEventListener('click', () => {
            startWrittenExam('full');
        });
    }

    // Start Daily 8-Question Drill (38-Day Plan)
    const btnStartDaily8 = document.getElementById('btn-start-daily-8');
    if (btnStartDaily8) {
        btnStartDaily8.addEventListener('click', () => {
            startWrittenExam('daily8');
        });
    }

    // Model Papers Buttons (6 Papers)
    const modelPaperBtns = document.querySelectorAll('.model-paper-btn');
    modelPaperBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const paperNum = parseInt(btn.dataset.paper) || 1;
            startWrittenExam('paper', paperNum);
        });
    });

    const drillBtns = document.querySelectorAll('.drill-subject-btn');
    drillBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const subject = btn.dataset.subject;
            startWrittenExam('drill', subject);
        });
    });

    document.getElementById('btn-prev-q')?.addEventListener('click', () => navigateQuestion(-1));
    document.getElementById('btn-next-q')?.addEventListener('click', () => {
        if (writtenExamState.currentIndex === writtenExamState.questions.length - 1) {
            confirmSubmitExam();
        } else {
            navigateQuestion(1);
        }
    });
    document.getElementById('btn-flag-q')?.addEventListener('click', toggleFlagCurrentQuestion);
    document.getElementById('btn-clear-q')?.addEventListener('click', clearCurrentAnswer);
    document.getElementById('btn-submit-exam-early')?.addEventListener('click', confirmSubmitExam);
    document.getElementById('btn-start-test-confirm')?.addEventListener('click', () => {
        document.getElementById('exam-start-setup-card').classList.add('d-none');
        const examGrid = document.querySelector('.exam-main-grid');
        if (examGrid) examGrid.classList.remove('d-none');
        startExamTimer();
        renderCurrentQuestion();
    });
    document.getElementById('btn-quit-exam')?.addEventListener('click', async () => {
        const confirmed = await showFancyConfirm("Exit Test", "Are you sure you want to exit the current test? Your progress will not be saved.", "quit", "Yes, Exit", "No");
        if (confirmed) {
            if (writtenExamState.timerInterval) clearInterval(writtenExamState.timerInterval);
            writtenExamState.active = false;
            resetWrittenExamUI();
            showToast("Test exited.");
        }
    });
    document.getElementById('btn-retake-exam')?.addEventListener('click', resetWrittenExamUI);
    document.getElementById('btn-result-back')?.addEventListener('click', resetWrittenExamUI);

    document.getElementById('btn-review-answers')?.addEventListener('click', () => openAnswerReview('all'));
    document.getElementById('btn-back-to-results')?.addEventListener('click', showResultScreen);

    const reviewFilterBtns = document.querySelectorAll('.review-filter-btn');
    reviewFilterBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            reviewFilterBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderAnswerReviewList(btn.dataset.filter);
        });
    });

    document.getElementById('btn-clear-exam-history')?.addEventListener('click', async () => {
        const confirmed = await showFancyConfirm("Clear History", "Are you sure you want to clear all recorded exam attempts? This cannot be undone.", "quit", "Yes, Clear", "No");
        if (confirmed) {
            localStorage.removeItem('bps16_written_history');
            loadExamHistoryTable();
            showToast("Exam history cleared.");
        }
    });

    // Subject Cards click to open 40-Days selection
    const subjectCards = document.querySelectorAll('.prep-subject-card');
    subjectCards.forEach(card => {
        card.addEventListener('click', () => {
            const subject = card.dataset.subject;
            const subjectName = card.querySelector('.syll-title').innerText;
            
            document.getElementById('selected-subject-title').innerText = subjectName;
            document.getElementById('prep-subject-grid').classList.add('d-none');
            document.getElementById('prep-day-selector').classList.remove('d-none');
            
            // Push history state
            if (!historyStateUpdateInProgress) {
                history.pushState({ type: 'day-selector', subject: subject, subjectName: subjectName }, '', window.location.pathname);
            }

            // Build Days Grid
            const daysGrid = document.getElementById('subject-days-grid');
            if (daysGrid) {
                daysGrid.innerHTML = '';
                for (let d = 1; d <= 40; d++) {
                    const dayBtn = document.createElement('button');
                    dayBtn.className = 'palette-btn';
                    dayBtn.style.padding = '10px';
                    dayBtn.style.fontSize = '0.9rem';
                    dayBtn.innerText = `Day ${d}`;
                    dayBtn.addEventListener('click', () => {
                        startWrittenExam('daily-subject-day', { subject, dayNum: d });
                    });
                    daysGrid.appendChild(dayBtn);
                }
            }
        });
    });

    document.getElementById('btn-back-to-subjects')?.addEventListener('click', () => {
        history.back();
    });

    updateDailyStudyPlanUI();
    loadExamHistoryTable();
}

// Start Written Exam (Full 100 Marks, Model Papers, Subject Drill, or Daily 8 MCQs)
function startWrittenExam(mode, param = null) {
    if (!historyStateUpdateInProgress) {
        history.pushState({ type: 'exam', examMode: mode, param: param }, '', window.location.pathname);
    }

    writtenExamState.mode = mode;
    writtenExamState.subjectKey = (mode === 'drill') ? param : null;
    writtenExamState.active = true;
    writtenExamState.currentIndex = 0;
    writtenExamState.flagged = new Set();
    writtenExamState.startTime = new Date();

    if (mode === 'full') {
        writtenExamState.questions = [...WRITTEN_EXAM_DB];
        writtenExamState.timerSeconds = 90 * 60; // 90 minutes
        document.getElementById('exam-mode-title').innerText = "Full 100 Marks Written Mock Exam";
    } else if (mode === 'paper') {
        const paperNum = param || 1;
        // Shift questions deterministically based on paper number for authentic distinct papers
        const shift = (paperNum - 1) * 15;
        writtenExamState.questions = WRITTEN_EXAM_DB.map((q, idx) => {
            const targetIdx = (idx + shift) % WRITTEN_EXAM_DB.length;
            return { ...WRITTEN_EXAM_DB[targetIdx], id: idx + 1 };
        });
        writtenExamState.timerSeconds = 90 * 60;
        document.getElementById('exam-mode-title').innerText = `Model Paper ${paperNum}: Official BPS-16 Past Paper`;
    } else if (mode === 'daily8') {
        // Pick 8 daily questions based on current day
        const dayNum = state.currentDay || 1;
        const startIdx = ((dayNum - 1) * 8) % WRITTEN_EXAM_DB.length;
        const count = (dayNum === 38) ? 4 : 8;
        writtenExamState.questions = [];
        for (let i = 0; i < count; i++) {
            const q = WRITTEN_EXAM_DB[(startIdx + i) % WRITTEN_EXAM_DB.length];
            writtenExamState.questions.push({ ...q, id: i + 1 });
        }
        writtenExamState.timerSeconds = count * 2 * 60; // 2 mins per question
        document.getElementById('exam-mode-title').innerText = `Daily Study Drill (Day ${dayNum} - ${count} MCQs Target)`;
    } else if (mode === 'daily-subject-day') {
        const subject = param.subject;
        const dayNum = param.dayNum;
        const subjectQuestions = WRITTEN_EXAM_DB.filter(q => q.section === subject);
        
        // Pick exactly 8 questions deterministically for this day & subject
        const startIdx = ((dayNum - 1) * 8) % subjectQuestions.length;
        writtenExamState.questions = [];
        for (let i = 0; i < 8; i++) {
            const q = subjectQuestions[(startIdx + i) % subjectQuestions.length];
            writtenExamState.questions.push({ ...q, id: i + 1 });
        }
        writtenExamState.timerSeconds = 8 * 2 * 60; // 16 minutes
        const subjectTitle = document.getElementById('selected-subject-title').innerText;
        document.getElementById('exam-mode-title').innerText = `${subjectTitle} (Day ${dayNum} - 8 MCQs)`;
    } else {
        writtenExamState.questions = WRITTEN_EXAM_DB.filter(q => q.section === param);
        writtenExamState.timerSeconds = Math.max(15 * 60, writtenExamState.questions.length * 90);
        const sampleQ = writtenExamState.questions[0];
        document.getElementById('exam-mode-title').innerText = `${sampleQ ? sampleQ.sectionName : 'Subject'} Practice Drill`;
    }

    writtenExamState.userAnswers = new Array(writtenExamState.questions.length).fill(null);

    // Update UI Visibility
    document.getElementById('written-mode-select').classList.add('d-none');
    document.getElementById('written-exam-result').classList.add('d-none');
    document.getElementById('written-answer-review').classList.add('d-none');
    document.getElementById('written-exam-active').classList.remove('d-none');

    // Build Palette
    buildQuestionPalette();

    // Show start setup screen and hide active test grid
    const setupCard = document.getElementById('exam-start-setup-card');
    const examGrid = document.querySelector('.exam-main-grid');
    if (setupCard) setupCard.classList.remove('d-none');
    if (examGrid) examGrid.classList.add('d-none');

    // Update setup screen text
    const modeTitle = document.getElementById('exam-mode-title').innerText;
    document.getElementById('setup-exam-title').innerText = modeTitle;
    document.getElementById('setup-exam-info').innerHTML = `
        This test contains <strong>${writtenExamState.questions.length} Questions</strong>. 
        You have <strong>${Math.round(writtenExamState.timerSeconds / 60)} Minutes</strong> to complete it.
    `;

    // Reset initial timer display
    updateTimerDisplay();

    showToast(`Exam loaded! Ready to start.`);
}

// Build Question Palette Buttons (1 to N)
function buildQuestionPalette() {
    const grid = document.getElementById('question-palette-grid');
    if (!grid) return;
    grid.innerHTML = '';

    writtenExamState.questions.forEach((q, idx) => {
        const btn = document.createElement('button');
        btn.className = 'palette-btn';
        btn.innerText = idx + 1;
        btn.id = `palette-btn-${idx}`;
        btn.addEventListener('click', () => {
            writtenExamState.currentIndex = idx;
            renderCurrentQuestion();
        });
        grid.appendChild(btn);
    });

    updatePaletteState();
}

// Update Palette Button Classes
function updatePaletteState() {
    writtenExamState.questions.forEach((q, idx) => {
        const btn = document.getElementById(`palette-btn-${idx}`);
        if (!btn) return;

        btn.classList.remove('active-q', 'answered', 'flagged');

        if (idx === writtenExamState.currentIndex) {
            btn.classList.add('active-q');
        }

        if (writtenExamState.flagged.has(idx)) {
            btn.classList.add('flagged');
        } else if (writtenExamState.userAnswers[idx] !== null) {
            btn.classList.add('answered');
        }
    });

    const answeredCount = writtenExamState.userAnswers.filter(a => a !== null).length;
    const answeredLbl = document.getElementById('answered-counter-label');
    if (answeredLbl) {
        answeredLbl.innerText = `${answeredCount} / ${writtenExamState.questions.length} Answered`;
    }
}

// Render Active Question
function renderCurrentQuestion() {
    const idx = writtenExamState.currentIndex;
    const q = writtenExamState.questions[idx];
    if (!q) return;

    // Header counter and progress
    document.getElementById('q-curr-num').innerText = idx + 1;
    document.getElementById('q-total-num').innerText = writtenExamState.questions.length;

    const progressPct = Math.round(((idx + 1) / writtenExamState.questions.length) * 100);
    const progressFill = document.getElementById('exam-progress-fill');
    if (progressFill) progressFill.style.width = `${progressPct}%`;

    // Section Tag
    document.getElementById('q-section-badge').innerText = q.sectionName.toUpperCase();

    // Status Badge
    const statusBadge = document.getElementById('q-status-badge');
    if (statusBadge) {
        statusBadge.className = 'q-status-tag';
        if (writtenExamState.flagged.has(idx)) {
            statusBadge.classList.add('flagged');
            statusBadge.innerText = '🚩 Flagged';
        } else if (writtenExamState.userAnswers[idx] !== null) {
            statusBadge.classList.add('answered');
            statusBadge.innerText = '✓ Answered';
        } else {
            statusBadge.innerText = 'Unattempted';
        }
    }

    // Question Text
    document.getElementById('q-text-heading').innerText = `${idx + 1}. ${q.question}`;

    // Render 4 Option Cards
    const optionsContainer = document.getElementById('q-options-list');
    optionsContainer.innerHTML = '';

    const letters = ['A', 'B', 'C', 'D'];
    const userAns = writtenExamState.userAnswers[idx];
    const isAnswered = (userAns !== null);

    q.options.forEach((optText, optIdx) => {
        const optBtn = document.createElement('div');
        optBtn.className = 'option-btn';
        
        if (isAnswered) {
            optBtn.classList.add('disabled');
            if (optIdx === q.answer) {
                optBtn.classList.add('correct-choice');
            } else if (optIdx === userAns) {
                optBtn.classList.add('wrong-choice');
            }
        } else {
            if (userAns === optIdx) {
                optBtn.classList.add('selected');
            }
        }

        optBtn.innerHTML = `
            <div class="option-letter">${letters[optIdx]}</div>
            <div>${optText}</div>
        `;

        if (!isAnswered) {
            optBtn.addEventListener('click', () => {
                writtenExamState.userAnswers[idx] = optIdx;
                renderCurrentQuestion();
                updatePaletteState();
            });
        }

        optionsContainer.appendChild(optBtn);
    });

    // Render Explanation Box if Answered
    const explanationBox = document.getElementById('q-explanation-box');
    if (explanationBox) {
        if (isAnswered) {
            explanationBox.innerHTML = `
                <strong style="color: var(--accent-gold); display: block; margin-bottom: 4px;">💡 Study Explanation & Reason:</strong>
                ${q.explanation}
            `;
            explanationBox.classList.remove('d-none');
        } else {
            explanationBox.classList.add('d-none');
            explanationBox.innerHTML = '';
        }
    }

    // Control Buttons
    const prevBtn = document.getElementById('btn-prev-q');
    const nextBtn = document.getElementById('btn-next-q');

    if (prevBtn) prevBtn.disabled = (idx === 0);
    if (nextBtn) {
        if (idx === writtenExamState.questions.length - 1) {
            nextBtn.innerText = 'Submit Exam ✓';
        } else {
            nextBtn.innerText = 'Next Question →';
        }
    }

    updatePaletteState();
}

// Navigate Question
function navigateQuestion(direction) {
    const newIdx = writtenExamState.currentIndex + direction;
    if (newIdx >= 0 && newIdx < writtenExamState.questions.length) {
        writtenExamState.currentIndex = newIdx;
        renderCurrentQuestion();
    }
}

// Flag / Unflag Question
function toggleFlagCurrentQuestion() {
    const idx = writtenExamState.currentIndex;
    if (writtenExamState.flagged.has(idx)) {
        writtenExamState.flagged.delete(idx);
        showToast(`Question ${idx + 1} unflagged.`);
    } else {
        writtenExamState.flagged.add(idx);
        showToast(`Question ${idx + 1} flagged for review.`);
    }
    renderCurrentQuestion();
}

// Clear Selection for Current Question
function clearCurrentAnswer() {
    const idx = writtenExamState.currentIndex;
    writtenExamState.userAnswers[idx] = null;
    renderCurrentQuestion();
    showToast(`Cleared answer for Question ${idx + 1}.`);
}

// Start Timer Interval
function startExamTimer() {
    if (writtenExamState.timerInterval) clearInterval(writtenExamState.timerInterval);

    updateTimerDisplay();

    writtenExamState.timerInterval = setInterval(() => {
        writtenExamState.timerSeconds--;

        updateTimerDisplay();

        if (writtenExamState.timerSeconds <= 0) {
            clearInterval(writtenExamState.timerInterval);
            showToast("Time expired! Submitting your exam automatically...");
            submitWrittenExam();
        }
    }, 1000);
}

// Update Timer UI
function updateTimerDisplay() {
    const mins = Math.floor(writtenExamState.timerSeconds / 60);
    const secs = writtenExamState.timerSeconds % 60;
    const formatted = `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;

    const timerDisplay = document.getElementById('exam-timer-display');
    const timerBox = document.getElementById('exam-timer-box');

    if (timerDisplay) timerDisplay.innerText = formatted;

    if (timerBox) {
        if (writtenExamState.timerSeconds <= 600) { // 10 mins warning
            timerBox.classList.add('warning');
        } else {
            timerBox.classList.remove('warning');
        }
    }
}

// Confirm Submit
async function confirmSubmitExam() {
    const unanswered = writtenExamState.userAnswers.filter(a => a === null).length;
    let msg = "Are you sure you want to submit your Written Exam?";
    if (unanswered > 0) {
        msg = `You have ${unanswered} unanswered question(s). Are you sure you want to submit now?`;
    }

    const confirmed = await showFancyConfirm("Submit Written Exam", msg, "submit", "Yes, Submit", "No");
    if (confirmed) {
        submitWrittenExam();
    }
}

// Calculate & Submit Written Exam
function submitWrittenExam() {
    if (writtenExamState.timerInterval) clearInterval(writtenExamState.timerInterval);
    writtenExamState.active = false;

    const totalQ = writtenExamState.questions.length;
    let correctCount = 0;
    let wrongCount = 0;
    let unattemptedCount = 0;

    writtenExamState.questions.forEach((q, idx) => {
        const userAns = writtenExamState.userAnswers[idx];
        if (userAns === null) {
            unattemptedCount++;
        } else if (userAns === q.answer) {
            correctCount++;
        } else {
            wrongCount++;
        }
    });

    const percentage = Math.round((correctCount / totalQ) * 100);
    const timeSpentSec = Math.floor((new Date() - writtenExamState.startTime) / 1000);
    const timeMins = Math.floor(timeSpentSec / 60);
    const timeSecs = timeSpentSec % 60;

    // Render Results Header
    document.getElementById('result-score-num').innerText = correctCount;
    document.getElementById('result-percentage-text').innerText = `Percentage: ${percentage}%`;

    document.getElementById('r-stat-correct').innerText = correctCount;
    document.getElementById('r-stat-wrong').innerText = wrongCount;
    document.getElementById('r-stat-unattempted').innerText = unattemptedCount;
    document.getElementById('r-stat-time').innerText = `${timeMins}m ${timeSecs}s`;

    const statusBadge = document.getElementById('result-status-badge');
    if (statusBadge) {
        if (percentage >= 50) {
            statusBadge.className = 'result-badge pass';
            statusBadge.innerText = 'PASSED - QUALIFIED FOR SPEED TEST';
        } else {
            statusBadge.className = 'result-badge fail';
            statusBadge.innerText = 'NEEDS IMPROVEMENT (CUTOFF: 50%)';
        }
    }

    // Render Section Breakdown Table
    renderSectionBreakdownTable();

    // Show Result Panel
    showResultScreen();

    let modeLabel = "100 Marks Written Mock Exam";
    const titleEl = document.getElementById('exam-mode-title');
    if (titleEl) modeLabel = titleEl.innerText;

    saveExamAttemptToHistory(correctCount, totalQ, percentage, modeLabel, timeSpentSec);

    showToast(`Exam submitted! You scored ${correctCount}/${totalQ} (${percentage}%).`);
}

// Save Exam Attempt to LocalStorage History
function saveExamAttemptToHistory(score, total, percentage, modeName, timeSpentSec) {
    try {
        const historyJson = localStorage.getItem('bps16_written_history');
        const history = historyJson ? JSON.parse(historyJson) : [];
        
        const date = new Date().toISOString().split('T')[0];
        history.unshift({
            date,
            modeName,
            score,
            total,
            percentage,
            timeSpentSec,
            passed: percentage >= 50
        });
        
        localStorage.setItem('bps16_written_history', JSON.stringify(history.slice(0, 50)));
        
        // If daily drill, update total completed MCQs counter
        if (modeName.includes('Daily Study Drill')) {
            const currentFinished = parseInt(localStorage.getItem('bps16_daily_finished_mcqs') || '0') + score;
            localStorage.setItem('bps16_daily_finished_mcqs', currentFinished.toString());
        }
        
        updateDailyStudyPlanUI();
        loadExamHistoryTable();
    } catch (e) {
        console.warn("Failed to save exam history to LocalStorage:", e);
    }
}

// Load Exam History Table from LocalStorage
function loadExamHistoryTable() {
    const tbody = document.getElementById('written-history-table-body');
    if (!tbody) return;
    
    try {
        const historyJson = localStorage.getItem('bps16_written_history');
        const history = historyJson ? JSON.parse(historyJson) : [];
        
        if (history.length === 0) {
            tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted); padding: 20px;">No exam attempts logged yet. Complete a daily drill or model paper to record your history!</td></tr>`;
            return;
        }
        
        tbody.innerHTML = '';
        history.forEach((item) => {
            const statusBadge = item.passed ? 
                `<span class="category-badge cat-emerald">PASSED</span>` : 
                `<span class="category-badge cat-rose">FAILED</span>`;
                
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="time-badge">${item.date}</span></td>
                <td><strong>${item.modeName}</strong></td>
                <td><strong style="color: var(--accent-gold);">${item.score} / ${item.total}</strong></td>
                <td>${item.percentage}%</td>
                <td>${statusBadge}</td>
                <td><button onclick="openAnswerReview('all')" class="btn-secondary" style="padding: 4px 10px; font-size: 0.75rem; width: auto;">Review</button></td>
            `;
            tbody.appendChild(tr);
        });
    } catch (e) {
        console.warn("Failed to load exam history:", e);
    }
}

// Update Daily Study Plan UI Widget
function updateDailyStudyPlanUI() {
    const finishedMcqs = parseInt(localStorage.getItem('bps16_daily_finished_mcqs') || '0');
    const dayNum = Math.min(38, Math.floor(finishedMcqs / 8) + 1);
    
    const counterText = document.getElementById('daily-plan-counter-text');
    if (counterText) {
        counterText.innerText = `${finishedMcqs} / 300 Questions Finished (Day ${dayNum} of 38)`;
    }
    
    const progressFill = document.getElementById('daily-plan-progress-fill');
    if (progressFill) {
        const pct = Math.min(100, Math.round((finishedMcqs / 300) * 100));
        progressFill.style.width = `${pct}%`;
    }
}

// Section Performance Table Generator
function renderSectionBreakdownTable() {
    const tbody = document.getElementById('result-section-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const sectionsMap = {};

    writtenExamState.questions.forEach((q, idx) => {
        if (!sectionsMap[q.sectionName]) {
            sectionsMap[q.sectionName] = { total: 0, correct: 0, wrong: 0, unattempted: 0 };
        }
        const s = sectionsMap[q.sectionName];
        s.total++;

        const uAns = writtenExamState.userAnswers[idx];
        if (uAns === null) s.unattempted++;
        else if (uAns === q.answer) s.correct++;
        else s.wrong++;
    });

    Object.keys(sectionsMap).forEach(secName => {
        const s = sectionsMap[secName];
        const secPct = Math.round((s.correct / s.total) * 100);

        let perfBadge = `<span class="category-badge cat-emerald">Excellent</span>`;
        if (secPct < 50) perfBadge = `<span class="category-badge cat-rose">Weak</span>`;
        else if (secPct < 75) perfBadge = `<span class="category-badge cat-gold">Moderate</span>`;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${secName}</strong></td>
            <td>${s.total}</td>
            <td style="color: var(--accent-emerald); font-weight:700;">${s.correct}</td>
            <td style="color: var(--accent-rose);">${s.wrong}</td>
            <td><strong>${secPct}%</strong></td>
            <td>${perfBadge}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Display Result Screen
function showResultScreen() {
    document.getElementById('written-exam-active').classList.add('d-none');
    document.getElementById('written-mode-select').classList.add('d-none');
    document.getElementById('written-answer-review').classList.add('d-none');
    document.getElementById('written-exam-result').classList.remove('d-none');
}

// Reset UI back to Mode Selector
function resetWrittenExamUI() {
    document.getElementById('written-exam-active').classList.add('d-none');
    document.getElementById('written-exam-result').classList.add('d-none');
    document.getElementById('written-answer-review').classList.add('d-none');
    document.getElementById('written-mode-select').classList.remove('d-none');
}

// Open Answer Review View
function openAnswerReview(filter = 'all') {
    document.getElementById('written-exam-result').classList.add('d-none');
    document.getElementById('written-answer-review').classList.remove('d-none');

    renderAnswerReviewList(filter);
}

// Render Answer Review Questions List
function renderAnswerReviewList(filter) {
    const listContainer = document.getElementById('review-questions-list');
    if (!listContainer) return;
    listContainer.innerHTML = '';

    const letters = ['A', 'B', 'C', 'D'];

    writtenExamState.questions.forEach((q, idx) => {
        const uAns = writtenExamState.userAnswers[idx];
        const isCorrect = (uAns === q.answer);
        const isUnattempted = (uAns === null);
        const isFlagged = writtenExamState.flagged.has(idx);

        // Filter Check
        if (filter === 'correct' && !isCorrect) return;
        if (filter === 'incorrect' && (isCorrect || isUnattempted)) return;
        if (filter === 'flagged' && !isFlagged) return;

        let cardClass = 'review-item-card';
        let statusTag = '';

        if (isUnattempted) {
            cardClass += ' unattempted';
            statusTag = `<span class="q-status-tag">Not Attempted</span>`;
        } else if (isCorrect) {
            cardClass += ' correct';
            statusTag = `<span class="q-status-tag answered">✓ Correct (+1)</span>`;
        } else {
            cardClass += ' incorrect';
            statusTag = `<span class="q-status-tag" style="background: rgba(244,63,94,0.2); color: var(--accent-rose);">✗ Incorrect</span>`;
        }

        if (isFlagged) {
            statusTag += ` <span class="q-status-tag flagged">🚩 Flagged</span>`;
        }

        const card = document.createElement('div');
        card.className = cardClass;

        let optionsHtml = '';
        q.options.forEach((optText, optIdx) => {
            let optClass = 'review-opt-box';
            if (optIdx === q.answer) {
                optClass += ' correct-choice';
            } else if (optIdx === uAns && !isCorrect) {
                optClass += ' user-wrong-choice';
            }

            optionsHtml += `
                <div class="${optClass}">
                    <strong>${letters[optIdx]}.</strong> ${optText}
                    ${optIdx === q.answer ? ' ✓ (Correct Answer)' : ''}
                    ${(optIdx === uAns && !isCorrect) ? ' ✗ (Your Answer)' : ''}
                </div>
            `;
        });

        card.innerHTML = `
            <div class="review-q-header">
                <span class="category-badge cat-msoffice">${q.sectionName.toUpperCase()}</span>
                <div>${statusTag}</div>
            </div>
            <div class="review-q-text">
                Question ${idx + 1}: ${q.question}
            </div>
            <div class="review-options-grid">
                ${optionsHtml}
            </div>
            <div class="review-explanation-box">
                <strong style="color: var(--accent-gold); display: block; margin-bottom: 4px;">Detailed Explanation & Note:</strong>
                ${q.explanation}
            </div>
        `;

        listContainer.appendChild(card);
    });

    if (listContainer.children.length === 0) {
        listContainer.innerHTML = `<div style="text-align: center; color: var(--text-muted); padding: 30px;">No questions match the selected filter.</div>`;
    }
}

// Fancy Custom Confirmation Modal
function showFancyConfirm(title, message, iconType = 'confirm', actionLabel = '', cancelLabel = 'Cancel') {
    return new Promise((resolve) => {
        const modal = document.getElementById('fancy-confirm-modal');
        const iconEl = document.getElementById('fancy-confirm-icon');
        const titleEl = document.getElementById('fancy-confirm-title');
        const messageEl = document.getElementById('fancy-confirm-message');
        const cancelBtn = document.getElementById('btn-fancy-confirm-cancel');
        const actionBtn = document.getElementById('btn-fancy-confirm-action');

        if (!modal) {
            resolve(confirm(message));
            return;
        }

        // Set content
        titleEl.innerText = title;
        messageEl.innerText = message;
        cancelBtn.innerText = cancelLabel;

        // Set styles depending on actions
        if (iconType === 'quit') {
            iconEl.innerText = '✕';
            iconEl.style.color = 'var(--accent-rose)';
            actionBtn.style.background = 'linear-gradient(135deg, #ef4444, #dc2626)';
            actionBtn.innerText = actionLabel || 'Exit Test';
        } else if (iconType === 'submit') {
            iconEl.innerText = '✓';
            iconEl.style.color = 'var(--accent-emerald)';
            actionBtn.style.background = 'linear-gradient(135deg, #10b981, #059669)';
            actionBtn.innerText = actionLabel || 'Submit Test';
        } else {
            iconEl.innerText = '❓';
            iconEl.style.color = 'var(--accent-gold)';
            actionBtn.style.background = 'linear-gradient(135deg, #f59e0b, #d97706)';
            actionBtn.innerText = actionLabel || 'Yes, Proceed';
        }

        // Show modal
        modal.classList.remove('d-none');

        const handleCancel = () => {
            modal.classList.add('d-none');
            cleanup();
            resolve(false);
        };

        const handleAction = () => {
            modal.classList.add('d-none');
            cleanup();
            resolve(true);
        };

        const cleanup = () => {
            cancelBtn.removeEventListener('click', handleCancel);
            actionBtn.removeEventListener('click', handleAction);
        };

        cancelBtn.addEventListener('click', handleCancel);
        actionBtn.addEventListener('click', handleAction);
    });
}

// Popstate Browser Back/Forward navigation listener
window.addEventListener('popstate', (event) => {
    const stateObj = event.state;
    if (stateObj) {
        handleNavigationState(stateObj);
    } else {
        // Fallback to home view
        handleNavigationState({ type: 'tab', tabId: 'dashboard' });
    }
});

// Handle popstate SPA route switches
function handleNavigationState(stateObj) {
    historyStateUpdateInProgress = true;
    
    // Close confirmation modals
    const modal = document.getElementById('fancy-confirm-modal');
    if (modal) modal.classList.add('d-none');

    // Cleanly reset exam status if moving away from an active exam state
    if (writtenExamState.active && stateObj.type !== 'exam') {
        if (writtenExamState.timerInterval) clearInterval(writtenExamState.timerInterval);
        writtenExamState.active = false;
        resetWrittenExamUI();
    }

    if (stateObj.type === 'tab') {
        // Hide exam sections if tab switched
        resetWrittenExamUI();
        
        // Show correct main tab content
        const tabs = document.querySelectorAll('.tab-btn');
        const contents = document.querySelectorAll('.tab-content');
        tabs.forEach(t => t.classList.remove('active'));
        contents.forEach(c => c.classList.remove('active'));

        const tabBtn = document.querySelector(`.tab-btn[data-tab="${stateObj.tabId}"]`);
        const contentEl = document.getElementById(stateObj.tabId);
        
        if (tabBtn && contentEl) {
            tabBtn.classList.add('active');
            contentEl.classList.add('active');
            
            if (stateObj.tabId === 'analytics') {
                renderCharts();
            } else if (stateObj.tabId === 'typing') {
                setTimeout(() => {
                    document.getElementById('typing-passage-container')?.focus();
                }, 50);
            } else if (stateObj.tabId === 'written') {
                document.getElementById('prep-day-selector').classList.add('d-none');
                document.getElementById('prep-subject-grid').classList.remove('d-none');
            }
        }
    } else if (stateObj.type === 'day-selector') {
        // Navigate back to the 40-days day selection view inside Written tab
        const tabBtn = document.querySelector('.tab-btn[data-tab="written"]');
        const contentEl = document.getElementById('written');
        
        const tabs = document.querySelectorAll('.tab-btn');
        const contents = document.querySelectorAll('.tab-content');
        tabs.forEach(t => t.classList.remove('active'));
        contents.forEach(c => c.classList.remove('active'));

        if (tabBtn && contentEl) {
            tabBtn.classList.add('active');
            contentEl.classList.add('active');
        }

        // Show day selector, hide subject grid & active exam layouts
        document.getElementById('prep-subject-grid').classList.add('d-none');
        document.getElementById('prep-day-selector').classList.remove('d-none');
        document.getElementById('selected-subject-title').innerText = stateObj.subjectName;

        // Rebuild days buttons grid
        const daysGrid = document.getElementById('subject-days-grid');
        if (daysGrid) {
            daysGrid.innerHTML = '';
            for (let d = 1; d <= 40; d++) {
                const dayBtn = document.createElement('button');
                dayBtn.className = 'palette-btn';
                dayBtn.style.padding = '10px';
                dayBtn.style.fontSize = '0.9rem';
                dayBtn.innerText = `Day ${d}`;
                dayBtn.addEventListener('click', () => {
                    startWrittenExam('daily-subject-day', { subject: stateObj.subject, dayNum: d });
                });
                daysGrid.appendChild(dayBtn);
            }
        }
    } else if (stateObj.type === 'written-home') {
        // Go back to the main list of papers and subjects
        const tabBtn = document.querySelector('.tab-btn[data-tab="written"]');
        const contentEl = document.getElementById('written');
        
        const tabs = document.querySelectorAll('.tab-btn');
        const contents = document.querySelectorAll('.tab-content');
        tabs.forEach(t => t.classList.remove('active'));
        contents.forEach(c => c.classList.remove('active'));

        if (tabBtn && contentEl) {
            tabBtn.classList.add('active');
            contentEl.classList.add('active');
        }

        document.getElementById('prep-day-selector').classList.add('d-none');
        document.getElementById('prep-subject-grid').classList.remove('d-none');
    }

    historyStateUpdateInProgress = false;
}
