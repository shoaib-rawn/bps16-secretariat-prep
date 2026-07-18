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
        day: 45,
        category: "theory",
        title: "Exam Day Protocol & Final Checklist",
        description: "Final exam strategy: 1. Keep hands warm for shorthand and typing. 2. In shorthand, focus on outline accuracy; don't stress over a single missed word. 3. In typing, accuracy yields speed; typing at 99% accuracy is faster than correcting mistakes. 4. In MS Office, save your file (Ctrl+S) every 30 seconds to prevent lost work during server or PC issues.",
        shortcut: "Ctrl + S (Save Document - use continuously during practical exams!)",
        quickTip: "Bring your own shorthand notepad and pencil (2B lead is softer and faster for drawing outlines) to the test center.",
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

// 2. Typing Practice Passages for Custom Drill
const TYPING_PASSAGES = [
    "The President's Secretariat, Aiwan-e-Sadr, functions as the administrative core for the head of state, coordinating official state affairs, diplomatic receptions, and statutory declarations. As an Assistant Private Secretary, precision in official communication, speed in shorthand dictation, and mastery over file management systems are core prerequisites for managing the workflow of high-ranking government officials.",
    "Microsoft Word section breaks are fundamental tools for crafting professional reports. They allow the author to divide a document into segments that can have distinct page numbering formats, header templates, and orientation margins. Without section breaks, uncoupling a title page or inserting landscape sheets within a portrait dossier is impossible, which is why examiners focus heavily on this topic.",
    "Excel spreadsheet models are widely deployed in ministries to track budgets, personnel registers, and monthly allocations. Private secretaries are expected to know lookup functions like INDEX and MATCH, logical calculations using nested IFS, and pivot summaries. Ensuring proper print layouts with repeating title rows is critical to creating clean reports for executive reviews."
];

// 3. Motivational Quotes & Guidance
const MOTIVATION_INSIGHTS = [
    {
        quote: "Your current shorthand speed of 120 WPM is exceptionally good. Only a tiny fraction of applicants can achieve this speed. Protect this asset: do 30 minutes of dictation daily, and focus heavily on translation accuracy onto the keyboard. The BPS-16 post is within your reach.",
        author: "Personal Coach"
    },
    {
        quote: "BPS-16 at the President's Secretariat (Aiwan-e-Sadr) is not just a job; it is a prestigious career path with high security, excellent pension benefits, and proximity to the governance center. Maintaining 50 WPM and building MS Office mastery will secure your success.",
        author: "Secretariat Protocol Coach"
    },
    {
        quote: "Success is the sum of small efforts, repeated day in and day out. 45 days of focused discipline will separate you from the competition. Today's MS Office Tip is your stepping stone.",
        author: "Robert Collier"
    }
];

// 4. Application State (Loaded from Local Storage or initialized)
let state = {
    currentDay: 1,
    completedDays: [],
    shorthandLogs: [
        { date: '2026-07-10', speed: 120, accuracy: 95 },
        { date: '2026-07-12', speed: 120, accuracy: 96 },
        { date: '2026-07-15', speed: 122, accuracy: 95 },
        { date: '2026-07-18', speed: 120, accuracy: 97 }
    ],
    typingLogs: [
        { date: '2026-07-10', wpm: 50, accuracy: 96 },
        { date: '2026-07-12', wpm: 51, accuracy: 97 },
        { date: '2026-07-15', wpm: 52, accuracy: 98 },
        { date: '2026-07-18', wpm: 50, accuracy: 99 }
    ]
};

// Load state from local storage
function loadState() {
    try {
        const saved = localStorage.getItem('bps16_prep_state');
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
        localStorage.setItem('bps16_prep_state', JSON.stringify(state));
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
    updateCountdown();
    showToast("Dashboard Loaded successfully. Welcome, candidate!");
});

// Tab Switch Logic
function setupTabs() {
    const tabs = document.querySelectorAll('.tab-btn');
    const contents = document.querySelectorAll('.tab-content');
    
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            contents.forEach(c => c.classList.remove('active'));
            
            tab.classList.add('active');
            const targetId = tab.dataset.tab;
            document.getElementById(targetId).classList.add('active');
            
            if (targetId === 'analytics') {
                renderCharts();
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
    
    // Quick tip card header on dashboard
    const todayTip = SYLLABUS[state.currentDay - 1] || SYLLABUS[0];
    document.getElementById('dash-today-title').innerText = `Day ${todayTip.day}: ${todayTip.title}`;
    document.getElementById('dash-today-desc').innerText = todayTip.description.replace(/\*\*/g, '').substring(0, 160) + "...";
    document.getElementById('dash-today-category').innerText = todayTip.category.toUpperCase();
    document.getElementById('dash-today-category').className = `category-badge cat-${todayTip.category}`;
    
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
    
    const w = container.clientWidth || 400;
    const h = container.clientHeight || 220;
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
        <svg class="chart-svg">
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
    const inputArea = document.getElementById('typing-drill-input');
    const resetBtn = document.getElementById('btn-reset-drill');
    const wpmResult = document.getElementById('drill-wpm');
    const accResult = document.getElementById('drill-acc');
    
    if (!drillText || !inputArea) return;
    
    function resetDrill() {
        characterIndex = 0;
        typingErrors = 0;
        typingStartTime = null;
        drillActive = false;
        inputArea.value = '';
        wpmResult.innerText = '0';
        accResult.innerText = '100%';
        
        // Render characters
        const text = TYPING_PASSAGES[currentPassageIndex];
        drillText.innerHTML = text.split('').map((char, index) => {
            return `<span class="${index === 0 ? 'current' : ''}">${char}</span>`;
        }).join('');
    }
    
    // Choose other passages
    document.getElementById('btn-next-passage')?.addEventListener('click', () => {
        currentPassageIndex = (currentPassageIndex + 1) % TYPING_PASSAGES.length;
        resetDrill();
        showToast("Loaded new drill passage!");
    });
    
    inputArea.addEventListener('input', (e) => {
        const textSpans = drillText.querySelectorAll('span');
        const typedVal = inputArea.value;
        
        if (!typingStartTime && typedVal.length > 0) {
            typingStartTime = new Date();
            drillActive = true;
        }
        
        characterIndex = typedVal.length;
        
        // Calculate WPM and Accuracy
        const text = TYPING_PASSAGES[currentPassageIndex];
        typingErrors = 0;
        
        for (let i = 0; i < textSpans.length; i++) {
            if (i < typedVal.length) {
                if (typedVal[i] === text[i]) {
                    textSpans[i].className = 'correct';
                } else {
                    textSpans[i].className = 'incorrect';
                    typingErrors++;
                }
            } else if (i === typedVal.length) {
                textSpans[i].className = 'current';
            } else {
                textSpans[i].className = '';
            }
        }
        
        // Update stats
        if (typedVal.length > 0) {
            const timeDiffMin = (new Date() - typingStartTime) / 60000;
            const grossWpm = Math.round((typedVal.length / 5) / (timeDiffMin || 0.01));
            const netWpm = Math.max(0, grossWpm - Math.round(typingErrors / (timeDiffMin || 0.01)));
            const accuracy = Math.round(((typedVal.length - typingErrors) / typedVal.length) * 100);
            
            wpmResult.innerText = netWpm;
            accResult.innerText = `${accuracy}%`;
            
            // Finish Condition
            if (typedVal.length >= text.length) {
                drillActive = false;
                showToast(`Drill complete! net speed: ${netWpm} WPM | Acc: ${accuracy}%`);
                
                // Add to typing log automatically
                const date = new Date().toISOString().split('T')[0];
                state.typingLogs.push({ date, wpm: netWpm, accuracy });
                saveState();
                updateLogTables();
            }
        }
    });
    
    resetBtn?.addEventListener('click', resetDrill);
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

// 9. Countdown to Skills Test
function updateCountdown() {
    // Standard 45 Days duration from today's target (assuming Sep 1st, 2026)
    const testDate = new Date(2026, 8, 1, 9, 0, 0);
    const today = new Date();
    
    const diffTime = Math.abs(testDate - today);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    const countdownEl = document.getElementById('countdown-days');
    if (countdownEl) {
        countdownEl.innerText = diffDays > 0 ? diffDays : "0";
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
