/*

# Priority TODO
• Remake sidebar UI to emphasize source links
• various uploads tools (renaming, stabilization) in Adminhelper
• detect that frozen rows were probably unfrozen, and suggest to restore them
• Factor out WorksheetSectionBuilder and provide worksheet options to it (such as date)
  • or make a separate group options for sections
  • or check whether the worksheet has a date, and if it doesn't, default to empty date.
• Detect empty files in uploads and question them (or replace hash part with EMPTY and detect in in conditional formatting).
• Detect invalid URLs in uploads and reject them?

# Priority TODO (2)
• make markers invisible in published mode (add a cf rule to whiten them).
• Add action: sort rows in name order
• Convert worksheet to theory
• Worksheet plan
  • optimize worksheet generation to only load and save cfrules once.
• Upload dialog
  • upon hitting upload button, add any unconfirmed text in the text adder to file list.

# TODO
• make a notice to link insertion (in file upload) that only file links and Drive links are ok, not Overleaf links.
• mark “burning” problems with fire emoji — requires a lot of workarounds, though.
• presets of worksheets: game, theory, … (a list is saved in spreadsheet metadata);
• upload process must explicitly fail if the file is not found or is a directory;
• Upload solutions:
  • add it to the sidebar
• Sidebar refactoring:
  [done] make sidebar validation not acquire the lock unless necessary
  • factor out contents item as a class of its own.
  • Sidebar upload optimization: open upload dialog immediately, filling necessary fields and enabling ui as data is validated
    • concept: OpportunisticPromise; it has current value, which may get updated eventually.
    • refactor category_css to use data-category attribute
• StudyGroup creator
  • proper interface for editing timetable
• StudyGroup metadata editor
• Spreadsheet metadata editor
• Upload configuration editor
• Resolve Actions/Worksheets XXX
• Resolve any other XXX
• indent files with 4 spaces
• make StudyGroup resistant to the deletion of the last column
  • maybe hide it
  • maybe just output a message that would suggest copying a separator column
    and moving it in correct place
• Admin mode and introduction
• Multiadd worksheets: add several worksheets or add a worksheet to several groups at once
• All formulas in WorksheetLig/Worksheet and WorksheetLib/StudyGroup should use SpreadsheetLib/Formula to guarantee locale compatibility.
• Function to reorder section in worksheet?
• Rename S3Lib/Signer to Upload/S3Signer
• Regenerate:
  • list (combined lists of students for printing)
  • s-summary (special)
  • s-sample (special)
• Refactor stylesheets to use CSS custom properties?
• Special courses:
  • Script to create special course groups
  • Recreate s-summary
• a function that detects whether attendance was filled the last week
• make a separate (hidden) row for problem status (and maybe make some statistics based on it — like how many problems remain for a given student)
• make CFormatting reasonably resistant to #REF! errors in formulas

*/
