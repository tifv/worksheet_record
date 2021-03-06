/*

# Priority TODO
• detect duplicate uploads
• detect duplicate upload names
• detect obsoleted uploads
• various uploads tools (renaming, stabilization) in Adminhelper
• detect that sheet is probably a copy of proper StudyGroup and suggest to restore metadata
• detect that frozen rows were probably unfrozen, and suggest to restore them

# Second-priority TODO
• make markers invisible in published mode (add a cf rule to whiten them).
• make allow_subproblems operate on worksheets that surround every range of range list.
• Add action: sort rows in name order
• Convert worksheet to theory
• Worksheet plan
  • optimize worksheet generation to only load and save cfrules once.

# TODO
• mark “burning” problems with fire emoji — requires a lot of workarounds, though.
• presets of worksheets: game, theory, … (a list is saved in spreadsheet metadata);
• upload process must explicitly fail if the file is not found or is a directory;
• Upload solutions:
  • add it to the sidebar
• Sidebar refactoring:
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

# Metadata
(either JSON or string)

SPREADHSHEET
…

SHEET
worksheet__group = null
worksheet__group--categories = ["a", "g", "c", "o"]
worksheet__group--color_scheme = "gray"
worksheet__group--timetable = {"2019-11-11": [1, 2, 3], …}

COLUMN
worksheet--start_col
worksheet--end_col


# DocumentProperties
upload_folder = "…folder id…"
upload_name[email]


# UserProperties
%WORKSHEET_RECORD_ID%:upload_subfolder = "…subfolder id…"
upload_author = "…"


# HLS to RGB
https://stackoverflow.com/questions/2353211/
https://stackoverflow.com/questions/36721830/


# Upload development
Upload menu split into two tabs: "Доступ" is about
• access to the adminrecord
  • for the owner: creating adminrecord
• access to the upload_folder
  • for the owner: creating upload_folder
• existence the upload_subfolder
• ability to verify that all files in upload_subfolder were copied over and not needed anymore
and "Загрузка файлов" is about
• displaying all information stored:
  • worksheet group
  • worksheet date
  • worksheet category
  • worksheet title (can be edited)
  • uploader email
  • author name (preserved across uploads in upload_author user property)
• uploading files
  • a notice is shown:
    Загруженные файлы будут доступны по ссылке.
    Не загружайте материалы, которые нельзя выкладывать в интернет.
  • large labels "PDF" и " with small comments
  • as files are prepared to upload, label area decreases, down to some limit

For the owner, there is a separate dialog to duplicate all «living» files in her own Drive, and also replace links.

Also, try including forms in the response — maybe files are working again.

# Unlinked
Move all "Unlinked" code to Adminrecord. Nobody can use it anyway.


# Regeneration
Create scripts that regenerate worksheets sheet and other sheets…


# Metadata
Make yet another library?


# Garbage

[Загрузка листочков / Adminhelper]
По окончании сборов нужно выдавать авторам список их файлов, чтобы они могли их удалить из Drive.

[Worksheet]
• Убери ограничение при удалении столбцов. Или нет

[Функции]
• Обойди все обращения к Date, напиши собственные функции для этого.
• Зачисти пространство функций, доступных снаружи.

[Посещаемость]
Напиши функцию, которая выдает ошибку, если посещаемость за последнюю неделю не была заполнена.
* Ну может и не надо.

*/
