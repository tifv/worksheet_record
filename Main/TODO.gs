/*

# FIXME future
• make worksheet sheet resistant to deletion of the last column
  • maybe hide it

# Metadata
(either JSON or string)

SPREADHSHEET
worksheet__meta--color_schemes = {
  "gray" : {marks: [h, l, s], limit: [h, l, s]},
  … }
worksheet__meta--categories = {
  "a" : {color: [h, l, s]},
  … }

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