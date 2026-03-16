Set shell = CreateObject("WScript.Shell")

' 1. Показываем ошибку (сохраняй в ANSI, чтобы был русский текст!)
MsgBox "Критическая ошибка системы!" & vbCrLf & "Обнаружен взлом Шлёпы.", 16, "Windows Security"

' 2. Путь к картинке
picPath = "C:\shlepa.jpg"

' 3. Просто открываем картинку на весь экран
shell.Run picPath, 3, False