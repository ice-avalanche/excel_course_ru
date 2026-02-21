# Урок 15. Интеграция Excel с приложениями Office

## 1. Назначение

Excel может работать не изолированно, а как часть экосистемы Microsoft Office.  
Интеграция позволяет:

- автоматически формировать отчёты в Word;
- отправлять письма из Outlook;
- создавать презентации в PowerPoint;
- импортировать данные из других приложений;
- строить корпоративные процессы на базе Excel.

## 2. Интеграция с Word

### 2.1 Связанные объекты

Можно вставить таблицу Excel в Word:

`Word → Вставка → Объект → Из файла → Связать с файлом`

Преимущество:
- при обновлении Excel документ Word обновляется автоматически.

### 2.2 Слияние документов (Mail Merge)

Excel используется как источник данных:

1. Подготовить таблицу (ФИО, Должность, Оклад).
2. В Word → Рассылки → Выбрать получателей → Использовать существующий список.
3. Подключить Excel-файл.
4. Вставить поля слияния.
5. Создать персонализированные документы.

Применение:
- договоры;
- письма;
- уведомления;
- справки.

### 2.3 Генерация Word-отчета через VBA

Пример кода:

```vba
Sub CreateWordReport()

    Dim wdApp As Object
    Dim wdDoc As Object

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    Set wdDoc = wdApp.Documents.Add

    wdDoc.Content.InsertAfter "Отчет по продажам" & vbCrLf
    wdDoc.Content.InsertAfter "Общая выручка: " & Range("B2").Value

End Sub
````

## 3. Интеграция с Outlook

### 3.1 Отправка писем из Excel

Пример:

```vba
Sub SendMail()

    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = "example@mail.com"
        .Subject = "Отчет"
        .Body = "См. данные в Excel."
        .Display
    End With

End Sub
```

Можно:

* прикреплять файлы;
* автоматически подставлять текст;
* отправлять персонализированные письма.

### 3.2 Рассылка по списку

Цикл по таблице:

```vba
For i = 2 To 100
    OutMail.To = Cells(i, 1).Value
    OutMail.Body = "Здравствуйте, " & Cells(i, 2).Value
Next i
```

## 4. Интеграция с PowerPoint

### 4.1 Вставка диаграмм

Excel-графики можно:

* вставлять как связанные объекты;
* обновлять автоматически;
* использовать как основу презентаций.

### 4.2 Генерация презентации через VBA

```vba
Sub CreatePresentation()

    Dim ppApp As Object
    Dim ppPres As Object
    Dim slide As Object

    Set ppApp = CreateObject("PowerPoint.Application")
    ppApp.Visible = True

    Set ppPres = ppApp.Presentations.Add
    Set slide = ppPres.Slides.Add(1, 1)

    slide.Shapes(1).TextFrame.TextRange.Text = "Финансовый отчет"
    slide.Shapes(2).TextFrame.TextRange.Text = "Выручка: " & Range("B2").Value

End Sub
```

## 5. Интеграция с PDF

Excel позволяет:

* сохранять отчеты в PDF;
* формировать автоматическую выгрузку.

```vba
ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:="C:\Report.pdf"
```

## 6. Импорт данных из других приложений

### Источники:

| Источник   | Способ                        |
| ---------- | ----------------------------- |
| Word       | Копирование / Power Query     |
| Outlook    | Экспорт в CSV                 |
| Access     | Подключение через Power Query |
| Web        | Импорт через Power Query      |
| SharePoint | Онлайн-связь                  |

## 7. Интеграция через Power Query

Power Query позволяет:

* объединять данные из нескольких файлов;
* подключаться к SharePoint;
* получать данные из веб-API;
* обновлять данные автоматически.

## 8. Создание корпоративного процесса

Пример сценария:

1. Excel рассчитывает показатели.
2. Макрос формирует Word-отчет.
3. Отчет сохраняется в PDF.
4. Outlook отправляет файл руководителю.
5. PowerPoint создаёт презентацию для совещания.

## 9. Безопасность и ограничения

| Риск                     | Решение                       |
| ------------------------ | ----------------------------- |
| Блокировка макросов      | Подписывать файлы             |
| Несовместимость версий   | Проверка версий Office        |
| Ограничение прав доступа | Использовать доверенные папки |
