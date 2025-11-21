Для конвертации doc,odt,rtf в doc,docx и т.д. через Word используйте метод Documents.OpenNoRepairDialog для максимальной автоматизации.  
А то открываются окна, которые мешают работе робота. Так как требуют их обработки.
1) Окно подтверждения исправлений, если конвертируемый файл не исправен и Word автоматически исправляет файл.
2) Окно не корректно кодировки.
3) Окно ввода пароля для открытия в режиме записи.  

Пример  
```python

    try:
        word = win32com.client.Dispatch("Word.Application")
    except Exception as e:
        raise
    else:
        try:
            #word.Visible = False
            # 5 аргумент у OpenNoRepairDialog это пароль. Специально рандомный написал, чтобы выдавало сразу ошибку при открытии документа. А то Word требуется пароль и окно с требованием пароля висит бесконечно и блочит робота.
            # OpenAndRepair в False блокирует окно подтверждения исправлений.
            # 3 аргумент в True - это открытие файла в режиме чтения. Его хватает для конвертации в другой формат
            wordDoc = word.Documents.OpenNoRepairDialog(PathToFile , False,True,False,"34125" ,OpenAndRepair = False)
            wordDoc.SaveAs2(PathToConvertFile, FileFormat = FormatConvert.value)
            wordDoc.Close()
        except Exception as e2:
            try:
                word.Quit()
            except Exception as e3:
                pass
            raise
    word.Quit()






    
class Format(enum.Enum):
    wdFormatDocument =	0
    wdFormatDOSText =	4
    wdFormatDOSTextLineBreaks =	5
    wdFormatEncodedText =	7
    wdFormatFilteredHTML =	10
    wdFormatFlatXML =	19
    wdFormatFlatXMLMacroEnabled =	20
    wdFormatFlatXMLTemplate =	21
    wdFormatFlatXMLTemplateMacroEnabled =	22
    wdFormatOpenDocumentText =	23
    wdFormatHTML =	8
    wdFormatRTF =	6
    wdFormatStrictOpenXMLDocument =	24
    wdFormatTemplate =	1
    wdFormatText =	2
    wdFormatTextLineBreaks =	3
    wdFormatUnicodeText =	7
    wdFormatWebArchive =	9
    wdFormatXML =	11
    wdFormatDocument97 =	0
    wdFormatDocumentDefault =	16 #docx
    wdFormatPDF =	17
    wdFormatTemplate97 =	1
    wdFormatXMLDocument =	12
    wdFormatXMLDocumentMacroEnabled =	13
    wdFormatXMLTemplate =	14
    wdFormatXMLTemplateMacroEnabled =	15
    wdFormatXPS = 18


ListCorrectFormat = [".docx",".doc",".docm",".dot",".dotm",".dotx",".htm",".html",".mhtml",".odt",".pdf",".rtf",".txt",".wps",".xml",".xps"]

```
