# Freeoffice

Freeoffice - это библиотека для создания docx документов позволяет добавлять такие элементы как текст, изображение и таблицы

## Установка

Выполните команду в терминале:

```bash
go get github.com/AlexsRyzhkov/freeoffice
```

## Работа с библиотекой


##### Пример работы

```go
import (
    "github.com/AlexsRyzhkov/freeoffice"
)

func main(){
	
	// Создание документа
	d := docx.New()
	
	// Получение документа
	document := d.GetDocument()
	
	//  (text, property)
	document.AddParagraph("text", &fragments.TextProperty{})
	
	//  (url, property)
	document.AddImage("url", &fragments.ImageProperty{})
    
	//  (row, col)
	document.AddTable(2, 3)
}
```

## Доступные методы для текстового параграфа

```go
type ITextParagraph interface {
	// Измененяет текст
	SetText(string) ITextParagraph
    
	// Делает текс жирным
	SetBold() ITextParagraph
	// Делает текст курсивным
	SetItalic() ITextParagraph
	// Делает текст подчеркнутым "
	SetUnderline() ITextParagraph
	// Делает текст зачеркнутым
	SetStrike() ITextParagraph
    
	// Противоположные верхним
	UnSetBold() ITextParagraph
	UnSetItalic() ITextParagraph
	UnSetUnderline() ITextParagraph
	UnSetStrike() ITextParagraph
    
	// Изменяет font family 
	SetFontFamily(string) ITextParagraph
	// Изменяет размер текста
	SetFontSize(int) ITextParagraph
    
	// Изменяет цвет текста
	SetTextColor(string) ITextParagraph
	// Изменяет цвет выделения
	SetTextHighlightColor(string) ITextParagraph
    
	// Добавляет выравнивание
	SetJustify(string) ITextParagraph
    // Добавляет отступ слева
	SetLeftOffSet(string) ITextParagraph
    // Добавляет отступ справа
	SetRightOffSet(string) ITextParagraph
}
```