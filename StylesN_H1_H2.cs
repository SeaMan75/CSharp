using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Style = Microsoft.Office.Interop.Word.Style;
using Word = Microsoft.Office.Interop.Word;

namespace Styles
{
    public partial class StylesN_H1_H2
    {
        // ✅ ГЛОБАЛЬНЫЕ СВОЙСТВА — видны во всех методах класса,
        // всегда возвращают актуальные значения (активный документ может меняться).
        private Word.Application App => Globals.ThisAddIn.Application;
        private Word.Document Doc => App.ActiveDocument;

        // ✅ НОВОЕ: типы inline-форматирования
        private enum InlineFormat
        {
            None = 0,
            Bold,
            Italic,
            BoldItalic,
            Underline,
            Strikethrough,
            Code,
            Link,
            Image
        }

        // ✅ НОВОЕ: структура для хранения найденной inline-правки
        // ✅ ОБНОВЛЕНО: добавлено свойство extraData для хранения URL ссылок и картинок
        private class InlineEdit
        {
            public int fullStart;      // начало всего match (включая маркеры)
            public int fullEnd;        // конец всего match (включая маркеры)
            public int innerStart;     // начало внутреннего текста
            public int innerEnd;       // конец внутреннего текста
            public InlineFormat format;
            public string extraData;   // здесь будем хранить URL для ссылок и картинок
        }

        private void StylesN_H1_H2_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                var wordVersion = App.Version;
                var interopVersion = typeof(Microsoft.Office.Interop.Word.Application).Assembly.GetName().Version;

                MessageBox.Show($"Add-in v{version}\nWord v{wordVersion}\nInterop v{interopVersion}");
            }
            catch { }
        }

        #region CLICK    
        /// <summary>
        /// Настройка стилей документа
        /// </summary>
        private void buttonStyles_Click(object sender, RibbonControlEventArgs e)
        {
            // Убираем инициализацию переменных через null - это не нужно
            var doc = Doc;

            if (doc == null)
            {
                MessageBox.Show("Нет активного документа.");
                return;
            }

            try
            {
                // Настройка стиля "Картинка"
                Style pictureStyle = null;

                try
                {
                    pictureStyle = doc.Styles["Картинка"];
                }
                catch
                {
                    pictureStyle = doc.Styles.Add("Картинка", Word.WdStyleType.wdStyleTypeParagraph);
                }

                // Применяем параметры стиля
                pictureStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                pictureStyle.ParagraphFormat.SpaceBefore = 14; // ≈ одна строка при размере шрифта 14 pt
                pictureStyle.ParagraphFormat.SpaceAfter = 0;
                pictureStyle.ParagraphFormat.FirstLineIndent = 0;
                pictureStyle.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                pictureStyle.set_BaseStyle(""); // без базового стиля
                //Настройка дополнительного стиля    
                // Получение стиля "Заголовок 1"
                var heading1Style = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading1];
                if (heading1Style != null)
                {
                    // Создание нового стиля "Заголовок 1 Дополнительный"
                    Style additionalHeading1Style = null;
                    try
                    {
                        additionalHeading1Style = doc.Styles["Заголовок 1 Дополнительный"];
                    }
                    catch
                    {
                        additionalHeading1Style = doc.Styles.Add("Заголовок 1 Дополнительный", Word.WdStyleType.wdStyleTypeParagraph);
                    }

                    additionalHeading1Style.Font.AllCaps = 1; // Все буквы заглавные
                    additionalHeading1Style.set_BaseStyle(heading1Style.NameLocal);
                }
                else
                {
                    MessageBox.Show("Стиль 'Заголовок 1' не найден.");
                }


                // Настройка стиля для подписи картинки:
                Style pictureCaptionStyle = null;

                try
                {
                    pictureCaptionStyle = doc.Styles["Подпись картинки"];
                }
                catch
                {
                    pictureCaptionStyle = doc.Styles.Add("Подпись картинки", Word.WdStyleType.wdStyleTypeParagraph);
                }

                pictureCaptionStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                pictureCaptionStyle.ParagraphFormat.SpaceBefore = 0;
                pictureCaptionStyle.ParagraphFormat.SpaceAfter = 28;
                pictureCaptionStyle.ParagraphFormat.FirstLineIndent = 0;
                pictureCaptionStyle.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                pictureCaptionStyle.set_BaseStyle("");


                // Настройка стиля для подписи таблицы:

                Style tableCaptionStyle = null;

                try
                {
                    tableCaptionStyle = doc.Styles["Подпись таблицы"];
                }
                catch
                {
                    tableCaptionStyle = doc.Styles.Add("Подпись таблицы", Word.WdStyleType.wdStyleTypeParagraph);
                }

                tableCaptionStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableCaptionStyle.ParagraphFormat.SpaceBefore = 28;
                tableCaptionStyle.ParagraphFormat.SpaceAfter = 0;
                tableCaptionStyle.ParagraphFormat.FirstLineIndent = App.CentimetersToPoints(1.25f);
                tableCaptionStyle.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                tableCaptionStyle.set_BaseStyle("");

                // Настройка стиля "Обычный"
                var normalStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
                if (normalStyle != null)
                {
                    normalStyle.Font.Name = "Times New Roman Cyr";
                    normalStyle.Font.Size = 14;
                    normalStyle.Font.Bold = 0; // false
                    normalStyle.Font.Color = Word.WdColor.wdColorAutomatic;

                    normalStyle.ParagraphFormat.FirstLineIndent = App.CentimetersToPoints(1.25f);
                    normalStyle.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                    normalStyle.ParagraphFormat.SpaceBefore = 0;
                    normalStyle.ParagraphFormat.SpaceAfter = 0;
                    normalStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else
                {
                    MessageBox.Show("Стиль 'Обычный' не найден.");
                }

                // Настройка стиля "Заголовок 1"
                if (heading1Style != null)
                {
                    heading1Style.Font.Name = "Times New Roman Cyr";
                    heading1Style.Font.Size = 14;
                    heading1Style.Font.Bold = 1; // true
                    heading1Style.Font.Color = Word.WdColor.wdColorAutomatic;

                    heading1Style.ParagraphFormat.SpaceBefore = 0;
                    heading1Style.ParagraphFormat.SpaceAfter = 28;
                    heading1Style.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                    heading1Style.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    heading1Style.ParagraphFormat.FirstLineIndent = 0;
                    heading1Style.ParagraphFormat.PageBreakBefore = -1; // true

                    heading1Style.set_BaseStyle("");
                }
                else
                {
                    MessageBox.Show("Стиль 'Заголовок 1' не найден.");
                }

                // Настройка стиля "Заголовок 2"
                var heading2Style = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading2];
                if (heading2Style != null)
                {
                    // ❌ БЫЛО: heading2Style.Font = new Word.Font { ... };
                    // ✅ СТАЛО:
                    heading2Style.Font.Name = "Times New Roman Cyr";
                    heading2Style.Font.Size = 14;
                    heading2Style.Font.Bold = 1; // true
                    heading2Style.Font.Color = Word.WdColor.wdColorAutomatic;

                    heading2Style.ParagraphFormat.SpaceBefore = 0;
                    heading2Style.ParagraphFormat.SpaceAfter = 28;
                    heading2Style.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                    heading2Style.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    heading2Style.ParagraphFormat.FirstLineIndent = 0;
                    heading2Style.ParagraphFormat.PageBreakBefore = 0; // false

                    heading2Style.set_BaseStyle("");
                }
                else
                {
                    MessageBox.Show("Стиль 'Заголовок 2' не найден.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            // ❌ УДАЛЕНО: finally с Marshal.ReleaseComObject
            // Стили и документы управляются Word - не нужно освобождать их вручную!
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            setIndent(1.25f, 0);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Doc;

            if (doc == null)
            {
                MessageBox.Show("Нет активного документа.");
                return;
            }

            try
            {
                // Установка полей документа
                doc.PageSetup.TopMargin = App.CentimetersToPoints(2);
                doc.PageSetup.BottomMargin = App.CentimetersToPoints(2);
                doc.PageSetup.LeftMargin = App.CentimetersToPoints(3);
                doc.PageSetup.RightMargin = App.CentimetersToPoints(1.5f);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            // ❌ УДАЛЕНО: finally с Marshal.ReleaseComObject
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Doc;

            if (doc == null)
            {
                MessageBox.Show("Нет активного документа.");
                return;
            }

            try
            {
                // Установка полей документа
                doc.PageSetup.TopMargin = App.CentimetersToPoints(2);
                doc.PageSetup.BottomMargin = App.CentimetersToPoints(2);
                doc.PageSetup.LeftMargin = App.CentimetersToPoints(3);
                doc.PageSetup.RightMargin = App.CentimetersToPoints(1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            // ❌ УДАЛЕНО: finally с Marshal.ReleaseComObject
        }




        // В button8_Click через button11_Click ошибок нет - они правильно работают с объектами
        // без создания новых Font/ParagraphFormat через new

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = App.Selection;

            if (selection.Tables.Count > 0)
            {
                var table = selection.Tables[1];

                // Обработка шапки таблицы
                var headerRow = table.Rows[1];
                foreach (Word.Cell cell in headerRow.Cells)
                {
                    cell.Range.Font.Bold = 1;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.Range.ParagraphFormat.FirstLineIndent = 0;
                }

                // Обработка остальной части таблицы
                for (int i = 2; i <= table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    foreach (Word.Cell cell in row.Cells)
                    {
                        cell.Range.Font.Bold = 0;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        cell.Range.ParagraphFormat.FirstLineIndent = 0;
                    }
                }

                // Общие настройки для всего выделения
                selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                selection.Font.Size = 12;
                selection.Font.Name = "Times New Roman Cyr";
                selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            }
            else
            {
                MessageBox.Show("Выделенная таблица не найдена.");
            }
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Doc;
            var selection = App.Selection;

            if (selection.Tables.Count > 0)
            {
                string tableName = Clipboard.GetText();
                if (string.IsNullOrEmpty(tableName))
                {
                    MessageBox.Show("Буфер обмена пуст.");
                    return;
                }

                var table = selection.Tables[1];
                var range = table.Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                var captionRange = range.Duplicate;
                captionRange.InsertCaption(
                    Label: "Таблица",
                    Title: $" «{tableName}»",
                    TitleAutoText: null,
                    Position: Word.WdCaptionPosition.wdCaptionPositionAbove,
                    ExcludeLabel: false
                );

                // Применяем стиль "Подпись к таблице", если он существует
                Style captionStyle = null;
                try
                {
                    captionStyle = doc.Styles["Подпись к таблице"];
                }
                catch
                {
                    // Стиль не найден, ничего не делаем
                }

                if (captionStyle != null)
                {
                    range.Paragraphs[1].Range.set_Style(captionStyle);
                }
            }
            else
            {
                MessageBox.Show("Выделенная таблица не найдена.");
            }
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Doc;
            var selection = App.Selection;

            if (selection.InlineShapes.Count > 0 || selection.ShapeRange.Count > 0)
            {
                string pictureName = Clipboard.GetText();
                if (string.IsNullOrEmpty(pictureName))
                {
                    MessageBox.Show("Буфер обмена пуст.");
                    return;
                }

                Word.Range range;
                if (selection.InlineShapes.Count > 0)
                {
                    range = selection.InlineShapes[1].Range;
                }
                else
                {
                    range = selection.ShapeRange[1].Anchor;
                }
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                var captionRange = range.Duplicate;
                captionRange.InsertCaption(
                    Label: "Рисунок",
                    Title: $" {char.ConvertFromUtf32(0x2014)} {pictureName}",
                    TitleAutoText: null,
                    Position: Word.WdCaptionPosition.wdCaptionPositionBelow,
                    ExcludeLabel: false
                );

                Style captionStyle = null;
                try
                {
                    captionStyle = doc.Styles["Подпись картинки"];
                }
                catch
                {
                    MessageBox.Show("Стиль 'Подпись картинки' не найден.");
                    return;
                }

                if (captionStyle != null)
                {
                    captionRange.Paragraphs[1].Range.set_Style(captionStyle);
                }
            }
            else
            {
                MessageBox.Show("Выделенная картинка не найдена.");
            }
        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Doc;
            var selection = App.Selection;

            if (selection.InlineShapes.Count > 0 || selection.ShapeRange.Count > 0)
            {
                string pictureName = Clipboard.GetText();
                if (string.IsNullOrEmpty(pictureName))
                {
                    MessageBox.Show("Буфер обмена пуст.");
                    return;
                }

                Word.Range range;
                if (selection.InlineShapes.Count > 0)
                {
                    range = selection.InlineShapes[1].Range;
                }
                else
                {
                    range = selection.ShapeRange[1].Anchor;
                }
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                var captionRange = range.Duplicate;
                captionRange.Text = $"Рисунок {doc.InlineShapes.Count + doc.Shapes.Count} — «{pictureName}»";

                Style captionStyle = null;
                try
                {
                    captionStyle = doc.Styles["Подпись картинки"];
                }
                catch
                {
                    MessageBox.Show("Стиль 'Подпись картинки' не найден.");
                    return;
                }

                if (captionStyle != null)
                {
                    captionRange.Paragraphs[1].Range.set_Style(captionStyle);
                }
            }
            else
            {
                MessageBox.Show("Выделенная картинка не найдена.");
            }
        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            setIndent(1.75f, 1.25f);
        }

        #endregion CLICK

        #region UTILITIES
        private void setIndent(float aFirstlineIndent, float aLeftIndent)
        {
            var doc = Doc;

            if (doc == null)
            {
                MessageBox.Show("Нет активного документа.");
                return;
            }

            try
            {
                var selection = App.Selection;
                if (selection.Range.Start == selection.Range.End)
                {
                    MessageBox.Show("Нет выделенного текста.");
                    return;
                }

                foreach (Word.Paragraph paragraph in doc.Paragraphs)
                {
                    var paragraphRange = paragraph.Range;
                    if (paragraphRange.InRange(selection.Range))
                    {
                        // ✅ Правильно: изменяем свойства напрямую
                        paragraph.Format.LeftIndent = App.CentimetersToPoints(aLeftIndent);
                        paragraph.Format.RightIndent = 0;
                        paragraph.Format.FirstLineIndent = App.CentimetersToPoints(aFirstlineIndent);

                        if (paragraph.Range.ListFormat?.List != null)
                        {
                            var listLevel = paragraph.Range.ListFormat.ListTemplate.ListLevels[paragraph.Range.ListFormat.ListLevelNumber];
                            listLevel.NumberPosition = App.CentimetersToPoints(aLeftIndent);
                            listLevel.TextPosition = App.CentimetersToPoints(aLeftIndent + 0.5f);
                            listLevel.TabPosition = App.CentimetersToPoints(aLeftIndent + 0.5f);
                            listLevel.ResetOnHigher = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            // ❌ УДАЛЕНО: finally с Marshal.ReleaseComObject
        }

        private void setIndent()
        {
            var doc = Doc;

            if (doc == null)
            {
                MessageBox.Show("Нет активного документа.");
                return;
            }

            try
            {
                var selection = App.Selection;
                if (selection.Range.Start == selection.Range.End)
                {
                    MessageBox.Show("Нет выделенного текста.");
                    return;
                }

                // ❌ БЫЛО: selection.ParagraphFormat = new Word.ParagraphFormat { ... };
                // ✅ СТАЛО: изменяем свойства напрямую
                selection.ParagraphFormat.LeftIndent = 0;
                selection.ParagraphFormat.RightIndent = 0;
                selection.ParagraphFormat.FirstLineIndent = App.CentimetersToPoints(1.25f);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            // ❌ УДАЛЕНО: finally с Marshal.ReleaseComObject
        }

        #endregion UTILITIES

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Doc;

            if (doc == null)
            {
                MessageBox.Show("Нет активного документа.");
                return;
            }

            // Проверим, есть ли изображение в буфере обмена
            if (!Clipboard.ContainsImage())
            {
                MessageBox.Show("В буфере обмена нет изображения.");
                return;
            }

            // Получаем текущее выделение — вставим картинку туда
            var selection = App.Selection;

            // Вставляем изображение как InlineShape
            Word.InlineShape inlineShape = null;
            try
            {
                inlineShape = selection.InlineShapes.AddPicture(
                    FileName: string.Empty,
                    LinkToFile: Type.Missing,
                    SaveWithDocument: true,
                    Range: selection.Range
                );
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Fallback: попробуем через Paste (иногда AddPicture не работает при пустом Range)
                selection.Paste();
                // После вставки пытаемся найти последнюю картинку
                if (selection.InlineShapes.Count > 0)
                {
                    inlineShape = selection.InlineShapes[selection.InlineShapes.Count];
                }
                else if (doc.InlineShapes.Count > 0)
                {
                    inlineShape = doc.InlineShapes[doc.InlineShapes.Count];
                }
            }

            if (inlineShape == null)
            {
                MessageBox.Show("Не удалось вставить изображение.");
                return;
            }

            // Запрашиваем подпись у пользователя через WPF-диалог
            string captionText = ShowInputBox("Введите подпись к рисунку:", "Подпись");
            if (captionText == null)
                return; // пользователь отменил

            // Позиционируем курсор после картинки
            var rangeAfterPicture = inlineShape.Range.Duplicate;
            rangeAfterPicture.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            rangeAfterPicture.InsertParagraphAfter(); // добавляем абзац
            rangeAfterPicture.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // Формируем текст подписи
            int pictureNumber = doc.InlineShapes.Count + doc.Shapes.Count; // или логика нумерации по вашему выбору
            string fullCaption = $"Рисунок {pictureNumber} — {captionText}";

            // Вставляем подпись
            rangeAfterPicture.Text = fullCaption;

            // Применяем стиль "Подпись картинки"
            Style captionStyle = null;
            try
            {
                captionStyle = doc.Styles["Picture"];
            }
            catch
            {
                MessageBox.Show("Стиль 'Подпись картинки' не найден.");
                return;
            }

            if (captionStyle != null)
            {
                rangeAfterPicture.Paragraphs[1].Range.set_Style(captionStyle);
            }
        }

        // Вспомогательный метод: простой WPF-диалог ввода текста (без System.Windows.Forms)
        private string ShowInputBox(string prompt, string title)
        {
            using (var form = new Form())
            {
                form.Width = 400;
                form.Height = 150;
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.TopMost = true;
                form.MinimizeBox = false;
                form.MaximizeBox = false;

                var label = new Label
                {
                    Text = prompt,
                    Location = new System.Drawing.Point(15, 15),
                    Width = 360,
                    Height = 20
                };

                var textBox = new TextBox
                {
                    Location = new System.Drawing.Point(15, 40),
                    Width = 360,
                    Height = 26
                };

                var okButton = new Button
                {
                    Text = "OK",
                    Location = new System.Drawing.Point(215, 75),
                    Width = 75
                };
                okButton.Click += (_, __) => form.DialogResult = DialogResult.OK;

                var cancelButton = new Button
                {
                    Text = "Отмена",
                    Location = new System.Drawing.Point(295, 75),
                    Width = 75
                };
                cancelButton.Click += (_, __) => form.DialogResult = DialogResult.Cancel;

                textBox.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Enter) form.DialogResult = DialogResult.OK;
                    else if (e.KeyCode == Keys.Escape) form.DialogResult = DialogResult.Cancel;
                };

                form.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
                textBox.Focus();

                if (form.ShowDialog() == DialogResult.OK)
                {
                    return textBox.Text;
                }
                return null;
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = App.Selection;

            // Сбрасываем форматирование выделенного текста к стилю "Обычный"
            ResetToNormalStyle(selection);

            // Применяем маркированный список
            ApplyDashListStyle(selection);

            // Настраиваем отступы (если нужно специфично, иначе список сам задаст стандартные)
            setIndent(1.25f, 0); // Вашу функцию можно вызвать после, если она сбрасывает стили списка
        }

        /// <summary>
        /// Применяет стиль списка с тире (длинным маркером)
        /// </summary>
        private void ApplyDashListStyle(Word.Selection selection)
        {
            if (selection == null || selection.Range == null) return;

            try
            {
                // Получаем коллекцию шаблонов списков из активного документа
                Word.ListTemplates listTemplates = Doc.ListTemplates;

                // Создаем новый шаблон списка или используем существующий. 
                // Проще всего применить встроенный уровень через ListGallery.

                // wdListGalleryTypeBullet = 1 (Маркированные)
                // Index 2 или 3 обычно соответствует тире или ромбу, зависит от версии Word и языка.
                // Самый надежный способ получить именно "Тире" — создать свой простой шаблон.

                Word.ListTemplate lt = listTemplates.Add();

                // Настраиваем первый уровень списка
                Word.ListLevel level = lt.ListLevels[1];

                // Устанавливаем символ маркера: Тире (En Dash или Em Dash)
                // 8212 — это длинное тире (—), 8211 — среднее (–), 45 - обычное (-)

                /*
                 *  Среднее тире (En dash) level.NumberFormat = "\u2013";  Среднее 
                 *  Длинное тире (Em dash) level.NumberFormat = "\u2014";  Самое длинное 
                 *  Дефис (обычный, с клавиатуры) level.NumberFormat = "\u002D";  
                 *  или просто level.NumberFormat = "-";
                 */


                level.NumberFormat = "\u2014";

                // Выравнивание маркера
                level.Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;

                // Отступ текста от маркера (в пунктах)
                level.TextPosition = 25f;

                // Отступ самого маркера от левого края (в пунктах)
                level.NumberPosition = 0f;

                // Применяем этот шаблон к выделению
                selection.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate: lt,
                    ContinuePreviousList: false,
                    ApplyTo: Word.WdListApplyTo.wdListApplyToSelection,
                    DefaultListBehavior: Word.WdDefaultListBehavior.wdWord10ListBehavior
                );
            }
            catch (Exception ex)
            {
            }
        }

        /// <summary>
        /// Сбрасывает форматирование выделенного текста к стилю "Обычный"
        /// </summary>
        private void ResetToNormalStyle(Word.Selection selection)
        {
            if (selection == null || selection.Range == null) return;

            try
            {
                // Получаем стиль "Обычный" из шаблона документа
                Word.Style normalStyle = selection.Document.Styles[Word.WdBuiltinStyle.wdStyleNormal];

                // Применяем стиль к выделенному диапазону
                selection.Range.set_Style(normalStyle);

                // Дополнительно: сбрасываем ручное форматирование (жирный, курсив и т.д.)
                selection.Range.Font.Reset();
                selection.Range.ParagraphFormat.Reset();

                // Убираем маркировку/нумерацию, если она была
                selection.Range.ListFormat.RemoveNumbers(Word.WdNumberType.wdNumberParagraph);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка сброса стиля: " + ex.Message);
            }
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = App.Selection;
            // Сбрасываем форматирование выделенного текста к стилю "Обычный"
            ResetToNormalStyle(selection);
            ApplyNumberedListStyle(selection);
            setIndent(1.25f, 0);
        }

        /// <summary>
        /// Применяет нумерованный список (1., 2., 3.) с настройкой отступов
        /// </summary>
        private void ApplyNumberedListStyle(Word.Selection selection)
        {
            if (selection == null || selection.Range == null) return;

            try
            {
                Word.ListTemplates listTemplates = Doc.ListTemplates;

                // Создаем новый шаблон, чтобы сбросить связь с предыдущими списками
                Word.ListTemplate lt = listTemplates.Add();

                // Настраиваем первый уровень
                Word.ListLevel level = lt.ListLevels[1];

                // 1. Формат номера: "1." (число + точка + пробел)
                // \u0001 - это код для номера списка (арабские цифры 1,2,3...)
                level.NumberFormat = "%1.";

                // 2. Гарантируем начало с 1
                level.StartAt = 1;

                // 3. Настройка отступов (в пунктах, 1 см ≈ 28.35 pt)
                // Положение самого номера (цифры) относительно левого края
                level.NumberPosition = 0f;

                // Положение текста после номера. 
                // Если поставить 1.25 см (≈ 35.4 pt), то будет стандартный абзацный отступ
                level.TextPosition = 35.4f;

                // Выравнивание номера по левому краю
                level.Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;

                // Применяем к выделению
                selection.Range.ListFormat.ApplyListTemplateWithLevel(
                    ListTemplate: lt,
                    ContinuePreviousList: false, // ВАЖНО: не продолжать предыдущий список
                    ApplyTo: Word.WdListApplyTo.wdListApplyToSelection,
                    DefaultListBehavior: Word.WdDefaultListBehavior.wdWord10ListBehavior
                );

                // Дополнительно: можно сбросить отступ первого строки, если он мешает
                selection.ParagraphFormat.FirstLineIndent = 0;
            }
            catch (Exception ex)
            {
            }
        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = App.Selection;

            if (selection == null || selection.Range == null)
            {
                MessageBox.Show("Нет активного выделения.");
                return;
            }

            try
            {
                var doc = Doc;
                if (doc == null)
                {
                    MessageBox.Show("Нет активного документа.");
                    return;
                }

                // Получаем встроенный стиль "Обычный" (Normal)
                // wdStyleNormal = -1 — это константа встроенного стиля
                var normalStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];

                // Применяем стиль к выделению.
                // set_Style ведёт себя идентично штатной кнопке:
                // - если выделен диапазон внутри абзаца — стиль применяется к этому диапазону;
                // - если выделен весь абзац (или несколько) — стиль применяется ко всему абзацу/абзацам;
                // - если выделения нет (просто курсор) — стиль применяется к текущему абзацу.
                selection.Range.set_Style(normalStyle);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при применении стиля: " + ex.Message);
            }
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = App.Selection;
            if (selection == null || selection.Range == null) return;

            try
            {
                var doc = Doc;
                if (doc == null) return;

                var normalStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];

                // Сначала применяем стиль (как штатная кнопка)
                selection.Range.set_Style(normalStyle);

                // Затем сбрасываем ручное форматирование поверх стиля
                // (именно это делает Ctrl+Space для шрифта и Ctrl+Q для абзаца)
                selection.Range.Font.Reset();
                selection.Range.ParagraphFormat.Reset();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void button9_Click_1(object sender, RibbonControlEventArgs e)
        {
            var selection = App.Selection;

            if (selection == null || selection.Range == null)
            {
                MessageBox.Show("Нет выделения.");
                return;
            }

            try
            {
                var range = selection.Range;
                int docBaseStart = range.Start; // абсолютная позиция начала выделения в документе
                string text = range.Text;

                // Собираем список правок: (абсолютная позиция в документе, тип правки)
                // Тип 1 — удалить пробел перед двоеточием
                // Тип 2 — сделать строчной букву после двоеточия (только если есть пробел)
                var edits = new List<(int absPos, int type)>();

                for (int i = 0; i < text.Length; i++)
                {
                    if (text[i] == ':')
                    {
                        int absColonPos = docBaseStart + i;

                        // 1. Если перед двоеточием стоит пробел — убираем его
                        if (i > 0 && text[i - 1] == ' ')
                        {
                            edits.Add((absColonPos - 1, 1));
                        }

                        // 2. ПРОВЕРКА: есть ли хотя бы один пробел сразу после двоеточия?
                        if (i + 1 < text.Length && text[i + 1] == ' ')
                        {
                            // Если пробел есть, ищем первую букву после всех пробелов
                            int j = i + 1;
                            while (j < text.Length && text[j] == ' ')
                            {
                                j++;
                            }

                            // Если нашли заглавную букву — делаем её строчной
                            if (j < text.Length && char.IsUpper(text[j]))
                            {
                                edits.Add((docBaseStart + j, 2));
                            }
                        }
                        // Если text[i + 1] != ' ', блок просто пропускается, ничего не меняется.
                    }
                }

                // Применяем правки строго с конца к началу,
                // чтобы удаление/замена символов не сдвигала позиции ещё не обработанных правок
                edits.Sort((a, b) => b.absPos.CompareTo(a.absPos));

                foreach (var edit in edits)
                {
                    var charRange = Doc.Range(edit.absPos, edit.absPos + 1);

                    if (edit.type == 1)
                    {
                        charRange.Text = ""; // удаляем пробел перед двоеточием
                    }
                    else if (edit.type == 2)
                    {
                        charRange.Text = charRange.Text.ToLower(); // делаем строчной букву после пробела
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        /// <summary>
        /// ✅ ОБНОВЛЕНО: вставка Markdown с поддержкой блоков кода и inline-форматирования
        /// </summary>
        private void button10_Click_1(object sender, RibbonControlEventArgs e)
        {
            var selection = App.Selection;

            if (!Clipboard.ContainsText())
            {
                MessageBox.Show("Буфер обмена не содержит текста.");
                return;
            }

            string markdownText = Clipboard.GetText();
            if (string.IsNullOrEmpty(markdownText))
            {
                MessageBox.Show("Буфер обмена пуст.");
                return;
            }

            try
            {
                var doc = Doc;
                if (doc == null)
                {
                    MessageBox.Show("Нет активного документа.");
                    return;
                }

                string[] lines = markdownText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                int totalLines = lines.Length;

                // ✅ НОВОЕ: Создаем и показываем форму прогресса
                using (var progressForm = new MarkdownProgressForm(totalLines))
                {
                    progressForm.Show();

                    int i = 0;
                    while (i < lines.Length)
                    {
                        string line = lines[i];

                        // ✅ Обновляем прогресс-бар каждые 50 строк (или на последней строке)
                        if (i % 50 == 0 || i == totalLines - 1)
                        {
                            progressForm.UpdateProgress(i, totalLines, "Обработка текста...");
                        }

                        if (string.IsNullOrWhiteSpace(line))
                        {
                            selection.InsertParagraphAfter();
                            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            i++;
                            continue;
                        }

                        if (IsHorizontalRule(line))
                        {
                            i++;
                            continue;
                        }

                        if (line.TrimStart().StartsWith("```"))
                        {
                            progressForm.UpdateProgress(i, totalLines, "Вставка блока кода...");
                            var codeLines = new List<string>();
                            i++;
                            while (i < lines.Length && !lines[i].TrimStart().StartsWith("```"))
                            {
                                codeLines.Add(lines[i]);
                                i++;
                            }
                            if (i < lines.Length) i++;

                            InsertCodeBlock(selection, codeLines);
                            continue;
                        }

                        if (line.TrimStart().StartsWith("|") && line.TrimEnd().EndsWith("|"))
                        {
                            progressForm.UpdateProgress(i, totalLines, "Вставка таблицы...");
                            var tableLines = new List<string>();
                            while (i < lines.Length && lines[i].TrimStart().StartsWith("|") && lines[i].TrimEnd().EndsWith("|"))
                            {
                                tableLines.Add(lines[i]);
                                i++;
                            }

                            InsertMarkdownTable(selection, tableLines);
                            continue;
                        }

                        if (line.StartsWith("#### "))
                        {
                            InsertHeading(selection, line.Substring(5), 4);
                        }
                        else if (line.StartsWith("### "))
                        {
                            InsertHeading(selection, line.Substring(4), 3);
                        }
                        else if (line.StartsWith("## "))
                        {
                            InsertHeading(selection, line.Substring(3), 2);
                        }
                        else if (line.StartsWith("# "))
                        {
                            InsertHeading(selection, line.Substring(2), 1);
                        }
                        else if (Regex.IsMatch(line, @"^\s*\d+\.\s+"))
                        {
                            progressForm.UpdateProgress(i, totalLines, "Вставка нумерованного списка...");
                            var listLines = new List<string>();
                            while (i < lines.Length && (Regex.IsMatch(lines[i], @"^\s*\d+\.\s+") ||
                                   Regex.IsMatch(lines[i], @"^\s*[-*]\s+") ||
                                   (lines[i].StartsWith("  ") && listLines.Count > 0)))
                            {
                                listLines.Add(lines[i]);
                                i++;
                            }

                            InsertNumberedList(selection, listLines);
                            continue;
                        }
                        else if (Regex.IsMatch(line, @"^\s*[-*]\s+"))
                        {
                            progressForm.UpdateProgress(i, totalLines, "Вставка маркированного списка...");
                            var listLines = new List<string>();
                            while (i < lines.Length && (Regex.IsMatch(lines[i], @"^\s*[-*]\s+") ||
                                   (lines[i].StartsWith("  ") && listLines.Count > 0)))
                            {
                                listLines.Add(lines[i]);
                                i++;
                            }

                            InsertBulletedList(selection, listLines);
                            continue;
                        }
                        else if (line.StartsWith("> "))
                                {
                                    InsertBlockquote(selection, line.Substring(2));
                                }
                                // ✅ НОВОЕ: Списки задач (- [ ] или - [x])
                                else if (Regex.IsMatch(line, @"^\s*[-*]\s*\[\s*[xX\s]\s*\]\s*.*"))
                                {
                                    InsertTaskList(selection, line);
                                }
                        else
                        {
                            InsertNormalText(selection, line);
                        }

                        i++;
                    }

                    progressForm.UpdateProgress(totalLines, totalLines, "Завершено!");
                    System.Threading.Thread.Sleep(300); // Небольшая пауза, чтобы пользователь увидел "Завершено"
                } // ✅ Форма автоматически закроется здесь благодаря using
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при вставке Markdown: " + ex.Message);
            }
        }

        private void InsertHeading(Word.Selection selection, string text, int level)
        {
            // ✅ Запоминаем позицию ПЕРЕД вставкой — это надёжнее, чем MoveUp
            int startPos = selection.Range.Start;

            // Вставляем текст
            selection.TypeText(text);

            // Получаем диапазон вставленного текста по сохранённым координатам
            int endPos = selection.Range.Start;
            var range = Doc.Range(startPos, endPos);

            // Применяем стиль заголовка
            Word.WdBuiltinStyle style;
            switch (level)
            {
                case 1: style = Word.WdBuiltinStyle.wdStyleHeading1; break;
                case 2: style = Word.WdBuiltinStyle.wdStyleHeading2; break;
                case 3: style = Word.WdBuiltinStyle.wdStyleHeading3; break;
                case 4: style = Word.WdBuiltinStyle.wdStyleHeading4; break;
                default: style = Word.WdBuiltinStyle.wdStyleHeading1; break;
            }

            range.set_Style(style);

            // ✅ НОВОЕ: применяем inline-форматирование (жирный, курсив, код и т.д.)
            ApplyInlineFormatting(range);

            // Возвращаем курсор в конец вставленного текста
            selection.SetRange(endPos, endPos);

            // Перемещаем курсор в конец
            selection.InsertParagraphAfter();
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        private void InsertNormalText(Word.Selection selection, string text)
        {
            int startPos = selection.Range.Start;
            selection.TypeText(text);
            int endPos = selection.Range.Start;
            var range = Doc.Range(startPos, endPos);

            // ✅ НОВОЕ: применяем inline-форматирование
            ApplyInlineFormatting(range);

            selection.SetRange(endPos, endPos);
            selection.InsertParagraphAfter();
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        private void InsertNumberedList(Word.Selection selection, List<string> listLines)
        {
            // ✅ Запоминаем позицию перед вставкой всего списка
            int startPos = selection.Range.Start;

            // Вставляем все элементы списка
            foreach (var line in listLines)
            {
                // Определяем уровень вложенности (по отступам)
                int indentLevel = 0;
                string cleanLine = line;

                // Подсчитываем отступы (2 или 4 пробела = 1 уровень)
                int spaces = 0;
                foreach (char c in line)
                {
                    if (c == ' ') spaces++;
                    else break;
                }
                indentLevel = spaces / 4; // каждые 4 пробела = 1 уровень

                // Убираем номер и пробелы в начале
                cleanLine = Regex.Replace(line, @"^\s*\d+\.\s+", "");

                selection.TypeText(cleanLine);
                selection.InsertParagraphAfter();
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            // ✅ Выделяем весь вставленный список через запомненные координаты
            int endPos = selection.Range.Start;
            var range = Doc.Range(startPos, endPos);

            // Чтобы вызвать функции, принимающие Selection — временно выделяем диапазон
            range.Select();
            var tempSelection = App.Selection;

            // Применяем стиль нумерованного списка
            ResetToNormalStyle(tempSelection);
            ApplyNumberedListStyle(tempSelection);

            // Настраиваем отступы (если нужно специфично, иначе список сам задаст стандартные)
            setIndent(1.25f, 0);

            // ✅ НОВОЕ: применяем inline-форматирование ПОСЛЕ сброса стилей
            ApplyInlineFormatting(range);

            // Возвращаем курсор в конец списка
            selection.SetRange(endPos, endPos);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        private void InsertBulletedList(Word.Selection selection, List<string> listLines)
        {
            // ✅ Запоминаем позицию перед вставкой всего списка
            int startPos = selection.Range.Start;

            // Вставляем все элементы списка
            foreach (var line in listLines)
            {
                // Определяем уровень вложенности
                int spaces = 0;
                foreach (char c in line)
                {
                    if (c == ' ') spaces++;
                    else break;
                }
                int indentLevel = spaces / 4;

                // Убираем маркер и пробелы в начале
                string cleanLine = Regex.Replace(line, @"^\s*[-*]\s+", "");

                selection.TypeText(cleanLine);
                selection.InsertParagraphAfter();
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            // ✅ Выделяем весь вставленный список через запомненные координаты
            int endPos = selection.Range.Start;
            var range = Doc.Range(startPos, endPos);

            // Чтобы вызвать функции, принимающие Selection — временно выделяем диапазон
            range.Select();
            var tempSelection = App.Selection;

            // Применяем стиль маркированного списка
            ResetToNormalStyle(tempSelection);
            ApplyDashListStyle(tempSelection);

            // Настраиваем отступы (если нужно специфично, иначе список сам задаст стандартные)
            setIndent(1.25f, 0);

            // ✅ НОВОЕ: применяем inline-форматирование ПОСЛЕ сброса стилей
            ApplyInlineFormatting(range);

            // Возвращаем курсор в конец списка
            selection.SetRange(endPos, endPos);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        /// <summary>
        /// ✅ НОВОЕ: Вставка цитаты с левой границей
        /// </summary>
        private void InsertBlockquote(Word.Selection selection, string text)
        {
            int startPos = selection.Range.Start;
            selection.TypeText(text);
            int endPos = selection.Range.Start;
            var range = Doc.Range(startPos, endPos);

            // Применяем inline-форматирование внутри цитаты
            ApplyInlineFormatting(range);

            // Настраиваем левую границу (полоску)
            range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            range.Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
            range.Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray50;

            // Небольшой отступ слева для красоты
            range.ParagraphFormat.LeftIndent = App.CentimetersToPoints(1.25f);
            range.ParagraphFormat.SpaceBefore = 6;
            range.ParagraphFormat.SpaceAfter = 6;

            selection.SetRange(endPos, endPos);
            selection.InsertParagraphAfter();
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        /// <summary>
        /// ✅ НОВОЕ: Вставка списка задач с чекбоксами (☐ / ☑)
        /// </summary>
        private void InsertTaskList(Word.Selection selection, string line)
        {
            int startPos = selection.Range.Start;

            // Определяем состояние чекбокса
            bool isChecked = Regex.IsMatch(line, @"\[\s*[xX]\s*\]");
            string checkboxChar = isChecked ? "☑ " : "☐ "; // Unicode символы

            // Убираем маркдаун-синтаксис задачи, оставляем только текст
            string cleanText = Regex.Replace(line, @"^\s*[-*]\s*\[\s*[xX\s]\s*\]\s*", "");

            selection.TypeText(checkboxChar + cleanText);
            selection.InsertParagraphAfter();
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            int endPos = selection.Range.Start;
            var range = Doc.Range(startPos, endPos);

            // Делаем шрифт для чекбокса чуть крупнее и красивее (Segoe UI Symbol отлично подходит)
            range.Font.Name = "Segoe UI Symbol";

            // Применяем стандартный отступ, как у списков
            range.ParagraphFormat.LeftIndent = App.CentimetersToPoints(1.25f);
            range.ParagraphFormat.FirstLineIndent = App.CentimetersToPoints(-1.25f); // Выступ для маркера

            selection.SetRange(endPos, endPos);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }






        /// <summary>
        /// ✅ НОВОЕ: вставка блока кода ``` ... ``` с Courier New, одинарным интервалом, без отступов
        /// </summary>
        /// <summary>
        /// ✅ ОБНОВЛЕНО: вставка блока кода ``` ... ``` с Courier New, одинарным интервалом, 
        /// без отступов и БЕЗ интервалов до/после абзаца (SpaceAfter = 0).
        /// </summary>
        private void InsertCodeBlock(Word.Selection selection, List<string> codeLines)
        {
            int startPos = selection.Range.Start;

            foreach (var line in codeLines)
            {
                selection.TypeText(line);
                selection.InsertParagraphAfter();
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            int endPos = selection.Range.Start;
            var range = Doc.Range(startPos, endPos);

            // Базовое форматирование шрифта и междустрочного интервала
            range.Font.Name = "Courier New";
            range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

            // ✅ Явно убираем все отступы и интервалы у каждого абзаца внутри блока кода
            foreach (Word.Paragraph p in range.Paragraphs)
            {
                p.Format.FirstLineIndent = 0;
                p.Format.LeftIndent = 0;
                p.Format.RightIndent = 0;

                // ВОТ ЭТО РЕШАЕТ ПРОБЛЕМУ С 8 pt:
                p.Format.SpaceBefore = 0;
                p.Format.SpaceAfter = 0;
            }

            // Возвращаем курсор в конец блока кода и добавляем один пустой абзац после него
            selection.SetRange(endPos, endPos);
            selection.InsertParagraphAfter();
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        /// <summary>
        /// проверяет, является ли строка горизонтальным разделителем (---, ***, ===, ___ и т.д.)
        /// </summary>
        private bool IsHorizontalRule(string line)
        {
            string trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed)) return false;

            // Матчит строки, состоящие только из 3+ одинаковых символов -, *, _ или = с возможными пробелами
            // Примеры: ---, ***, ___, ===, - - -, * * *, --- --- ---
            return Regex.IsMatch(trimmed, @"^\s*([-*_=])(\s*\1){2,}\s*$");
        }

        /// <summary>
        /// надёжная проверка строки-разделителя таблицы (|---|---| или | :---: | ---: |)
        /// </summary>
        private bool IsTableSeparator(string line)
        {
            string trimmed = line.Trim();
            if (!trimmed.StartsWith("|") || !trimmed.EndsWith("|"))
                return false;

            // Разбиваем по | и берём ячейки
            string[] parts = trimmed.Split('|');
            // parts[0] и parts[last] — пустые (до первого и после последнего |)
            if (parts.Length < 3) return false;

            for (int i = 1; i < parts.Length - 1; i++)
            {
                string cell = parts[i].Trim();
                if (cell.Length == 0) return false;
                // Каждая ячейка должна состоять только из -, : и пробелов
                if (!Regex.IsMatch(cell, @"^[\-:\s]+$"))
                    return false;
                // И должен быть хотя бы один дефис
                if (!cell.Contains("-"))
                    return false;
            }

            return true;
        }

        private void InsertMarkdownTable(Word.Selection selection, List<string> tableLines)
        {
            if (tableLines.Count < 2) return; // нужна хотя бы шапка и разделитель

            // Парсим строки таблицы
            var rows = new List<string[]>();

            foreach (var line in tableLines)
            {
                if (IsTableSeparator(line)) continue;

                string[] cells = line.Split('|')
                    .Skip(1)
                    .Take(line.Split('|').Length - 2)
                    .Select(c => c.Trim())
                    .ToArray();

                rows.Add(cells);
            }

            if (rows.Count == 0) return;

            int rowCount = rows.Count;
            int colCount = rows.Max(r => r.Length);

            var range = selection.Range;
            var table = Doc.Tables.Add(range, rowCount, colCount);

            // ✅ НОВОЕ: Явно задаем стандартные черные границы для всей таблицы
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Borders.InsideColor = Word.WdColor.wdColorBlack;
            table.Borders.OutsideColor = Word.WdColor.wdColorBlack;

            // Заполняем таблицу
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < rows[i].Length && j < colCount; j++)
                {
                    var cellRange = table.Cell(i + 1, j + 1).Range;
                    cellRange.Text = rows[i][j];
                    ApplyInlineFormatting(cellRange);
                }
            }

            // Форматируем таблицу (шапка жирным и по центру)
            if (rowCount > 0)
            {
                var headerRow = table.Rows[1];
                foreach (Word.Cell cell in headerRow.Cells)
                {
                    cell.Range.Font.Bold = 1;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }

            // Применяем стиль шрифта ко всей таблице
            table.Select();
            var tableSelection = App.Selection;
            tableSelection.Font.Size = 12;
            tableSelection.Font.Name = "Times New Roman Cyr";
            tableSelection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

            // Перемещаем курсор после таблицы
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.MoveDown(Word.WdUnits.wdLine, 1);
            selection.InsertParagraphAfter();
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        /// <summary>
        /// ✅ НОВОЕ: применяет inline-форматирование Markdown к диапазону:
        /// **bold**, *italic*, __underline__, ~~strikethrough~~, `code`
        /// Работает с конца к началу, чтобы удаление маркеров не сбивало позиции.
        /// </summary>
        /// <summary>
        /// ✅ ОБНОВЛЕНО: применяет inline-форматирование с использованием строгих регулярных выражений,
        /// которые гарантированно не конфликтуют друг с другом (*** vs ** vs *).
        /// </summary>
        /// <summary>
        /// ✅ ПУЛЕНЕПРОБИВАЕМАЯ ВЕРСИЯ: применяет inline-форматирование.
        /// Использует строгие regex и безопасное удаление маркеров через .Delete().
        /// </summary>
        /// <summary>
        /// ✅ ФИНАЛЬНАЯ, ПУЛЕНЕПРОБИВАЕМАЯ ВЕРСИЯ.
        /// Использует одно единое регулярное выражение с именованными группами.
        /// Приоритет строго задан порядком: *** -> ** -> *
        /// Это физически исключает ситуацию, когда ** срабатывает внутри ***.
        /// </summary>
        /// <summary>
        /// Использует одно единое регулярное выражение с именованными группами.
        /// Приоритет строго задан порядком: *** -> ** -> *
        /// Это физически исключает ситуацию, когда ** срабатывает внутри ***.
        /// </summary>
        /// <summary>
        /// Использует одно единое регулярное выражение с именованными группами.
        /// Приоритет строго задан порядком: *** -> ** -> *
        /// Это физически исключает ситуацию, когда ** срабатывает внутри ***.
        /// </summary>
        private void ApplyInlineFormatting(Word.Range range)
        {
            if (range == null || string.IsNullOrEmpty(range.Text)) return;

            int rangeStart = range.Start;
            string text = range.Text;

            var edits = new List<InlineEdit>();

            // ✅ ОБНОВЛЕННЫЙ REGEX: добавлены Link и Image. Порядок приоритета сохранен!
            string pattern = @"(?<code>`(?<codeInner>[^`]+)`)|" +
                             @"(?<bolditalic>\*{3}(?<biInner>.*?)\*{3})|" +
                             @"(?<bold>\*{2}(?<bInner>.*?)\*{2})|" +
                             @"(?<italic>\*(?<iInner>[^*]*?)\*)|" +
                             @"(?<underline>_{2}(?<uInner>.*?)_{2})|" +
                             @"(?<strike>~{2}(?<sInner>.*?)~{2})|" +
                             @"(?<link>\[(?<linkText>[^\]]*)\]\((?<linkUrl>[^)]+)\))|" +
                             @"(?<image>!\[(?<imgAlt>[^\]]*)\]\((?<imgUrl>[^)]+)\))";

            foreach (Match m in Regex.Matches(text, pattern, RegexOptions.Singleline))
            {
                InlineFormat format = InlineFormat.None;
                int innerStart = 0, innerEnd = 0;
                string extraData = ""; // Для хранения URL ссылки или картинки

                if (m.Groups["code"].Success)
                {
                    format = InlineFormat.Code;
                    innerStart = m.Groups["codeInner"].Index;
                    innerEnd = innerStart + m.Groups["codeInner"].Length;
                }
                else if (m.Groups["bolditalic"].Success)
                {
                    format = InlineFormat.BoldItalic;
                    innerStart = m.Groups["biInner"].Index;
                    innerEnd = innerStart + m.Groups["biInner"].Length;
                }
                else if (m.Groups["bold"].Success)
                {
                    format = InlineFormat.Bold;
                    innerStart = m.Groups["bInner"].Index;
                    innerEnd = innerStart + m.Groups["bInner"].Length;
                }
                else if (m.Groups["italic"].Success)
                {
                    format = InlineFormat.Italic;
                    innerStart = m.Groups["iInner"].Index;
                    innerEnd = innerStart + m.Groups["iInner"].Length;
                }
                else if (m.Groups["underline"].Success)
                {
                    format = InlineFormat.Underline;
                    innerStart = m.Groups["uInner"].Index;
                    innerEnd = innerStart + m.Groups["uInner"].Length;
                }
                else if (m.Groups["strike"].Success)
                {
                    format = InlineFormat.Strikethrough;
                    innerStart = m.Groups["sInner"].Index;
                    innerEnd = innerStart + m.Groups["sInner"].Length;
                }
                else if (m.Groups["link"].Success)
                {
                    format = InlineFormat.Link;
                    innerStart = m.Groups["linkText"].Index;
                    innerEnd = innerStart + m.Groups["linkText"].Length;
                    extraData = m.Groups["linkUrl"].Value;
                }
                else if (m.Groups["image"].Success)
                {
                    format = InlineFormat.Image;
                    // Для картинки мы заменяем весь матч на изображение, innerStart/End не так важны, 
                    // но оставим их для совместимости логики удаления маркеров
                    innerStart = m.Index;
                    innerEnd = m.Index + m.Length;
                    extraData = m.Groups["imgUrl"].Value;
                }

                if (format != InlineFormat.None)
                {
                    edits.Add(new InlineEdit
                    {
                        fullStart = rangeStart + m.Index,
                        fullEnd = rangeStart + m.Index + m.Length,
                        innerStart = rangeStart + innerStart,
                        innerEnd = rangeStart + innerEnd,
                        format = format,
                        extraData = extraData
                    });
                }
            }

            if (edits.Count == 0) return;

            edits.Sort((a, b) => b.fullStart.CompareTo(a.fullStart));

            foreach (var edit in edits)
            {
                try
                {
                    var innerRange = Doc.Range(edit.innerStart, edit.innerEnd);

                    switch (edit.format)
                    {
                        case InlineFormat.Bold:
                            innerRange.Font.Bold = 1;
                            ApplyInlineFormatting(innerRange); // ✅ РЕКУРСИЯ для вложенности!
                            break;
                        case InlineFormat.Italic:
                            innerRange.Font.Italic = 1;
                            ApplyInlineFormatting(innerRange); // ✅ РЕКУРСИЯ
                            break;
                        case InlineFormat.BoldItalic:
                            innerRange.Font.Bold = 1;
                            innerRange.Font.Italic = 1;
                            ApplyInlineFormatting(innerRange); // ✅ РЕКУРСИЯ
                            break;
                        case InlineFormat.Underline:
                            innerRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                            break;
                        case InlineFormat.Strikethrough:
                            innerRange.Font.StrikeThrough = 1;
                            break;
                        case InlineFormat.Code:
                            innerRange.Font.Name = "Courier New";
                            innerRange.Font.Size = 11;
                            innerRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                            break;
                        case InlineFormat.Link:
                            innerRange.Text = edit.extraData; // Сначала вставляем URL как текст
                                                              // Создаем гиперссылку поверх этого текста
                            Doc.Hyperlinks.Add(innerRange, edit.extraData, Type.Missing, Type.Missing, edit.extraData, Type.Missing);
                            innerRange.Font.Color = Word.WdColor.wdColorBlue;
                            innerRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                            break;
                        case InlineFormat.Image:
                            try
                            {
                                // Пытаемся вставить картинку по URL. 
                                // LinkToFile: false, SaveWithDocument: true
                                var shape = innerRange.InlineShapes.AddPicture(edit.extraData, false, true);
                                shape.AlternativeText = "Markdown Image";
                            }
                            catch
                            {
                                // Если Word блокирует URL или он невалидный, вставляем заглушку
                                innerRange.Text = $"[🖼 Изображение: {edit.extraData}]";
                                innerRange.Font.Color = Word.WdColor.wdColorRed;
                                innerRange.Font.Italic = 1;
                            }
                            break;
                    }

                    // Удаляем маркеры (для Link и Image это удалит исходный markdown-синтаксис)
                    var endMarker = Doc.Range(edit.innerEnd, edit.fullEnd);
                    endMarker.Delete();

                    var startMarker = Doc.Range(edit.fullStart, edit.innerStart);
                    startMarker.Delete();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Markdown Inline Error] {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ✅ НОВОЕ: Простая форма с прогресс-баром для отображения процесса вставки
        /// </summary>
        public class MarkdownProgressForm : Form
        {
            public ProgressBar ProgressBar { get; private set; }
            public Label StatusLabel { get; private set; }

            public MarkdownProgressForm(int maxLines)
            {
                this.Width = 400;
                this.Height = 120;
                this.Text = "Вставка Markdown...";
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.TopMost = true;
                this.ControlBox = false; // Убираем крестик, чтобы пользователь не закрыл его случайно

                StatusLabel = new Label
                {
                    Text = "Подготовка...",
                    Location = new System.Drawing.Point(15, 15),
                    Width = 350,
                    Height = 20,
                    Font = new System.Drawing.Font("Segoe UI", 9f)
                };

                ProgressBar = new ProgressBar
                {
                    Location = new System.Drawing.Point(15, 45),
                    Width = 350,
                    Height = 20,
                    Minimum = 0,
                    Maximum = maxLines > 0 ? maxLines : 100,
                    Style = ProgressBarStyle.Continuous
                };

                this.Controls.Add(StatusLabel);
                this.Controls.Add(ProgressBar);
            }

            public void UpdateProgress(int currentLine, int totalLines, string currentAction)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => UpdateProgress(currentLine, totalLines, currentAction)));
                    return;
                }

                ProgressBar.Value = Math.Min(currentLine, ProgressBar.Maximum);
                StatusLabel.Text = $"{currentAction} (строка {currentLine} из {totalLines})";

                // ✅ Ключевой момент для VSTO: позволяем форме перерисоваться и Word'у обработать сообщения
                System.Windows.Forms.Application.DoEvents();
            }
        }
    }
}
