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
            Bold,
            Italic,
            BoldItalic,
            Underline,
            Strikethrough,
            Code
        }

        // ✅ НОВОЕ: структура для хранения найденной inline-правки
        private class InlineEdit
        {
            public int fullStart;   // начало всего match (включая маркеры)
            public int fullEnd;     // конец всего match (включая маркеры)
            public int innerStart;  // начало внутреннего текста
            public int innerEnd;    // конец внутреннего текста
            public InlineFormat format;
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

                // Разбиваем текст на строки
                string[] lines = markdownText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                int i = 0;
                while (i < lines.Length)
                {
                    string line = lines[i];

                    // Пропускаем пустые строки (но добавляем абзац)
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        selection.InsertParagraphAfter();
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        i++;
                        continue;
                    }

                    // пропускаем горизонтальные разделители (---, ***, ===, ___)
                    if (IsHorizontalRule(line))
                    {
                        i++;
                        continue;
                    }

                    // проверяем блок кода ``` (может быть с указанием языка, например ```csharp)
                    if (line.TrimStart().StartsWith("```"))
                    {
                        var codeLines = new List<string>();
                        i++; // пропускаем открывающий ``` (вместе с языком — он игнорируется)
                        while (i < lines.Length && !lines[i].TrimStart().StartsWith("```"))
                        {
                            codeLines.Add(lines[i]);
                            i++;
                        }
                        if (i < lines.Length) i++; // пропускаем закрывающий ```
                        
                        InsertCodeBlock(selection, codeLines);
                        continue;
                    }

                    // Проверяем, является ли это таблицей Markdown
                    if (line.TrimStart().StartsWith("|") && line.TrimEnd().EndsWith("|"))
                    {
                        // Собираем все строки таблицы
                        var tableLines = new List<string>();
                        while (i < lines.Length && lines[i].TrimStart().StartsWith("|") && lines[i].TrimEnd().EndsWith("|"))
                        {
                            tableLines.Add(lines[i]);
                            i++;
                        }

                        // Вставляем таблицу
                        InsertMarkdownTable(selection, tableLines);
                        continue;
                    }

                    // Проверяем заголовки
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
                    // Проверяем нумерованный список (1. , 2. , и т.д.)
                    else if (Regex.IsMatch(line, @"^\s*\d+\.\s+"))
                    {
                        // Собираем все строки списка (включая вложенные)
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
                    // Проверяем маркированный список (- , * )
                    else if (Regex.IsMatch(line, @"^\s*[-*]\s+"))
                    {
                        // Собираем все строки списка (включая вложенные)
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
                    // Обычный текст
                    else
                    {
                        InsertNormalText(selection, line);
                    }

                    i++;
                }
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
        /// ✅ НОВОЕ: вставка блока кода ``` ... ``` с Courier New, одинарным интервалом, без отступов
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

            // Форматирование блока кода
            range.Font.Name = "Courier New";
            range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            
            // Убираем отступы у всех абзацев блока
            foreach (Word.Paragraph p in range.Paragraphs)
            {
                p.Format.FirstLineIndent = 0;
                p.Format.LeftIndent = 0;
            }

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
                // ✅ ИСПРАВЛЕНО: используем надёжную проверку строки-разделителя
                if (IsTableSeparator(line)) continue;

                // Разбиваем строку по |
                string[] cells = line.Split('|')
                    .Skip(1) // пропускаем первый пустой элемент
                    .Take(line.Split('|').Length - 2) // пропускаем последний пустой элемент
                    .Select(c => c.Trim())
                    .ToArray();

                rows.Add(cells);
            }

            if (rows.Count == 0) return;

            int rowCount = rows.Count;
            int colCount = rows.Max(r => r.Length);

            // Создаём таблицу в Word
            var range = selection.Range;
            var table = Doc.Tables.Add(range, rowCount, colCount);

            // Заполняем таблицу
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < rows[i].Length && j < colCount; j++)
                {
                    var cellRange = table.Cell(i + 1, j + 1).Range;
                    cellRange.Text = rows[i][j];
                    // ✅ НОВОЕ: применяем inline-форматирование к содержимому ячейки
                    ApplyInlineFormatting(cellRange);
                }
            }

            // Форматируем таблицу (шапка жирным)
            if (rowCount > 0)
            {
                var headerRow = table.Rows[1];
                foreach (Word.Cell cell in headerRow.Cells)
                {
                    cell.Range.Font.Bold = 1;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }

            // Применяем стиль ко всей таблице
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
        private void ApplyInlineFormatting(Word.Range range)
        {
            int rangeStart = range.Start;
            string text = range.Text;

            if (string.IsNullOrEmpty(text)) return;

            var edits = new List<InlineEdit>();

            // ✅ Inline code: `text` (высший приоритет — обрабатывается первым)
            foreach (Match m in Regex.Matches(text, @"(?<!`)`([^`]+)`(?!`)"))
            {
                edits.Add(new InlineEdit
                {
                    fullStart = rangeStart + m.Index,
                    fullEnd = rangeStart + m.Index + m.Length,
                    innerStart = rangeStart + m.Groups[1].Index,
                    innerEnd = rangeStart + m.Groups[1].Index + m.Groups[1].Length,
                    format = InlineFormat.Code
                });
            }

            // Bold+Italic: ***text*** (обрабатывается ПЕРЕД одиночными ** и *)
            foreach (Match m in Regex.Matches(text, @"\*\*\*(.+?)\*\*\*", RegexOptions.Singleline))
            {
                edits.Add(new InlineEdit
                {
                    fullStart = rangeStart + m.Index,
                    fullEnd = rangeStart + m.Index + m.Length,
                    innerStart = rangeStart + m.Groups[1].Index,
                    innerEnd = rangeStart + m.Groups[1].Index + m.Groups[1].Length,
                    format = InlineFormat.BoldItalic
                });
            }


            // ✅ Bold: **text**
            foreach (Match m in Regex.Matches(text, @"\*\*(.+?)\*\*", RegexOptions.Singleline))
            {
                edits.Add(new InlineEdit
                {
                    fullStart = rangeStart + m.Index,
                    fullEnd = rangeStart + m.Index + m.Length,
                    innerStart = rangeStart + m.Groups[1].Index,
                    innerEnd = rangeStart + m.Groups[1].Index + m.Groups[1].Length,
                    format = InlineFormat.Bold
                });
            }

            // ✅ Underline: __text__
            foreach (Match m in Regex.Matches(text, @"__(.+?)__", RegexOptions.Singleline))
            {
                edits.Add(new InlineEdit
                {
                    fullStart = rangeStart + m.Index,
                    fullEnd = rangeStart + m.Index + m.Length,
                    innerStart = rangeStart + m.Groups[1].Index,
                    innerEnd = rangeStart + m.Groups[1].Index + m.Groups[1].Length,
                    format = InlineFormat.Underline
                });
            }

            // ✅ Strikethrough: ~~text~~
            foreach (Match m in Regex.Matches(text, @"~~(.+?)~~", RegexOptions.Singleline))
            {
                edits.Add(new InlineEdit
                {
                    fullStart = rangeStart + m.Index,
                    fullEnd = rangeStart + m.Index + m.Length,
                    innerStart = rangeStart + m.Groups[1].Index,
                    innerEnd = rangeStart + m.Groups[1].Index + m.Groups[1].Length,
                    format = InlineFormat.Strikethrough
                });
            }

            // ✅ Italic: *text* (но не **, с negative lookbehind/lookahead)
            foreach (Match m in Regex.Matches(text, @"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", RegexOptions.Singleline))
            {
                edits.Add(new InlineEdit
                {
                    fullStart = rangeStart + m.Index,
                    fullEnd = rangeStart + m.Index + m.Length,
                    innerStart = rangeStart + m.Groups[1].Index,
                    innerEnd = rangeStart + m.Groups[1].Index + m.Groups[1].Length,
                    format = InlineFormat.Italic
                });
            }

            if (edits.Count == 0) return;

            // Сортируем с конца к началу, чтобы удаление маркеров не сбивало позиции
            edits.Sort((a, b) => b.fullStart.CompareTo(a.fullStart));

            // Убираем перекрывающиеся правки (например, если ** оказался внутри ```)
            // Code имеет высший приоритет — вложенные в него маркеры отсеиваются
            var validEdits = new List<InlineEdit>();
            foreach (var edit in edits)
            {
                bool overlaps = false;
                foreach (var existing in validEdits)
                {
                    // Проверка перекрытия интервалов
                    if (edit.fullStart < existing.fullEnd && edit.fullEnd > existing.fullStart)
                    {
                        overlaps = true;
                        break;
                    }
                }
                if (!overlaps)
                {
                    validEdits.Add(edit);
                }
            }

            // Применяем правки
            foreach (var edit in validEdits)
            {
                try
                {
                    // Форматирование внутреннего текста
                    var innerRange = Doc.Range(edit.innerStart, edit.innerEnd);

                    switch (edit.format)
                    {
                        case InlineFormat.Bold:
                            innerRange.Font.Bold = 1;
                            break;
                        case InlineFormat.Italic:
                            innerRange.Font.Italic = 1;
                            break;
                        case InlineFormat.BoldItalic:   
                            innerRange.Font.Bold = 1;
                            innerRange.Font.Italic = 1;
                            break;
                        case InlineFormat.Underline:
                            innerRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                            break;
                        case InlineFormat.Strikethrough:
                            innerRange.Font.StrikeThrough = 1;
                            break;
                        case InlineFormat.Code:
                            innerRange.Font.Name = "Courier New";
                            innerRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                            break;
                    }

                    // Удаляем конечный маркер (сначала, т.к. он дальше в документе)
                    var endMarker = Doc.Range(edit.innerEnd, edit.fullEnd);
                    endMarker.Text = "";

                    // Удаляем начальный маркер
                    var startMarker = Doc.Range(edit.fullStart, edit.innerStart);
                    startMarker.Text = "";
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Inline format error: {ex.Message}");
                }
            }
        }
    }

}