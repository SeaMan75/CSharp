using System;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace Styles
{
    public partial class StylesN_H1_H2
    {
        private void StylesN_H1_H2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonStyles_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = null;
            Word.Style normalStyle = null;
            Style heading1Style = null;
            Word.Style heading2Style = null;

            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("Нет активного документа.");
                    return;
                }

                // Проверка существования стиля "Normal"

                normalStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
                if (normalStyle != null)
                {
                    normalStyle.Font = new Word.Font
                    {
                        Name = "Times New Roman Cyr",
                        Size = 14,
                        Bold = 0,
                        Color = Word.WdColor.wdColorAutomatic
                    };
                    normalStyle.ParagraphFormat = new Word.ParagraphFormat
                    {
                        FirstLineIndent = Globals.ThisAddIn.Application.CentimetersToPoints(1.25f),
                        LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5,
                        SpaceBefore = 0,
                        SpaceAfter = 0,
                        Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    };
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Стиль 'Обычный' не найден.");
                }

                // Проверка существования стиля "Heading 1"
                heading1Style = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading1];
                if (heading1Style != null)
                {
                    heading1Style.Font = new Word.Font
                    {
                        Name = "Times New Roman Cyr",
                        Color = Word.WdColor.wdColorAutomatic,
                        Size = 14,
                        Bold = 1
                    };

                    heading1Style.ParagraphFormat = new Word.ParagraphFormat
                    {
                        SpaceBefore = 0,
                        SpaceAfter = 28,
                        LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5,
                        Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                        FirstLineIndent = 0,
                        PageBreakBefore = -1
                    };
                    heading1Style.set_BaseStyle("");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Стиль 'Заголовок 1' не найден.");
                }

                heading2Style = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading2];
                if (heading2Style != null)
                {
                    heading2Style.Font = new Word.Font
                    {
                        Name = "Times New Roman Cyr",
                        Size = 14,
                        Bold = 1,
                        Color = Word.WdColor.wdColorAutomatic
                    };
                    heading2Style.ParagraphFormat = new Word.ParagraphFormat
                    {
                        SpaceBefore = 0,
                        SpaceAfter = 28,
                        LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5,
                        Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                        FirstLineIndent = 0,
                        PageBreakBefore = 0
                    };
                    heading2Style.set_BaseStyle("");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Стиль 'Заголовок 2' не найден.");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            finally
            {
                if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                if (heading1Style != null) Marshal.ReleaseComObject(heading1Style);
                if (heading2Style != null) Marshal.ReleaseComObject(heading2Style);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = null;
            Style normalStyle = null;
            Style heading1Style = null;
            Style heading2Style = null;
                    
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("Нет активного документа.");
                    return;
                }
                // Установка отступов и первой строки для всего выделенного текста
                Selection selection = Globals.ThisAddIn.Application.Selection;
                selection.ParagraphFormat = new Word.ParagraphFormat
                {
                    LeftIndent = 0,
                    RightIndent = 0,
                    FirstLineIndent = Globals.ThisAddIn.Application.CentimetersToPoints((float)1.25)
                };
            }


            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            finally
            {
                if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                if (heading1Style != null) Marshal.ReleaseComObject(heading1Style);
                if (heading2Style != null) Marshal.ReleaseComObject(heading2Style);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = null;
            Style normalStyle = null;
            Style heading1Style = null;
            Style heading2Style = null;

            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("Нет активного документа.");
                    return;
                }
                // Установка полей документа
                doc.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
                doc.PageSetup.BottomMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
                doc.PageSetup.LeftMargin = Globals.ThisAddIn.Application.CentimetersToPoints(3);
                doc.PageSetup.RightMargin = Globals.ThisAddIn.Application.CentimetersToPoints((float)1.5);
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            finally
            {
                if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                if (heading1Style != null) Marshal.ReleaseComObject(heading1Style);
                if (heading2Style != null) Marshal.ReleaseComObject(heading2Style);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = null;
            Style normalStyle = null;
            Style heading1Style = null;
            Style heading2Style = null;

            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("Нет активного документа.");
                    return;
                }
                // Установка полей документа
                doc.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
                doc.PageSetup.BottomMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
                doc.PageSetup.LeftMargin = Globals.ThisAddIn.Application.CentimetersToPoints(3);
                doc.PageSetup.RightMargin = Globals.ThisAddIn.Application.CentimetersToPoints(1);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            finally
            {
                if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                if (heading1Style != null) Marshal.ReleaseComObject(heading1Style);
                if (heading2Style != null) Marshal.ReleaseComObject(heading2Style);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = null;
            Style normalStyle = null;
            Style heading1Style = null;
            Style heading2Style = null;

            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("Нет активного документа.");
                    return;
                }

                Style pictureStyle = null;

                // Проверка и настройка стиля "Картинка"
                try
                {
                    pictureStyle = doc.Styles["Картинка"];
                }
                catch
                {
                    pictureStyle = doc.Styles.Add("Картинка", Word.WdStyleType.wdStyleTypeParagraph);
                }

                pictureStyle.ParagraphFormat = new Word.ParagraphFormat
                {
                    Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                    SpaceBefore = 28,
                    SpaceAfter = 0,
                    LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5,
                    FirstLineIndent = 0
                };
                pictureStyle.set_BaseStyle("");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            finally
            {
                if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                if (heading1Style != null) Marshal.ReleaseComObject(heading1Style);
                if (heading2Style != null) Marshal.ReleaseComObject(heading2Style);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = null;
            Style normalStyle = null;
            Style heading1Style = null;
            Style heading2Style = null;

            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("Нет активного документа.");
                    return;
                }

                Style pictureCaptionStyle = null;

                try
                {
                    pictureCaptionStyle = doc.Styles["Подпись картинки"];
                }
                catch
                {
                    pictureCaptionStyle = doc.Styles.Add("Подпись картинки", Word.WdStyleType.wdStyleTypeParagraph);
                }

                pictureCaptionStyle.ParagraphFormat = new Word.ParagraphFormat
                {
                    Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                    SpaceBefore = 0,
                    SpaceAfter = 28,
                    FirstLineIndent = 0,
                    LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5
                };

                pictureCaptionStyle.set_BaseStyle("");
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            finally
            {
                if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                if (heading1Style != null) Marshal.ReleaseComObject(heading1Style);
                if (heading2Style != null) Marshal.ReleaseComObject(heading2Style);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = null;
            Style normalStyle = null;
            Style heading1Style = null;
            Style heading2Style = null;

            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (doc == null)
                {
                    System.Windows.Forms.MessageBox.Show("Нет активного документа.");
                    return;
                }

                Style tableCaptionStyle = null;

                try
                {
                    tableCaptionStyle = doc.Styles["Подпись таблицы"];
                }
                catch
                {
                    tableCaptionStyle = doc.Styles.Add("Подпись таблицы", Word.WdStyleType.wdStyleTypeParagraph);
                }

                tableCaptionStyle.ParagraphFormat = new Word.ParagraphFormat
                {
                    Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                    SpaceBefore = 28,
                    SpaceAfter = 0,
                    FirstLineIndent = Globals.ThisAddIn.Application.CentimetersToPoints((float)1.25),
                    LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5
                };
                tableCaptionStyle.set_BaseStyle("");
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
            finally
            {
                if (normalStyle != null) Marshal.ReleaseComObject(normalStyle);
                if (heading1Style != null) Marshal.ReleaseComObject(heading1Style);
                if (heading2Style != null) Marshal.ReleaseComObject(heading2Style);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }
    }
}
