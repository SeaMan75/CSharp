namespace Styles
{
    partial class StylesN_H1_H2 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public StylesN_H1_H2()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.groupStyles = this.Factory.CreateRibbonGroup();
            this.buttonStyles = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.groupFormat = this.Factory.CreateRibbonGroup();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button15 = this.Factory.CreateRibbonButton();
            this.button16 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button13 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.groupStyles.SuspendLayout();
            this.groupFormat.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.groupStyles);
            this.tab2.Groups.Add(this.groupFormat);
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Label = "Разное";
            this.tab2.Name = "tab2";
            // 
            // groupStyles
            // 
            this.groupStyles.Items.Add(this.buttonStyles);
            this.groupStyles.Items.Add(this.button1);
            this.groupStyles.Items.Add(this.button12);
            this.groupStyles.Items.Add(this.button2);
            this.groupStyles.Items.Add(this.button3);
            this.groupStyles.Label = "Стили";
            this.groupStyles.Name = "groupStyles";
            // 
            // buttonStyles
            // 
            this.buttonStyles.Label = "Настройка стилей";
            this.buttonStyles.Name = "buttonStyles";
            this.buttonStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStyles_Click);
            // 
            // button1
            // 
            this.button1.Label = "Отступы у выделенного 1.25;  0";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button12
            // 
            this.button12.Label = "Отступы у выделенного 2.50;  1.25";
            this.button12.Name = "button12";
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button12_Click);
            // 
            // button2
            // 
            this.button2.Label = "Поля: 2/2/3/1.5";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Label = "Поля 2/2/3/1";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // groupFormat
            // 
            this.groupFormat.Items.Add(this.button8);
            this.groupFormat.Items.Add(this.button5);
            this.groupFormat.Items.Add(this.button6);
            this.groupFormat.Label = "Форматирование";
            this.groupFormat.Name = "groupFormat";
            // 
            // button8
            // 
            this.button8.Label = "Формат таблицы";
            this.button8.Name = "button8";
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // button5
            // 
            this.button5.Label = "Список";
            this.button5.Name = "button5";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Label = "Нумерованный список";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button14);
            this.group2.Items.Add(this.button15);
            this.group2.Items.Add(this.button16);
            this.group2.Items.Add(this.button4);
            this.group2.Label = "Подписи";
            this.group2.Name = "group2";
            // 
            // button14
            // 
            this.button14.Label = "Подписать таблицу";
            this.button14.Name = "button14";
            // 
            // button15
            // 
            this.button15.Label = "Подписать картинку с учетом заголовков";
            this.button15.Name = "button15";
            this.button15.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // button16
            // 
            this.button16.Label = "Подписать картинку без учета заголовков";
            this.button16.Name = "button16";
            this.button16.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button11_Click);
            // 
            // button4
            // 
            this.button4.Label = "Вставить из буфера и подписать";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button13);
            this.group3.Items.Add(this.button7);
            this.group3.Items.Add(this.button9);
            this.group3.Items.Add(this.button10);
            this.group3.Label = "Форматирование";
            this.group3.Name = "group3";
            // 
            // button13
            // 
            this.button13.Label = "Стиль обычный";
            this.button13.Name = "button13";
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // button7
            // 
            this.button7.Label = "Стиль обычный со сбросом";
            this.button7.Name = "button7";
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button9
            // 
            this.button9.Label = "Двоеточие";
            this.button9.Name = "button9";
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click_1);
            // 
            // button10
            // 
            this.button10.Label = "Вставить из буфера Markdown";
            this.button10.Name = "button10";
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click_1);
            // 
            // StylesN_H1_H2
            // 
            this.Name = "StylesN_H1_H2";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.StylesN_H1_H2_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.groupStyles.ResumeLayout(false);
            this.groupStyles.PerformLayout();
            this.groupFormat.ResumeLayout(false);
            this.groupFormat.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button15;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button16;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
    }

    partial class ThisRibbonCollection
    {
        internal StylesN_H1_H2 StylesN_H1_H2
        {
            get { return this.GetRibbon<StylesN_H1_H2>(); }
        }
    }
}
