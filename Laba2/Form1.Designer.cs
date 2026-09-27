namespace Laba2 {
  partial class Form1 {
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing) {
      if (disposing && (components != null))
        components.Dispose();
      base.Dispose(disposing);
    }

    private void InitializeComponent() {
      menuStrip1 = new MenuStrip();
      miDataMenu = new ToolStripMenuItem();
      miGenerate = new ToolStripMenuItem();
      miLoadExcel = new ToolStripMenuItem();
      miLoadGoogle = new ToolStripMenuItem();
      miActionsMenu = new ToolStripMenuItem();
      miCalculate = new ToolStripMenuItem();
      miClear = new ToolStripMenuItem();
      miExit = new ToolStripMenuItem();
      dataGridView1 = new DataGridView();
      panel1 = new Panel();
      chkBubble = new CheckBox();
      chkInsert = new CheckBox();
      chkShaker = new CheckBox();
      chkQuick = new CheckBox();
      chkBogo = new CheckBox();
      gbDirection = new GroupBox();
      rbAsc = new RadioButton();
      rbDesc = new RadioButton();
      numCount = new NumericUpDown();
      lblCountCaption = new Label();
      lblBubbleCaption = new Label();
      lblInsertCaption = new Label();
      lblShakerCaption = new Label();
      lblQuickCaption = new Label();
      lblBogoCaption = new Label();
      lblBubbleTime = new Label();
      lblInsertTime = new Label();
      lblShakerTime = new Label();
      lblQuickTime = new Label();
      lblBogoTime = new Label();
      menuStrip1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
      gbDirection.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)numCount).BeginInit();
      SuspendLayout();

      menuStrip1.ImageScalingSize = new Size(20, 20);
      menuStrip1.Items.AddRange(new ToolStripItem[] { miDataMenu, miActionsMenu });
      menuStrip1.Location = new Point(0, 0);
      menuStrip1.Name = "menuStrip1";
      menuStrip1.Size = new Size(1184, 28);
      menuStrip1.TabIndex = 0;

      miDataMenu.DropDownItems.AddRange(new ToolStripItem[] { miGenerate, miLoadExcel, miLoadGoogle });
      miDataMenu.Name = "miDataMenu";
      miDataMenu.Size = new Size(78, 24);
      miDataMenu.Text = "Данные";

      miGenerate.Name = "miGenerate";
      miGenerate.Size = new Size(272, 26);
      miGenerate.Text = "Сгенерировать";
      miGenerate.Click += miGenerate_Click;

      miLoadExcel.Name = "miLoadExcel";
      miLoadExcel.Size = new Size(272, 26);
      miLoadExcel.Text = "Загрузить из Excel";
      miLoadExcel.Click += miLoadExcel_Click;

      miLoadGoogle.Name = "miLoadGoogle";
      miLoadGoogle.Size = new Size(272, 26);
      miLoadGoogle.Text = "Загрузить из Google Table";
      miLoadGoogle.Click += miLoadGoogle_Click;

      miActionsMenu.DropDownItems.AddRange(new ToolStripItem[] { miCalculate, miClear, miExit });
      miActionsMenu.Name = "miActionsMenu";
      miActionsMenu.Size = new Size(88, 24);
      miActionsMenu.Text = "Действия";

      miCalculate.Name = "miCalculate";
      miCalculate.Size = new Size(167, 26);
      miCalculate.Text = "Рассчитать";
      miCalculate.Click += miCalculate_Click;

      miClear.Name = "miClear";
      miClear.Size = new Size(167, 26);
      miClear.Text = "Очистить";
      miClear.Click += miClear_Click;

      miExit.Name = "miExit";
      miExit.Size = new Size(167, 26);
      miExit.Text = "Выход";
      miExit.Click += miExit_Click;

      dataGridView1.ColumnHeadersHeight = 29;
      dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
      dataGridView1.Location = new Point(12, 40);
      dataGridView1.Name = "dataGridView1";
      dataGridView1.RowHeadersWidth = 51;
      dataGridView1.Size = new Size(260, 380);
      dataGridView1.TabIndex = 1;
      dataGridView1.CellValidating += DataGridView1_CellValidating;

      panel1.BackColor = Color.White;
      panel1.BorderStyle = BorderStyle.FixedSingle;
      panel1.Location = new Point(290, 40);
      panel1.Name = "panel1";
      panel1.Size = new Size(880, 380);
      panel1.TabIndex = 2;
      panel1.Paint += panel1_Paint;

      chkBubble.Location = new Point(12, 430);
      chkBubble.Name = "chkBubble";
      chkBubble.Size = new Size(150, 22);
      chkBubble.TabIndex = 3;
      chkBubble.Text = "Пузырьковая";

      chkInsert.Location = new Point(12, 455);
      chkInsert.Name = "chkInsert";
      chkInsert.Size = new Size(150, 22);
      chkInsert.TabIndex = 4;
      chkInsert.Text = "Вставками";

      chkShaker.Location = new Point(12, 480);
      chkShaker.Name = "chkShaker";
      chkShaker.Size = new Size(150, 22);
      chkShaker.TabIndex = 5;
      chkShaker.Text = "Шейкерная";

      chkQuick.Location = new Point(12, 505);
      chkQuick.Name = "chkQuick";
      chkQuick.Size = new Size(150, 22);
      chkQuick.TabIndex = 6;
      chkQuick.Text = "Быстрая";

      chkBogo.Location = new Point(12, 530);
      chkBogo.Name = "chkBogo";
      chkBogo.Size = new Size(150, 22);
      chkBogo.TabIndex = 7;
      chkBogo.Text = "BOGO";

      gbDirection.Controls.Add(rbAsc);
      gbDirection.Controls.Add(rbDesc);
      gbDirection.Location = new Point(180, 430);
      gbDirection.Name = "gbDirection";
      gbDirection.Size = new Size(160, 80);
      gbDirection.TabIndex = 8;
      gbDirection.TabStop = false;
      gbDirection.Text = "Направление";

      rbAsc.Checked = true;
      rbAsc.Location = new Point(10, 20);
      rbAsc.Name = "rbAsc";
      rbAsc.Size = new Size(140, 20);
      rbAsc.TabIndex = 0;
      rbAsc.TabStop = true;
      rbAsc.Text = "По возрастанию";

      rbDesc.Location = new Point(10, 45);
      rbDesc.Name = "rbDesc";
      rbDesc.Size = new Size(140, 20);
      rbDesc.TabIndex = 1;
      rbDesc.Text = "По убыванию";

      numCount.Location = new Point(445, 433);
      numCount.Maximum = new decimal(new int[] { 200, 0, 0, 0 });
      numCount.Minimum = new decimal(new int[] { 2, 0, 0, 0 });
      numCount.Name = "numCount";
      numCount.Size = new Size(80, 27);
      numCount.TabIndex = 10;
      numCount.Value = new decimal(new int[] { 20, 0, 0, 0 });

      lblCountCaption.Location = new Point(360, 435);
      lblCountCaption.Name = "lblCountCaption";
      lblCountCaption.Size = new Size(80, 20);
      lblCountCaption.TabIndex = 9;
      lblCountCaption.Text = "Элементов:";

      lblBubbleCaption.Location = new Point(360, 470);
      lblBubbleCaption.Name = "lblBubbleCaption";
      lblBubbleCaption.Size = new Size(70, 20);
      lblBubbleCaption.TabIndex = 11;
      lblBubbleCaption.Text = "Bubble:";

      lblInsertCaption.Location = new Point(360, 495);
      lblInsertCaption.Name = "lblInsertCaption";
      lblInsertCaption.Size = new Size(70, 20);
      lblInsertCaption.TabIndex = 13;
      lblInsertCaption.Text = "Insert:";

      lblShakerCaption.Location = new Point(360, 520);
      lblShakerCaption.Name = "lblShakerCaption";
      lblShakerCaption.Size = new Size(70, 20);
      lblShakerCaption.TabIndex = 15;
      lblShakerCaption.Text = "Shaker:";

      lblQuickCaption.Location = new Point(570, 470);
      lblQuickCaption.Name = "lblQuickCaption";
      lblQuickCaption.Size = new Size(70, 20);
      lblQuickCaption.TabIndex = 17;
      lblQuickCaption.Text = "Quick:";

      lblBogoCaption.Location = new Point(570, 495);
      lblBogoCaption.Name = "lblBogoCaption";
      lblBogoCaption.Size = new Size(70, 20);
      lblBogoCaption.TabIndex = 19;
      lblBogoCaption.Text = "Bogo:";

      lblBubbleTime.Location = new Point(430, 470);
      lblBubbleTime.Name = "lblBubbleTime";
      lblBubbleTime.Size = new Size(120, 20);
      lblBubbleTime.TabIndex = 12;
      lblBubbleTime.Text = "—";

      lblInsertTime.Location = new Point(430, 495);
      lblInsertTime.Name = "lblInsertTime";
      lblInsertTime.Size = new Size(120, 20);
      lblInsertTime.TabIndex = 14;
      lblInsertTime.Text = "—";

      lblShakerTime.Location = new Point(430, 520);
      lblShakerTime.Name = "lblShakerTime";
      lblShakerTime.Size = new Size(120, 20);
      lblShakerTime.TabIndex = 16;
      lblShakerTime.Text = "—";

      lblQuickTime.Location = new Point(640, 470);
      lblQuickTime.Name = "lblQuickTime";
      lblQuickTime.Size = new Size(120, 20);
      lblQuickTime.TabIndex = 18;
      lblQuickTime.Text = "—";

      lblBogoTime.Location = new Point(640, 495);
      lblBogoTime.Name = "lblBogoTime";
      lblBogoTime.Size = new Size(120, 20);
      lblBogoTime.TabIndex = 20;
      lblBogoTime.Text = "—";

      ClientSize = new Size(1184, 561);
      Controls.Add(menuStrip1);
      Controls.Add(dataGridView1);
      Controls.Add(panel1);
      Controls.Add(chkBubble);
      Controls.Add(chkInsert);
      Controls.Add(chkShaker);
      Controls.Add(chkQuick);
      Controls.Add(chkBogo);
      Controls.Add(gbDirection);
      Controls.Add(lblCountCaption);
      Controls.Add(numCount);
      Controls.Add(lblBubbleCaption);
      Controls.Add(lblBubbleTime);
      Controls.Add(lblInsertCaption);
      Controls.Add(lblInsertTime);
      Controls.Add(lblShakerCaption);
      Controls.Add(lblShakerTime);
      Controls.Add(lblQuickCaption);
      Controls.Add(lblQuickTime);
      Controls.Add(lblBogoCaption);
      Controls.Add(lblBogoTime);
      MainMenuStrip = menuStrip1;
      Name = "Form1";
      StartPosition = FormStartPosition.CenterScreen;
      Text = "Лабораторная №2 — Олимпиадные сортировки";
      Load += Form1_Load;
      menuStrip1.ResumeLayout(false);
      menuStrip1.PerformLayout();
      ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
      gbDirection.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)numCount).EndInit();
      ResumeLayout(false);
      PerformLayout();
    }

    private System.Windows.Forms.MenuStrip menuStrip1;
    private System.Windows.Forms.ToolStripMenuItem miDataMenu;
    private System.Windows.Forms.ToolStripMenuItem miGenerate;
    private System.Windows.Forms.ToolStripMenuItem miLoadExcel;
    private System.Windows.Forms.ToolStripMenuItem miLoadGoogle;
    private System.Windows.Forms.ToolStripMenuItem miActionsMenu;
    private System.Windows.Forms.ToolStripMenuItem miCalculate;
    private System.Windows.Forms.ToolStripMenuItem miClear;
    private System.Windows.Forms.ToolStripMenuItem miExit;
    private System.Windows.Forms.DataGridView dataGridView1;
    private System.Windows.Forms.Panel panel1;
    private System.Windows.Forms.CheckBox chkBubble;
    private System.Windows.Forms.CheckBox chkInsert;
    private System.Windows.Forms.CheckBox chkShaker;
    private System.Windows.Forms.CheckBox chkQuick;
    private System.Windows.Forms.CheckBox chkBogo;
    private System.Windows.Forms.GroupBox gbDirection;
    private System.Windows.Forms.RadioButton rbAsc;
    private System.Windows.Forms.RadioButton rbDesc;
    private System.Windows.Forms.NumericUpDown numCount;
    private System.Windows.Forms.Label lblCountCaption;
    private System.Windows.Forms.Label lblBubbleCaption;
    private System.Windows.Forms.Label lblInsertCaption;
    private System.Windows.Forms.Label lblShakerCaption;
    private System.Windows.Forms.Label lblQuickCaption;
    private System.Windows.Forms.Label lblBogoCaption;
    private System.Windows.Forms.Label lblBubbleTime;
    private System.Windows.Forms.Label lblInsertTime;
    private System.Windows.Forms.Label lblShakerTime;
    private System.Windows.Forms.Label lblQuickTime;
    private System.Windows.Forms.Label lblBogoTime;
  }
}