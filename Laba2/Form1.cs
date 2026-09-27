using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laba2 {
  public partial class Form1 : Form {
    private const int MinValue = -100;
    private const int MaxValue = 101;
    private const int MinElements = 2;
    private const int GridColumnWidth = 220;
    private const int BarMinWidth = 2;
    private const int BarGap = 2;
    private const int BarVerticalPadding = 10;
    private const int StateOffset = 3;
    private const int ExcelSheetIndex = 1;
    private const int ExcelFirstRow = 1;
    private const int ExcelFirstCol = 1;
    private const int BarMinHeight = 1;

    private List<int> dataSet;
    private readonly Random randomizer;

    private readonly Dictionary<string, int[]> sortedStates;
    private readonly Dictionary<string, Color> palette;

    public Form1() {
      InitializeComponent();
      SetupGrid();

      dataSet = new List<int>();
      randomizer = new Random();
      sortedStates = new Dictionary<string, int[]>();
      palette = new Dictionary<string, Color>();
      palette["Bubble"] = Color.OrangeRed;
      palette["Insert"] = Color.SeaGreen;
      palette["Shaker"] = Color.RoyalBlue;
      palette["Quick"] = Color.Purple;
      palette["Bogo"] = Color.Brown;
    }

    private void SetupGrid() {
      dataGridView1.Columns.Clear();
      dataGridView1.Columns.Add("ValueColumn", "Значение");
      dataGridView1.Columns[0].Width = GridColumnWidth;
    }

    private void DataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e) {
      string rawValue;
      if (e.FormattedValue == null)
        rawValue = string.Empty;
      else
        rawValue = e.FormattedValue.ToString();

      if (string.IsNullOrWhiteSpace(rawValue)) {
        dataGridView1.Rows[e.RowIndex].ErrorText = string.Empty;
        return;
      }

      bool isParsed;
      isParsed = int.TryParse(rawValue, out _);

      if (!isParsed) {
        e.Cancel = true;
        dataGridView1.Rows[e.RowIndex].ErrorText = "Введите целое число!";
        MessageBox.Show(
            "Некорректное значение: «" + rawValue + "». Ожидается целое число.",
            "Ошибка ввода",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
      } else {
        dataGridView1.Rows[e.RowIndex].ErrorText = string.Empty;
      }
    }

    private void miGenerate_Click(object sender, EventArgs e) {
      int elementCount;
      elementCount = (int)numCount.Value;

      dataSet.Clear();
      for (int counter = 0; counter < elementCount; ++counter)
        dataSet.Add(randomizer.Next(MinValue, MaxValue));

      FillGrid();
      panel1.Invalidate();
    }

    private void FillGrid() {
      dataGridView1.Rows.Clear();
      foreach (int value in dataSet)
        dataGridView1.Rows.Add(value);
    }

    private bool ReadGrid() {
      dataSet.Clear();
      foreach (DataGridViewRow currentRow in dataGridView1.Rows) {
        if (currentRow.IsNewRow)
          continue;

        object cellValue;
        cellValue = currentRow.Cells[0].Value;

        if (cellValue == null)
          continue;

        string cellText;
        cellText = cellValue.ToString();

        if (string.IsNullOrWhiteSpace(cellText))
          continue;

        int parsedValue;
        bool isParsed;
        isParsed = int.TryParse(cellText, out parsedValue);

        if (!isParsed) {
          MessageBox.Show(
              "В таблице есть некорректные значения.",
              "Ошибка",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
          return false;
        }
        dataSet.Add(parsedValue);
      }

      if (dataSet.Count < MinElements) {
        MessageBox.Show(
            "Нужно минимум " + MinElements + " элемента.",
            "Ошибка",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return false;
      }
      return true;
    }

    private void miLoadExcel_Click(object sender, EventArgs e) {
      OpenFileDialog fileDialog;
      fileDialog = new OpenFileDialog();
      fileDialog.Filter = "Excel|*.xlsx;*.xls";

      if (fileDialog.ShowDialog() != DialogResult.OK)
        return;

      try {
        Microsoft.Office.Interop.Excel.Application excelApp;
        Microsoft.Office.Interop.Excel.Workbook workbook;
        Microsoft.Office.Interop.Excel.Worksheet worksheet;
        Microsoft.Office.Interop.Excel.Range usedRange;

        excelApp = new Microsoft.Office.Interop.Excel.Application();
        workbook = excelApp.Workbooks.Open(fileDialog.FileName);
        worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[ExcelSheetIndex];
        usedRange = worksheet.UsedRange;

        dataSet.Clear();
        for (int rowIndex = ExcelFirstRow; rowIndex <= usedRange.Rows.Count; ++rowIndex) {
          Microsoft.Office.Interop.Excel.Range cellRange;
          cellRange = usedRange.Cells[rowIndex, ExcelFirstCol] as Microsoft.Office.Interop.Excel.Range;

          object cellValue;
          if (cellRange == null)
            cellValue = null;
          else
            cellValue = cellRange.Value2;

          if (cellValue != null) {
            int parsedValue;
            if (int.TryParse(cellValue.ToString(), out parsedValue))
              dataSet.Add(parsedValue);
          }
        }

        workbook.Close(false);
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

        FillGrid();
        MessageBox.Show("Загружено " + dataSet.Count + " значений.", "Успех");
      } catch (Exception ex) {
        MessageBox.Show(
            "Ошибка загрузки Excel: " + ex.Message,
            "Ошибка",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
      }
    }

    private void miLoadGoogle_Click(object sender, EventArgs e) {
      string url;
      url = Microsoft.VisualBasic.Interaction.InputBox(
          "Введите ссылку на CSV-экспорт Google таблицы:",
          "Google Table",
          string.Empty);

      if (string.IsNullOrWhiteSpace(url))
        return;

      try {
        using (System.Net.WebClient webClient = new System.Net.WebClient()) {
          string rawText;
          rawText = webClient.DownloadString(url);

          char[] separators;
          separators = new char[] { '\n', '\r', ',', ';', '\t' };

          string[] rawTokens;
          rawTokens = rawText.Split(separators, StringSplitOptions.RemoveEmptyEntries);

          dataSet = new List<int>();

          foreach (string rawToken in rawTokens) {
            string cleanToken;
            cleanToken = rawToken.Trim();
            cleanToken = cleanToken.Trim('"');
            cleanToken = cleanToken.Replace("\u00A0", "");
            cleanToken = cleanToken.Replace(" ", "");

            if (cleanToken.Length == 0)
              continue;

            int parsedValue;
            bool isParsed;
            isParsed = int.TryParse(cleanToken, out parsedValue);

            if (isParsed)
              dataSet.Add(parsedValue);
          }

          FillGrid();
          MessageBox.Show("Загружено " + dataSet.Count + " значений.", "Успех");
        }
      } catch (Exception ex) {
        MessageBox.Show(
            "Ошибка загрузки: " + ex.Message,
            "Ошибка",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
      }
    }

    private void miClear_Click(object sender, EventArgs e) {
      dataSet.Clear();
      sortedStates.Clear();
      dataGridView1.Rows.Clear();
      panel1.Invalidate();
      lblBubbleTime.Text = "—";
      lblInsertTime.Text = "—";
      lblShakerTime.Text = "—";
      lblQuickTime.Text = "—";
      lblBogoTime.Text = "—";
    }

    private void miExit_Click(object sender, EventArgs e) {
      Close();
    }

    private async void miCalculate_Click(object sender, EventArgs e) {
      if (!ReadGrid())
        return;

      List<Tuple<string, Func<int[], bool, IProgress<int[]>, Task>>> selectedAlgorithms;
      selectedAlgorithms = new List<Tuple<string, Func<int[], bool, IProgress<int[]>, Task>>>();

      if (chkBubble.Checked)
        selectedAlgorithms.Add(Tuple.Create("Bubble", (Func<int[], bool, IProgress<int[]>, Task>)SortEngine.BubbleSortAsync));
      if (chkInsert.Checked)
        selectedAlgorithms.Add(Tuple.Create("Insert", (Func<int[], bool, IProgress<int[]>, Task>)SortEngine.InsertSortAsync));
      if (chkShaker.Checked)
        selectedAlgorithms.Add(Tuple.Create("Shaker", (Func<int[], bool, IProgress<int[]>, Task>)SortEngine.ShakerSortAsync));
      if (chkQuick.Checked)
        selectedAlgorithms.Add(Tuple.Create("Quick", (Func<int[], bool, IProgress<int[]>, Task>)SortEngine.QuickSortAsync));
      if (chkBogo.Checked)
        selectedAlgorithms.Add(Tuple.Create("Bogo", (Func<int[], bool, IProgress<int[]>, Task>)SortEngine.BogoSortAsync));

      if (selectedAlgorithms.Count == 0) {
        MessageBox.Show("Выберите хотя бы один алгоритм.", "Внимание");
        return;
      }

      bool ascendingOrder;
      ascendingOrder = rbAsc.Checked;

      sortedStates.Clear();
      panel1.Invalidate();

      Dictionary<string, long> algorithmTimings;
      List<Task> runningTasks;

      algorithmTimings = new Dictionary<string, long>();
      runningTasks = new List<Task>();

      foreach (Tuple<string, Func<int[], bool, IProgress<int[]>, Task>> selectedItem in selectedAlgorithms) {
        string algorithmName;
        algorithmName = selectedItem.Item1;

        int[] inputCopy;
        inputCopy = dataSet.ToArray();

        Progress<int[]> progressHandler;
        progressHandler = new Progress<int[]>(sortedSnapshot => {
          sortedStates[algorithmName] = sortedSnapshot;
          panel1.Invalidate();
        });

        string capturedName;
        capturedName = algorithmName;

        Func<int[], bool, IProgress<int[]>, Task> algorithmDelegate;
        algorithmDelegate = selectedItem.Item2;

        runningTasks.Add(Task.Run(async () => {
          Stopwatch stopwatch;
          stopwatch = Stopwatch.StartNew();

          await algorithmDelegate(inputCopy, ascendingOrder, progressHandler);

          stopwatch.Stop();
          lock (algorithmTimings) {
            algorithmTimings[capturedName] = stopwatch.ElapsedMilliseconds;
          }
        }));
      }

      await Task.WhenAll(runningTasks);

      long elapsedTime;

      if (algorithmTimings.TryGetValue("Bubble", out elapsedTime))
        lblBubbleTime.Text = elapsedTime + " мс";
      if (algorithmTimings.TryGetValue("Insert", out elapsedTime))
        lblInsertTime.Text = elapsedTime + " мс";
      if (algorithmTimings.TryGetValue("Shaker", out elapsedTime))
        lblShakerTime.Text = elapsedTime + " мс";
      if (algorithmTimings.TryGetValue("Quick", out elapsedTime))
        lblQuickTime.Text = elapsedTime + " мс";
      if (algorithmTimings.TryGetValue("Bogo", out elapsedTime))
        lblBogoTime.Text = elapsedTime + " мс";

      KeyValuePair<string, long> fastestAlgorithm;
      fastestAlgorithm = algorithmTimings.OrderBy(pair => pair.Value).First();

      MessageBox.Show(
          "Самый быстрый алгоритм: " + fastestAlgorithm.Key + " (" + fastestAlgorithm.Value + " мс)",
          "Результат",
          MessageBoxButtons.OK,
          MessageBoxIcon.Information);
    }

    private void panel1_Paint(object sender, PaintEventArgs e) {
      Graphics graphics;
      graphics = e.Graphics;
      graphics.Clear(Color.White);

      Dictionary<string, int[]> statesToDraw;

      if (sortedStates.Count > 0) {
        statesToDraw = sortedStates;
      } else if (dataSet.Count > 0) {
        statesToDraw = new Dictionary<string, int[]>();
        statesToDraw["Source"] = dataSet.ToArray();
      } else {
        statesToDraw = new Dictionary<string, int[]>();
      }

      if (statesToDraw.Count == 0)
        return;

      int horizontalOffset;
      horizontalOffset = 0;

      foreach (KeyValuePair<string, int[]> currentState in statesToDraw) {
        Color barColor;
        if (palette.TryGetValue(currentState.Key, out barColor) == false)
          barColor = Color.Gray;

        DrawArray(graphics, currentState.Value, barColor, horizontalOffset);
        horizontalOffset += StateOffset;
      }
    }

    private void DrawArray(Graphics graphics, int[] numbers, Color barColor, int horizontalOffset) {
      if (numbers == null || numbers.Length == 0)
        return;

      int barWidth;
      barWidth = Math.Max(BarMinWidth, panel1.Width / numbers.Length);

      int maxAbsoluteValue;
      maxAbsoluteValue = 0;

      for (int scanIndex = 0; scanIndex < numbers.Length; ++scanIndex) {
        int currentAbsolute;
        currentAbsolute = Math.Abs(numbers[scanIndex]);
        if (currentAbsolute > maxAbsoluteValue)
          maxAbsoluteValue = currentAbsolute;
      }

      if (maxAbsoluteValue < 1)
        maxAbsoluteValue = 1;

      int centerLine;
      centerLine = panel1.Height / 2;

      int halfHeight;
      halfHeight = panel1.Height / 2 - BarVerticalPadding;

      using (SolidBrush barBrush = new SolidBrush(barColor)) {
        for (int barIndex = 0; barIndex < numbers.Length; ++barIndex) {
          int barHeight;
          barHeight = (int)((double)Math.Abs(numbers[barIndex]) / maxAbsoluteValue * halfHeight);

          int barX;
          barX = barIndex * barWidth + horizontalOffset;

          int barY;
          if (numbers[barIndex] >= 0)
            barY = centerLine - barHeight;
          else
            barY = centerLine;

          int drawWidth;
          drawWidth = Math.Max(BarMinHeight, barWidth - BarGap);

          int drawHeight;
          drawHeight = Math.Max(BarMinHeight, barHeight);

          graphics.FillRectangle(barBrush, barX, barY, drawWidth, drawHeight);
        }
      }
    }

    private void Form1_Load(object sender, EventArgs e) {
    }
  }
}