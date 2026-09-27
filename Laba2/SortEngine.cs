using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laba2 {
  public static class SortEngine {
    private const int SmallDelayMs = 1;
    private const int BogoDelayMs = 20;
    private const int BogoMaxTries = 100000;

    private static bool NeedSwap(int firstValue, int secondValue, bool ascending) {
      bool result;
      if (ascending)
        result = firstValue > secondValue;
      else
        result = firstValue < secondValue;
      return result;
    }

    public static async Task BubbleSortAsync(int[] numbers, bool ascending, IProgress<int[]> progress) {
      int length = numbers.Length;
      for (int outer = 0; outer < length - 1; ++outer) {
        for (int inner = 0; inner < length - 1 - outer; ++inner) {
          if (NeedSwap(numbers[inner], numbers[inner + 1], ascending)) {
            int temp = numbers[inner];
            numbers[inner] = numbers[inner + 1];
            numbers[inner + 1] = temp;

            if (progress != null)
              progress.Report((int[])numbers.Clone());
            await Task.Delay(SmallDelayMs);
          }
        }
      }
    }

    public static async Task InsertSortAsync(int[] numbers, bool ascending, IProgress<int[]> progress) {
      int length = numbers.Length;
      for (int currentIndex = 1; currentIndex < length; ++currentIndex) {
        int keyValue = numbers[currentIndex];
        int scanIndex = currentIndex - 1;
        while (scanIndex >= 0 && NeedSwap(numbers[scanIndex], keyValue, ascending)) {
          numbers[scanIndex + 1] = numbers[scanIndex];
          --scanIndex;
        }
        numbers[scanIndex + 1] = keyValue;

        if (progress != null)
          progress.Report((int[])numbers.Clone());
        await Task.Delay(SmallDelayMs);
      }
    }

    public static async Task ShakerSortAsync(int[] numbers, bool ascending, IProgress<int[]> progress) {
      int leftBorder = 0;
      int rightBorder = numbers.Length - 1;
      while (leftBorder < rightBorder) {
        for (int forwardIndex = leftBorder; forwardIndex < rightBorder; ++forwardIndex) {
          if (NeedSwap(numbers[forwardIndex], numbers[forwardIndex + 1], ascending)) {
            int temp = numbers[forwardIndex];
            numbers[forwardIndex] = numbers[forwardIndex + 1];
            numbers[forwardIndex + 1] = temp;
          }
        }
        --rightBorder;

        for (int backwardIndex = rightBorder; backwardIndex > leftBorder; --backwardIndex) {
          if (NeedSwap(numbers[backwardIndex - 1], numbers[backwardIndex], ascending)) {
            int temp = numbers[backwardIndex - 1];
            numbers[backwardIndex - 1] = numbers[backwardIndex];
            numbers[backwardIndex] = temp;
          }
        }
        ++leftBorder;

        if (progress != null)
          progress.Report((int[])numbers.Clone());
        await Task.Delay(SmallDelayMs);
      }
    }

    public static async Task QuickSortAsync(int[] numbers, bool ascending, IProgress<int[]> progress) {
      await QuickSortRecursive(numbers, 0, numbers.Length - 1, ascending, progress);
      if (progress != null)
        progress.Report((int[])numbers.Clone());
    }

    private static async Task QuickSortRecursive(int[] numbers, int lowIndex, int highIndex, bool ascending, IProgress<int[]> progress) {
      if (lowIndex >= highIndex)
        return;

      int pivotIndex = Partition(numbers, lowIndex, highIndex, ascending);

      await QuickSortRecursive(numbers, lowIndex, pivotIndex - 1, ascending, progress);
      await QuickSortRecursive(numbers, pivotIndex + 1, highIndex, ascending, progress);

      if (progress != null)
        progress.Report((int[])numbers.Clone());
      await Task.Delay(SmallDelayMs);
    }

    private static int Partition(int[] numbers, int lowIndex, int highIndex, bool ascending) {
      int pivotValue = numbers[highIndex];
      int smallerIndex = lowIndex - 1;

      for (int scanIndex = lowIndex; scanIndex < highIndex; ++scanIndex) {
        bool shouldMoveLeft;
        if (ascending)
          shouldMoveLeft = numbers[scanIndex] <= pivotValue;
        else
          shouldMoveLeft = numbers[scanIndex] >= pivotValue;

        if (shouldMoveLeft) {
          ++smallerIndex;
          int temp = numbers[smallerIndex];
          numbers[smallerIndex] = numbers[scanIndex];
          numbers[scanIndex] = temp;
        }
      }

      int finalTemp = numbers[smallerIndex + 1];
      numbers[smallerIndex + 1] = numbers[highIndex];
      numbers[highIndex] = finalTemp;

      return smallerIndex + 1;
    }

    public static async Task BogoSortAsync(int[] numbers, bool ascending, IProgress<int[]> progress) {
      Random randomizer = new Random();
      int attemptCounter = 0;

      while (!IsSorted(numbers, ascending)) {
        for (int swapIndex = numbers.Length - 1; swapIndex > 0; --swapIndex) {
          int partnerIndex = randomizer.Next(swapIndex + 1);
          int temp = numbers[swapIndex];
          numbers[swapIndex] = numbers[partnerIndex];
          numbers[partnerIndex] = temp;
        }

        if (progress != null)
          progress.Report((int[])numbers.Clone());
        await Task.Delay(BogoDelayMs);

        ++attemptCounter;
        if (attemptCounter > BogoMaxTries) {
          MessageBox.Show(
              "BOGO-сортировка прервана: слишком много попыток.",
              "Ограничение");
          break;
        }
      }
    }

    private static bool IsSorted(int[] numbers, bool ascending) {
      int length = numbers.Length;
      for (int checkIndex = 0; checkIndex < length - 1; ++checkIndex) {
        if (NeedSwap(numbers[checkIndex], numbers[checkIndex + 1], ascending))
          return false;
      }
      return true;
    }
  }
}