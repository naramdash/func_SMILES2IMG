using System;

namespace Smiles2Img.Tests;

internal static class TestAssert
{
    public static int PassedSuites { get; private set; }

    public static void ResetCounters() => PassedSuites = 0;

    public static void Check(bool condition, string message)
    {
        if (!condition)
            throw new Exception("Assertion failed: " + message);
    }

    public static void PassSuite(string suiteName)
    {
        PassedSuites++;
        Console.WriteLine($"  [PASS] {suiteName}");
    }

    public static void PrintCategoryHeader(string categoryTitle)
    {
        Console.WriteLine();
        Console.WriteLine("======================================================================");
        Console.WriteLine($"  {categoryTitle}");
        Console.WriteLine("======================================================================");
    }
}
