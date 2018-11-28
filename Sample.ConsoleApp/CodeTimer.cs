using System;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;

public static class CodeTimer
{
    public static void Initialize()
    {
        Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.RealTime;
        Thread.CurrentThread.Priority = ThreadPriority.Highest;
        Time(string.Empty, 0, () => { });
    }

    public static void Time(string name, int iteration, Action action)
    {
        if (String.IsNullOrEmpty(name)) return;

        // 1.
        var currentForeColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(name);

        // 2.
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
        var gcCounts = new int[GC.MaxGeneration + 1];
        for (var i = 0; i <= GC.MaxGeneration; i++)
        {
            gcCounts[i] = GC.CollectionCount(i);
        }

        // 3.
        var watch = new Stopwatch();
        watch.Start();
        var cycleCount = GetCycleCount();
        for (var i = 0; i < iteration; i++) action();
        var cpuCycles = GetCycleCount() - cycleCount;
        watch.Stop();

        // 4.
        Console.ForegroundColor = currentForeColor;
        Console.WriteLine("\tTime Elapsed:\t" + watch.ElapsedMilliseconds.ToString("N0") + "ms");
        Console.WriteLine("\tCPU Cycles:\t" + cpuCycles.ToString("N0"));

        // 5.
        for (var i = 0; i <= GC.MaxGeneration; i++)
        {
            var count = GC.CollectionCount(i) - gcCounts[i];
            Console.WriteLine("\tGen " + i + ": \t\t" + count);
        }

        Console.WriteLine();
    }

    static ulong GetCycleCount()
    {
        var cycleCount = 0UL;
        QueryThreadCycleTime(GetCurrentThread(), ref cycleCount);
        return cycleCount;
    }

    /// <summary>
    /// Retrieves the cycle time for the specified thread.
    /// </summary>
    /// <param name="threadHandle">A handle to the thread. The handle must have the PROCESS_QUERY_INFORMATION or PROCESS_QUERY_LIMITED_INFORMATION access right.</param>
    /// <param name="cycleTime">The number of CPU clock cycles used by the thread. This value includes cycles spent in both user mode and kernel mode.</param>
    /// <remarks>
    /// Minimum supported client: Windows Vista
    /// Minimum supported server: Windows Server 2008
    /// </remarks>
    /// <returns></returns>
    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    static extern bool QueryThreadCycleTime(IntPtr threadHandle, ref ulong cycleTime);

    /// <summary>
    /// Retrieves a pseudo handle for the calling thread.
    /// </summary>
    /// <returns></returns>
    [DllImport("kernel32.dll")]
    static extern IntPtr GetCurrentThread();
}