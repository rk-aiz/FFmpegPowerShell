using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

/// <summary>
/// Provides static methods for writing log messages that only appear when a debugger is attached. <br />
/// デバッガーがアタッチされている場合にのみ表示されるログメッセージを書き込むための静的メソッドを提供します。 <br />
/// This is useful for developer-centric logging that should not appear in production environments
/// or when running the application without debugging (e.g., via Ctrl+F5). <br />
/// これは、本番環境やデバッグなしでの実行時（例: Ctrl+F5）には表示すべきでない、開発者向けのログ記録に役立ちます。
/// </summary>
public static class DevLog
{
    /// <summary>
    /// Gets or sets a value indicating whether the DevLog is enabled. Defaults to true. <br />
    /// DevLogが有効かどうかを示す値を取得または設定します。既定値は true です。 <br />
    /// Set to false to globally disable DevLog output even when a debugger is attached. <br />
    /// デバッガーがアタッチされている場合でもDevLogの出力を全体的に無効にするには、false に設定します。
    /// </summary>
    public static bool IsEnabled { get; set; }

    private static bool IsDebugDetected
    {
        get
        {
            if (System.Diagnostics.Debugger.IsAttached) return true;

            bool ret = false;
            try
            {
                var runspace = Runspace.DefaultRunspace;
                ret = runspace.Debugger.IsActive;
            }
            catch
            {
            }
            return ret;
        }
    }

    /// <summary>
    /// Static constructor to initialize properties for .NET Framework 4.x compatibility.
    /// .NET Framework 4.x との互換性のためにプロパティを初期化する静的コンストラクター。
    /// </summary>
    static DevLog()
    {
        IsEnabled = true;
    }

    /// <summary>
    /// Writes the specified string value, followed by the current line terminator, to the standard output stream only if a debugger is attached. <br />
    /// デバッガーがアタッチされている場合にのみ、指定された文字列値と、それに続く現在の行終端記号を標準出力ストリームに書き込みます。
    /// </summary>
    /// <param name="message">The value to write. / 書き込む値。</param>
    public static void WriteLine(string message)
    {
        if (!IsEnabled || !IsDebugDetected) return;

        Console.ForegroundColor = ConsoleColor.DarkGray;
        try
        {
            Console.WriteLine("[DEV] " + message);
        }
        finally
        {
            Console.ResetColor();
        }
    }

    /// <summary>
    /// Writes the text representation of the specified object, followed by the current line terminator, to the standard output stream only if a debugger is attached. <br />
    /// デバッガーがアタッチされている場合にのみ、指定されたオブジェクトのテキスト表現と、それに続く現在の行終端記号を標準出力ストリームに書き込みます。
    /// </summary>
    /// <param name="value">The value to write. / 書き込む値。</param>
    public static void WriteLine(object value)
    {
        if (!IsEnabled || !IsDebugDetected) return;

        WriteLine(value != null ? value.ToString() : null);
    }

    /// <summary>
    /// Writes the specified formatted string and arguments, followed by the current line terminator, to the standard output stream only if a debugger is attached. <br />
    /// デバッガーがアタッチされている場合にのみ、指定された書式指定文字列と引数、およびそれに続く現在の行終端記号を標準出力ストリームに書き込みます。
    /// </summary>
    /// <param name="format">A composite format string. / 複合書式指定文字列。</param>
    /// <param name="args">An array of objects to write using format. / 書式設定して書き込むオブジェクトの配列。</param>
    public static void WriteLine(string format, params object[] args)
    {
        if (!IsEnabled || !IsDebugDetected) return;

        // We check again to avoid the cost of string.Format if not attached.
        Console.ForegroundColor = ConsoleColor.DarkGray;
        try
        {
            Console.WriteLine("[DEV] " + string.Format(format, args));
        }
        finally
        {
            Console.ResetColor();
        }
    }

    /// <summary>
    /// Writes a message with detailed caller information (file, method, line number) to the standard output stream only if a debugger is attached. <br />
    /// デバッガーがアタッチされている場合にのみ、詳細な呼び出し元情報（ファイル、メソッド、行番号）を含むメッセージを標準出力ストリームに書き込みます。
    /// </summary>
    /// <param name="message">The message to write. / 書き込むメッセージ。</param>
    /// <param name="callerName">The name of the calling method (automatically populated). / 呼び出し元のメソッド名（自動的に設定されます）。</param>
    /// <param name="callerFile">The path of the source file (automatically populated). / ソースファイルのパス（自動的に設定されます）。</param>
    /// <param name="callerLine">The line number in the source file (automatically populated). / ソースファイル内の行番号（自動的に設定されます）。</param>
    public static void WriteWithCallerInfo(
        string message,
        [CallerMemberName] string callerName = "",
        [CallerFilePath] string callerFile = "",
        [CallerLineNumber] int callerLine = 0)
    {
        if (!IsEnabled || !IsDebugDetected) return;

        Console.ForegroundColor = ConsoleColor.Gray; // A slightly brighter color for more detail
        try
        {
            string fileName = System.IO.Path.GetFileName(callerFile);
            Console.WriteLine("[DEV] " + message);
            Console.WriteLine(string.Format("      └─ at {0} in {1}:{2}", callerName, fileName, callerLine));
        }
        finally
        {
            Console.ResetColor();
        }
    }
}
