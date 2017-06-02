using System;
using System.Runtime.InteropServices;

namespace DLLHandler.Framework
{
    /// <summary>
    /// LoadLibraryEx関数の第3引数に渡すフラグ定数
    /// </summary>
    public enum DllEntryPointFlag : uint
    {
        DONT_RESOLVE_DLL_REFERENCES = 0x00000001,
        LOAD_IGNORE_CODE_AUTHZ_LEVEL = 0x00000010,
        LOAD_LIBRARY_AS_DATAFILE = 0x00000002,
        LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = 0x00000040,
        LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x00000020,
        LOAD_LIBRARY_SEARCH_APPLICATION_DIR = 0x00000200,
        LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = 0x00001000,
        LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR = 0x00000100,
        LOAD_LIBRARY_SEARCH_SYSTEM32 = 0x00000800,
        LOAD_LIBRARY_SEARCH_USER_DIRS = 0x00000400,
        LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008
    }

    /// <summary>
    /// DLLを動的に読み込むためのクラス
    /// </summary>
    public class DllHandler
    {

        // 型の対応は http://www.atmarkit.co.jp/fdotnet/dotnettips/024w32api/w32api.html を参照

        [DllImport("kernel32.dll", EntryPoint = "LoadLibrary", BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr LoadLibrary([MarshalAs(UnmanagedType.LPStr)] System.String lpLibFileName);

        [DllImport("kernel32.dll", EntryPoint = "LoadLibraryEx", BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr LoadLibraryEx([MarshalAs(UnmanagedType.LPStr)] System.String lpLibFileName, IntPtr hFile, uint dwFlags);

        [DllImport("kernel32.dll", EntryPoint = "GetProcAddress", BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr GetProcAddress(IntPtr hModule, [MarshalAs(UnmanagedType.LPStr)] string lpProcName);

        [DllImport("kernel32.dll", EntryPoint = "FreeLibrary")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FreeLibrary(IntPtr hModule);


        private IntPtr moduleHandle = IntPtr.Zero;
        private bool isAvailable = false;

        public IntPtr ModuleHandle
        {
            get
            {
                return this.moduleHandle;
            }
        }

        public bool IsAvailable
        {
            get
            {
                return this.isAvailable;
            }
        }

        /// <summary>
        /// LoadLibrary関数を使ってDLLを読み込む
        /// </summary>
        /// <param name="lpFileName"></param>
        /// <seealso cref="https://msdn.microsoft.com/ja-jp/library/cc429241.aspx"/>
        public DllHandler(string lpFileName)
        {
            if (!System.IO.File.Exists(lpFileName))
            {
                throw new System.IO.FileNotFoundException(lpFileName + " is not found");
            }

            moduleHandle = LoadLibrary(lpFileName);
            this.isAvailable = true;
        }

        /// <summary>
        /// LoadLibraryEx関数を使ってDLLを読み込む
        /// </summary>
        /// <param name="lpFileName"></param>
        /// <param name="dwFlags"></param>
        /// <seealso cref="https://msdn.microsoft.com/ja-jp/library/cc429243.aspx"/>
        public DllHandler(string lpFileName, DllEntryPointFlag dwFlags)
        {
            if (!System.IO.File.Exists(lpFileName))
            {
                throw new System.IO.FileNotFoundException(lpFileName + "is not found");
            }

            // 第2引数は
            moduleHandle = LoadLibraryEx(lpFileName, IntPtr.Zero, (uint)dwFlags);
            this.isAvailable = true;
        }

        /// <summary>
        /// 呼び出す関数のデリゲートを作成する
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="method">呼び出す関数の名前</param>
        /// <returns>デリゲート</returns>
        /// <example>
        /// delegate int Add(int a, int b); // 呼び出す関数
        /// DllHandler sampleDll = new DllHandler("sample.dll");
        /// Add add = sampleDll.GetProcDelegate<Add>("Add");
        /// add(3,5); // = 8
        /// </example>
        /// <seealso cref="https://msdn.microsoft.com/ja-jp/library/cc429133.aspx"/>
        public T GetProcDelegate<T>(string method) where T : class
        {
            if (!this.IsAvailable)
            {
                throw new System.IO.FileNotFoundException();
            }

            if (moduleHandle == IntPtr.Zero)
            {
                throw new DllNotFoundException();
            }

            IntPtr methodHandle = GetProcAddress(moduleHandle, method);
            T r = Marshal.GetDelegateForFunctionPointer(methodHandle, typeof(T)) as T;
            return r;
        }

        /// <summary>
        /// 読み込んだ動的ライブラリを解放する
        /// </summary>
        /// <seealso cref="https://msdn.microsoft.com/ja-jp/library/cc429103.aspx"/>
        public void Dispose()
        {
            if (this.IsAvailable && this.ModuleHandle != IntPtr.Zero)
            {
                FreeLibrary(moduleHandle);
            }
        }

    }
}
