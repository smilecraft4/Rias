using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using CommandLine;
using ImageMagick;

namespace Rias
{
    internal class Program
    {
        private static int CollectedFolderCount = 0;
        private static int ApplyCoverProgress = 0;
        private static int RemoveCoverProgress = 0;

        class Options
        {
            [Option('f', "folder", Default = ".\\", HelpText = "Where to start applying cover to folders")]
            public required string Path { get; set; }

            [Option('d', "depth", Default = 1, HelpText = "Maximum search depth for nested folders")]
            public int? Depth { get; set; }

            [Option('o', "overwrite", Default = false, HelpText = "Regenerate cover for folder with existing covers")]
            public bool Overwrite { get; set; }

            [Option('v', "verbose", Default = false, HelpText = "Enable verbose output")]
            public bool Verbose { get; set; }

            // Hidden

            [Option("folder-with-subfolders", Default = false, Hidden = true)]
            public bool IncludeParentFolder { get; set; }
        }

        [Verb("apply", HelpText = "Create covers for folders containg pictures")]
        class ApplyOptions : Options
        {
            public enum SortType
            {
                nameAsc,
                nameDes,
                dateAsc,
                dateDes,
                random
            }

            [Option('s', "sort", Default = SortType.nameAsc, HelpText = "Specify the sort type: (nameAsc, nameDes, dateAsc, dateDes, random)")]
            public SortType Sort { get; set; }

            // Hidden options

            [Option("cover-visible", Default = false, Hidden = true)]
            public bool CoverVisible { get; set; }

            [Option("ini-visible", Default = false, Hidden = true)]
            public bool DesktopIniVisible { get; set; }

            [Option("ico-resolutions", Default = "256, 48, 32, 24, 16", Hidden = true)]
            public required string Resolutions { get; set; }

            [Option("cover-source-filter", Default = ".jpg, .jpeg, .png, .webp", Hidden = true)]
            public required string Filter { get; set; }
        }

        [Verb("remove", HelpText = "Remove covers of folders")]
        class RemoveOptions : Options
        {
            [Option("ico", Default = "icon.ico", HelpText = "Remove the cover file with this name")]
            public required string CoverFile { get; set; }

            // Hidden

            [Option("everyhing", Default = false, HelpText = "Remove also every destkop.ico and cover.ico (including system like for the desktop)", Hidden = true)]
            public bool Everything { get; set; }
        }

        private static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            Parser.Default.ParseArguments<ApplyOptions, RemoveOptions>(args)
                .WithParsed<ApplyOptions>(Apply)
                .WithParsed<RemoveOptions>(Remove);
        }

        private static void Apply(ApplyOptions applyOptions)
        {
            var collectedFolders = CollectFolder(applyOptions);

            Console.WriteLine($"Processing applying covers for {collectedFolders.Count} folder");

            var processingStopwatch = Stopwatch.StartNew();
            Parallel.ForEach(collectedFolders, (dir) => ApplyCover(applyOptions, dir));
            processingStopwatch.Stop();

            var progressString = ApplyCoverProgress.ToString().PadLeft(CollectedFolderCount.ToString().Length);
            Console.WriteLine($"\rProcessed covers for {progressString}/{collectedFolders.Count} folder    ");

            // Update filesystem
            SHChangeNotify(0x8000000, 0x0, IntPtr.Zero, IntPtr.Zero);
        }

        private static void Remove(RemoveOptions removeOptions)
        {
            var collectedFolders = CollectFolder(removeOptions);

            Console.WriteLine($"Processing removing covers for {collectedFolders.Count} folder");


            var processingStopwatch = Stopwatch.StartNew();
            // Parallel.ForEach(collectedFolders, (dir) => RemoveCover(removeOptions, dir));

            foreach (var folder in collectedFolders)
            {
                RemoveCover(removeOptions, folder);
            }

            processingStopwatch.Stop();

            var progressString = RemoveCoverProgress.ToString().PadLeft(CollectedFolderCount.ToString().Length);
            Console.WriteLine($"\rRemoved covers for {progressString}/{collectedFolders.Count} folder    ");

            // Update filesystem
            SHChangeNotify(0x8000000, 0x0, IntPtr.Zero, IntPtr.Zero);
        }

        private static List<DirectoryInfo> CollectFolder(Options options)
        {
            var collectingStopwatch = Stopwatch.StartNew();
            var collectedFolders = new List<DirectoryInfo>();
            var pathFolderInfo = new DirectoryInfo(options.Path);
            RecursiveCollectFolder(options, pathFolderInfo, collectedFolders, 0);
            collectingStopwatch.Stop();

            CollectedFolderCount = collectedFolders.Count;

            // TODO: Add verbose message here for timings
            return collectedFolders;
        }

        private static void ApplyCover(ApplyOptions applyOptions, DirectoryInfo directoryInfo)
        {
            var fileInfos = directoryInfo.EnumerateFiles();
            var coverFile = SortFiles(applyOptions, fileInfos);

            // TODO convert "icon.ico" to a constant and thing about backwards compatibility

            // TODO consider adding a flag to specify this folder as a rias cover applied to it
            bool overwrite = applyOptions.Overwrite || !fileInfos.Any(f => f.Name == "icon.ico");

            if (overwrite)
            {
                { // create cover.ico
                    var iconDst = Path.Combine(directoryInfo.FullName, "icon.ico");

                    var tmpIconDst = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.ico");
                    using MagickImage image = new MagickImage(coverFile.FullName);
                    var maxSize = Math.Max(image.Width, image.Height);

                    image.Extent(maxSize, maxSize, Gravity.Center, MagickColor.FromRgba(255, 255, 255, 0));
                    image.Settings.SetDefine(MagickFormat.Ico, "auto-resize", applyOptions.Resolutions);
                    image.Write(tmpIconDst);
                    File.Move(tmpIconDst, iconDst, true);
                    File.Delete(tmpIconDst);

                    var attributes = File.GetAttributes(iconDst);
                    attributes |= FileAttributes.System;
                    if (!applyOptions.CoverVisible)
                    {
                        attributes |= FileAttributes.Hidden;
                    }
                    File.SetAttributes(iconDst, attributes);
                }

                { // create dekstop.ini
                    var desktopIniPath = Path.Combine(directoryInfo.FullName, "desktop.ini");

                    var tmpDesktopIni = Path.GetTempFileName();
                    var fileStream = new FileStream(tmpDesktopIni, FileMode.Create, FileAccess.Write);
                    using (StreamWriter sw = new StreamWriter(fileStream))
                    {
                        // TODO add version using exe version
                        sw.WriteLine("[.ShellClassInfo]");
                        sw.WriteLine($"IconResource=.\\icon.ico,0");
                        sw.WriteLine($"IconFile=.\\icon.ico");
                        sw.WriteLine($"IconIndex=0");
                        sw.WriteLine($";rias 0.1.0");
                        sw.Close();
                    }

                    File.Move(tmpDesktopIni, desktopIniPath, true);
                    File.Delete(tmpDesktopIni);

                    var attributes = File.GetAttributes(desktopIniPath);
                    attributes |= FileAttributes.System;
                    if (!applyOptions.DesktopIniVisible)
                    {
                        attributes |= FileAttributes.Hidden;
                    }
                    File.SetAttributes(desktopIniPath, attributes);
                }

                // Set directory properties here
                File.SetAttributes(directoryInfo.FullName,
                File.GetAttributes(directoryInfo.FullName)
                | FileAttributes.ReadOnly
                | FileAttributes.System
                );

                Interlocked.Increment(ref ApplyCoverProgress);

                var progressString = ApplyCoverProgress.ToString().PadLeft(CollectedFolderCount.ToString().Length);
                var msg = $"\rApplied {progressString}/{CollectedFolderCount} covers";
                Console.Write($"{msg}\r");
            }
        }

        private static void RemoveCover(RemoveOptions removeOptions, DirectoryInfo directoryInfo)
        {
            var fileInfos = directoryInfo.EnumerateFiles();

            var desktopIniFiles = fileInfos.Where(f => f.Name == "desktop.ini");
            if (desktopIniFiles.Any())
            {
                var coverFiles = fileInfos.Where(f => f.Name == removeOptions.CoverFile);
                if (coverFiles.Any())
                {
                    // TODO add verbose text here
                    foreach (var coverFile in coverFiles) { coverFile.Delete(); }
                    desktopIniFiles.First().Delete();
                }
                else if (removeOptions.Everything)
                {
                    // TODO: read the content of the files and find the references icons
                    // if the icons is stored in this folder delete it

                    // TODO add verbose text here
                    desktopIniFiles.First().Delete();
                }
            }

            Interlocked.Increment(ref RemoveCoverProgress);

            var progressString = RemoveCoverProgress.ToString().PadLeft(CollectedFolderCount.ToString().Length);
            var msg = $"\rRemoved {progressString}/{CollectedFolderCount} covers";
            Console.Write($"{msg}\r");
        }

        // TODO: Consider changing the ref collectedFolder to a Thread safe alternative
        private static bool RecursiveCollectFolder(Options options, DirectoryInfo dirInfo, List<DirectoryInfo> collectedFolders, int depth)
        {
            if (depth > options.Depth)
            {
                // TODO: Add verbose message here
                return false;
            }
            try
            {
                var subDirs = dirInfo.GetDirectories();
                var files = dirInfo.GetFiles();

                if (subDirs.Length == 0 && files.Length > 0)
                {
                    // TODO: Add verbose message here
                    // Console.WriteLine($"Adding {dirInfo.FullName}");
                    collectedFolders.Add(dirInfo);
                }
                else
                {
                    if (options.IncludeParentFolder && files.Length > 0)
                    {
                        // TODO: Add verbose message here
                        // Console.WriteLine($"Adding {dirInfo.FullName}");
                        collectedFolders.Add(dirInfo);
                    }

                    foreach (DirectoryInfo subDir in subDirs)
                    {
                        RecursiveCollectFolder(options, subDir, collectedFolders, depth + 1);
                    }
                }

                return true;
            }
            catch (UnauthorizedAccessException)
            {
                // TODO: Add verbose message here
                return false;
            }
        }

        private static FileInfo SortFiles(ApplyOptions applyOptions, IEnumerable<FileInfo> fileInfos)
        {
            var filteredFiles = fileInfos.Where(f => applyOptions.Filter.Contains(f.Extension)).ToList();
            switch (applyOptions.Sort)
            {
                default:
                case ApplyOptions.SortType.nameAsc:
                    return filteredFiles.OrderBy(f => f.Name).First();
                case ApplyOptions.SortType.nameDes:
                    return filteredFiles.OrderByDescending(f => f.Name).First();
                case ApplyOptions.SortType.dateAsc:
                    return filteredFiles.OrderBy(f => f.LastWriteTime).First();
                case ApplyOptions.SortType.dateDes:
                    return filteredFiles.OrderByDescending(f => f.LastWriteTime).First();
                case ApplyOptions.SortType.random:
                    var rng = new Random();
                    return filteredFiles[rng.Next(0, filteredFiles.Count - 1)];
            }
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);
    }
}