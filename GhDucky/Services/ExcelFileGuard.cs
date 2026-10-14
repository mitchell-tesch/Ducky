using System;
using System.IO;

namespace GhDucky.Services
{
    /// <summary>
    /// Detects whether an Excel workbook is currently locked (typically because
    /// it is open in Excel). Grasshopper solves run frequently and silently, so
    /// surfacing a clear "close the workbook" message is far more useful than
    /// letting a raw <see cref="IOException"/> escape.
    /// </summary>
    internal static class ExcelFileGuard
    {
        /// <summary>
        /// Returns true when the file cannot currently be opened for writing.
        /// <paramref name="reason"/> carries a user-facing explanation.
        /// </summary>
        public static bool IsLockedForWriting(string path, out string reason)
        {
            reason = null;

            if (!File.Exists(path))
                return false;

            // Excel drops a "~$Book.xlsx" owner file alongside an open workbook.
            // It is a strong hint even when the lock itself is advisory.
            var ownerFile = OwnerFilePath(path);
            var ownerExists = ownerFile != null && File.Exists(ownerFile);

            try
            {
                using var _ = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false;
            }
            catch (IOException)
            {
                reason = ownerExists
                    ? $"'{Path.GetFileName(path)}' is open in Excel. Close the workbook and try again."
                    : $"'{Path.GetFileName(path)}' is locked by another process.";
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                reason = $"'{Path.GetFileName(path)}' is read-only or access is denied.";
                return true;
            }
        }

        /// <summary>
        /// Returns true when the file cannot currently be opened for reading.
        /// Excel itself holds only a read-share lock, so reads normally succeed
        /// even while the workbook is open.
        /// </summary>
        public static bool IsLockedForReading(string path, out string reason)
        {
            reason = null;

            if (!File.Exists(path))
                return false;

            try
            {
                using var _ = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return false;
            }
            catch (IOException)
            {
                reason = $"'{Path.GetFileName(path)}' is locked by another process and cannot be read.";
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                reason = $"'{Path.GetFileName(path)}' cannot be read (access denied).";
                return true;
            }
        }

        /// <summary>True when the file has an extension this plugin can handle.</summary>
        public static bool IsSupportedExtension(string path)
        {
            var ext = Path.GetExtension(path);
            return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                   || ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Path of Excel's "~$" owner file for the given workbook.</summary>
        private static string OwnerFilePath(string path)
        {
            var dir = Path.GetDirectoryName(path);
            var name = Path.GetFileName(path);
            if (string.IsNullOrEmpty(name)) return null;
            return dir is null ? "~$" + name : Path.Combine(dir, "~$" + name);
        }
    }
}

