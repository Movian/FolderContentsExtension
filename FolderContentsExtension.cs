using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.Linq;

namespace FolderContentsExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Directory)]
    public class FolderContentsExtension : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            // Enable the menu only when there are files in the selected folder
            return GetFilesInFolder().Count > 0;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            // Create the context menu
            var menu = new ContextMenuStrip();

            // Create the sub-menu
            var subMenu = new ToolStripMenuItem("Contents");

            // Retrieve the list of files in the folder you right-click on
            List<string> files = GetFilesInFolder();

            foreach (string file in files)
            {
                try
                {
                    // Create a context menu item for each file
                    var fileMenuItem = new ToolStripMenuItem(Path.GetFileName(file));

                    // Get the file icon and set it as the menu item's image
                    IntPtr hIcon = GetFileIcon(file);
                    if (hIcon != IntPtr.Zero)
                    {
                        try
                        {
                            Icon fileIcon = Icon.FromHandle(hIcon);

                            // Check if the icon has a valid image
                            if (fileIcon?.Size.Width > 0 && fileIcon?.Size.Height > 0)
                            {
                                fileMenuItem.Image = fileIcon.ToBitmap();
                            }
                            else
                            {
                                // Fallback to a default icon
                                fileMenuItem.Image = SystemIcons.Application.ToBitmap();
                            }
                        }
                        finally
                        {
                            // Release the icon handle
                            DestroyIcon(hIcon);
                        }
                    }

                    fileMenuItem.Click += (sender, e) =>
                    {
                        // Handle the click event for the file menu item
                        // Implement the desired action when the user selects the menu item
                        try
                        {
                            Process.Start(file);
                        }
                        catch (Exception ex)
                        {
                            // Log the exception
                            LogError($"Error opening file: {file}", ex);
                        }
                    };

                    // Add the file menu item to the sub-menu
                    subMenu.DropDownItems.Add(fileMenuItem);
                }
                catch (Exception ex)
                {
                    // Log the exception
                    LogError($"Error creating menu item for file: {file}", ex);
                }
            }

            // Add the sub-menu to the main context menu
            menu.Items.Add(subMenu);

            return menu;
        }

        private List<string> GetFilesInFolder()
        {
            // Implement the logic to retrieve the list of files in the folder
            string folderPath = SelectedItemPaths.FirstOrDefault();
            return new List<string>(Directory.GetFiles(folderPath));
        }

        private void LogError(string message, Exception ex)
        {
            // Log the error message and exception to a file or console
            string logFilePath = @"C:\Temp\FolderContentsExtension_ErrorLog.txt";
            string logMessage = $"{DateTime.Now}: {message}\nException: {ex}\n\n";

            // Write the log message to a file
            File.AppendAllText(logFilePath, logMessage);

            // Alternatively, you can output the log message to the console
            Console.WriteLine(logMessage);
        }

        private const uint SHGFI_ICON = 0x100;
        private const uint SHGFI_SMALLICON = 0x1;

        [DllImport("shell32.dll")]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFileInfo psfi, uint cbSizeFileInfo, uint uFlags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool DestroyIcon(IntPtr handle);

        private static IntPtr GetFileIcon(string filePath)
        {
            SHFileInfo shinfo = new SHFileInfo();
            IntPtr hIcon = SHGetFileInfo(filePath, 0, ref shinfo, (uint)Marshal.SizeOf(shinfo), SHGFI_ICON | SHGFI_SMALLICON);
            return hIcon;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFileInfo
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }
    }
}
